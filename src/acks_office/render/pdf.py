"""主题驱动的 PDF：用 reportlab 直接排版，字体嵌入文件，在任何设备上显示一致。

版式与 Word 渲染相同，由同一套主题令牌决定：封面（title-block / standard / issue）、
章节（inline 编号标题 / opener 章节页）、栏数与首字下沉、提示块（tint / rule）、引文（bar / pull）、
表格表头（fill / label）与合计行（rules / bar）。页眉页脚来自主题外壳的模板。

排版两遍：第一遍得到总页数和各章所在的页码，第二遍写出文件（页脚的总页数、封面目录的页码）。
"""

import os
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape, unescape

from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (BaseDocTemplate, Flowable, Frame, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Table, TableStyle, XPreformatted)
from reportlab.platypus.flowables import BalancedColumns, HRFlowable, ImageAndFlowables

from .. import markdown_blocks as mdb
from ..themes import Theme, load_theme
from .common import PAGE_SIZES, attach_captions, display_width, document_meta, is_total_row, numeric_columns
from .pdffonts import PdfFontBook, is_cjk, script_runs

CM = 72 / 2.54
ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 4}
BULLETS = ("•", "◦", "▪")
_TAG = re.compile(r"<[^>]*>")


class _Para(Paragraph):
    """含中文的段落按中文规则折行；纯西文的段落按西文规则折行。

    reportlab 的中文折行处理纯西文段落时，段落被拆到下一页、下一栏或首字下沉旁边后，
    剩下部分会重新按中文规则折行，词与词之间的空格就丢了。
    """

    def __init__(self, text, style, *args, **kwargs):
        if style.wordWrap == "CJK" and not any(is_cjk(ch) for ch in unescape(_TAG.sub("", text or ""))):
            style = _latin_style(style)
        super().__init__(text, style, *args, **kwargs)

    def split(self, availWidth, availHeight):
        if not hasattr(self, "blPara"):
            self.wrap(availWidth, availHeight)
        cjk = getattr(getattr(self, "blPara", None), "kind", 0) == 1  # 按中文规则断的行
        parts = super().split(availWidth, availHeight)
        if cjk and len(parts) == 2 and getattr(parts[0], "blPara", None) is not None:
            parts[0]._lines = parts[0].blPara  # 拆分时已经断好的行
        return parts

    def wrap(self, availWidth, availHeight):
        # reportlab 拆分中文折行的段落后，前半段在同样宽度下会重新断行，遇到行尾悬挂的标点时
        # 行数可能变多，放不下就报 LayoutError。宽度不变时沿用拆分时断好的行。
        lines = getattr(self, "_lines", None)
        if lines is not None and abs(getattr(lines, "aW", -1) - availWidth) < 1e-6:
            style = self.style
            self.width = availWidth
            self._wrapWidths = [availWidth - style.leftIndent - style.firstLineIndent - style.rightIndent,
                                availWidth - style.leftIndent - style.rightIndent]
            self.blPara = lines
            self.height = len(lines.lines) * style.leading
            return self.width, self.height
        return super().wrap(availWidth, availHeight)


_LATIN: Dict[int, ParagraphStyle] = {}


def _latin_style(style: ParagraphStyle) -> ParagraphStyle:
    if id(style) not in _LATIN:
        _LATIN[id(style)] = ParagraphStyle(style.name + "-latin", parent=style, wordWrap=None)
    return _LATIN[id(style)]


class _Mark(Flowable):
    """不占位置的标记：画到某页时记下页码、更新页眉里的章节名、加 PDF 书签。"""

    def __init__(self, renderer: "PdfRenderer", chapter: Optional[str] = None, number: Optional[int] = None,
                 outline: Optional[Tuple[str, int]] = None):
        super().__init__()
        self.renderer, self.chapter, self.number, self.outline = renderer, chapter, number, outline
        self.width = self.height = 0
        self.keepWithNext = True  # 和后面的标题一起换页，书签与页码才准确

    def wrap(self, availWidth, availHeight):
        return 0, 0

    def draw(self):
        canvas, r = self.canv, self.renderer
        if self.chapter is not None:
            r.current_chapter = self.chapter
        if self.number is not None:
            r.chapter_pages[self.number] = canvas.getPageNumber()
        if self.outline:
            title, level = self.outline
            level = min(level, r.outline_level + 1)
            key = f"acks-{id(self)}"
            canvas.bookmarkPage(key)
            canvas.addOutlineEntry(title, key, level=level, closed=level > 0)
            r.outline_level = level


class _DropCap(Flowable):
    """首字下沉的大字：配合 ImageAndFlowables，让段落前几行绕排在它右侧。"""

    def __init__(self, char: str, font: str, size: float, color, height: float):
        super().__init__()
        self.char, self.font, self.size, self.color, self.cap_height = char, font, size, color, height

    def wrap(self, availWidth, availHeight):
        self.width = stringWidth(self.char, self.font, self.size)
        self.height = self.cap_height
        return self.width, self.height

    def _restrictSize(self, availWidth, availHeight):
        return self.wrap(availWidth, availHeight)

    def _unRestrictSize(self):
        pass

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.setFont(self.font, self.size)
        self.canv.drawString(0, self.cap_height - self.size * 0.86, self.char)


class PdfRenderer:
    def __init__(self, theme: Theme, meta: Dict[str, Any], base_dir: str, footer_label: Optional[str] = None):
        self.t = theme
        self.meta = meta
        self.base_dir = base_dir
        self.footer_label = footer_label
        self.f = PdfFontBook(theme)
        self.layout = theme.get("layout.doc")
        self.values = theme.chrome_values(meta)
        width, height = PAGE_SIZES[self.layout["page"]["size"]]
        self.page_w, self.page_h = width * CM, height * CM
        margin = self.layout["page"]["margin_cm"]
        self.top, self.bottom = margin["top"] * CM, margin["bottom"] * CM
        self.left, self.right = margin["left"] * CM, margin["right"] * CM
        self.text_width = self.page_w - self.left - self.right
        self.gap = float(self.layout["column_gap_cm"]) * CM
        self.columns = int(self.layout["columns"])
        self.cover = "title-block" if meta.get("cover") is False else self.layout["cover"]
        self.total_pages = 0
        self.toc_pages: Dict[int, int] = {}
        self.warnings: List[Dict[str, str]] = []
        self._reset()

    def _reset(self) -> None:
        """每一遍排版前清空计数。"""
        self.story: List[Flowable] = []
        self.pending: List[Flowable] = []
        self.table_no = self.figure_no = self.chapter_no = 0
        self.heading_no = [0, 0, 0]
        self.current_chapter = ""
        self.chapter_pages: Dict[int, int] = {}
        self.outline_level = -1
        self.in_columns = False
        self._fresh_page = True
        self._drop_cap_next = False
        self._styles: Dict[Tuple, ParagraphStyle] = {}

    # ------------------------------------------------------------ 尺寸、颜色、样式

    def sz(self, key: str) -> float:
        return float(self.t.get(f"type.doc.{key}"))

    def lead(self, key: str) -> float:
        return float(self.t.get(f"leading.{key}"))

    def space(self, key: str) -> float:
        return float(self.t.get(f"space.{key}"))

    def color(self, ref: str):
        return colors.HexColor(self.t.color(ref))

    def hex(self, ref: str) -> str:
        return self.t.color(ref)

    def style(self, role: str, size: float, color: str = "ink", leading: float = 1.3, before: float = 0,
              after: float = 0, align: str = "left", bold: Optional[bool] = None, keep: bool = False,
              left: float = 0, right: float = 0, bullet: float = 0) -> ParagraphStyle:
        key = (role, size, color, leading, before, after, align, bold, keep, left, right, bullet)
        if key not in self._styles:
            face = self.f.face(role, "cn", bold)
            self._styles[key] = ParagraphStyle(
                f"acks-{len(self._styles)}", fontName=face.name, fontSize=size, leading=size * leading,
                spaceBefore=before, spaceAfter=after, alignment=ALIGN[align], textColor=self.color(color),
                wordWrap="CJK", keepWithNext=keep, leftIndent=left, rightIndent=right, bulletIndent=bullet,
                bulletFontName=face.name, bulletFontSize=size, bulletColor=self.color(color))
        return self._styles[key]

    @property
    def flow_width(self) -> float:
        """当前内容的宽度：分栏时为一栏宽。"""
        if not self.in_columns:
            return self.text_width
        return (self.text_width - self.gap * (self.columns - 1)) / self.columns

    @property
    def body_align(self) -> str:
        return "justify" if self.layout["justify"] else "left"

    def text(self, text: str, role: str = "body", bold: Optional[bool] = None, italic: bool = False,
             color: Optional[str] = None, size: Optional[float] = None) -> str:
        return self.f.markup(str(text), role, bold, italic, self.hex(color) if color else None, size)

    def spans(self, spans: List[mdb.Span], role: str = "body", size: Optional[float] = None,
              color: Optional[str] = None, heading: bool = False) -> str:
        out = []
        for span in spans:
            if span.code:
                markup = self.text(span.text, "code", bold=False, color=None if heading else "accent",
                                   size=(size or self.sz("body")) * 0.92)
            else:
                markup = self.text(span.text, role, bold=True if span.bold else (None if heading else False),
                                   italic=span.italic, color=color)
            if span.strike:
                markup = f"<strike>{markup}</strike>"
            if span.link:
                href = escape(span.link, {'"': "&quot;"})
                markup = f'<a href="{href}" color="{self.hex("accent")}"><u>{markup}</u></a>'
            out.append(markup)
        return "".join(out)

    # ------------------------------------------------------------ 流式内容与分栏

    def add(self, *flowables: Flowable, full_width: bool = False) -> None:
        """正文内容：分栏开启时放进待排的栏里；表格、代码、图片等通栏内容先把已有的栏排出来。"""
        if self.in_columns and not full_width:
            self.pending.extend(flowables)
            return
        self.flush()
        self.story.extend(flowables)

    def flush(self) -> None:
        if self.pending:
            self.story.append(BalancedColumns(self.pending, nCols=self.columns, innerPadding=self.gap,
                                              spaceBefore=0, spaceAfter=0))
            self.pending = []

    # ------------------------------------------------------------ 页面

    def _templates(self) -> List[PageTemplate]:
        """第一页：封面（standard / issue）或带页脚的首页（title-block）；之后都是正文页。"""
        frame = dict(x1=self.left, y1=self.bottom, width=self.text_width, height=self.page_h - self.top - self.bottom,
                     leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        if self.cover == "title-block":
            first = PageTemplate("first", [Frame(id="first", **frame)], onPageEnd=self._first_page)
        else:
            first = PageTemplate("cover", [Frame(id="cover", **frame)], onPageEnd=self._cover_page)
        return [first, PageTemplate("body", [Frame(id="body", **frame)], onPageEnd=self._body_page)]

    def _first_page(self, canvas, doc) -> None:
        self._chrome(canvas, header=False)

    def _body_page(self, canvas, doc) -> None:
        self._chrome(canvas)

    def _cover_page(self, canvas, doc) -> None:
        if self.cover == "issue":
            self._cover_toc(canvas)
        else:
            self._cover_meta(canvas)

    def draw_text(self, canvas, x: float, y: float, text: str, role: str, size: float, color: str = "ink",
                  bold: Optional[bool] = None, align: str = "left") -> float:
        runs = [(self.f.face(role, script, bold).name, chunk) for script, chunk in script_runs(text)]
        widths = [stringWidth(chunk, name, size) for name, chunk in runs]
        x -= {"left": 0, "center": sum(widths) / 2, "right": sum(widths)}[align]
        canvas.setFillColor(self.color(color))
        for (name, chunk), width in zip(runs, widths, strict=True):
            canvas.setFont(name, size)
            canvas.drawString(x, y, chunk)
            x += width
        return sum(widths)

    def _chrome_text(self, path: str, page: int) -> str:
        parts = []
        for kind, value in self.t.template(path, dict(self.values, chapter=self.current_chapter)):
            if kind == "text":
                parts.append(value)
            else:
                parts.append(str(page if value == "page" else (self.total_pages or page)))
        return "".join(parts)

    def _chrome(self, canvas, header: bool = True) -> None:
        canvas.saveState()
        page = canvas.getPageNumber()
        size = self.sz("header")
        left_x, right_x = self.left, self.page_w - self.right
        if header:
            y = self.page_h - float(self.layout["page"]["header_cm"]) * CM - size
            self.draw_text(canvas, left_x, y, self._chrome_text("doc.header.left", page), "label", size, "muted", False)
            self.draw_text(canvas, right_x, y, self._chrome_text("doc.header.right", page), "label", size, "muted",
                           False, align="right")
            if self.layout["header_rule"]:
                canvas.setStrokeColor(self.color("rule"))
                canvas.setLineWidth(0.75)
                canvas.line(left_x, y - 5, right_x, y - 5)
        y = float(self.layout["page"]["footer_cm"]) * CM + size * 0.3
        left = self.footer_label if self.footer_label is not None else self._chrome_text("doc.footer.left", page)
        self.draw_text(canvas, left_x, y, left, "label", size, "muted", False)
        self.draw_text(canvas, right_x, y, self._chrome_text("doc.footer.right", page), "label", size, "muted",
                       False, align="right")
        canvas.restoreState()

    # ------------------------------------------------------------ 封面

    def _bar(self, width: float, weight: float, color: str, before: float = 0, after: float = 0) -> HRFlowable:
        return HRFlowable(width=width, thickness=weight, color=self.color(color), hAlign="LEFT",
                          spaceBefore=before, spaceAfter=after)

    def _title(self, before: float = 0, after: float = 0) -> Paragraph:
        size = self.sz("title")
        return _Para(self.text(self.meta.get("title", ""), "title"),
                         self.style("title", size, leading=self.lead("title"), before=before, after=after))

    def _two_sides(self, left: str, right: str) -> Table:
        table = Table([[_Para(left, self.style("label", 12, leading=1.3)),
                        _Para(right, self.style("label", 12, leading=1.3, align="right"))]],
                      colWidths=[self.text_width / 2] * 2)
        table.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                   ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                                   ("VALIGN", (0, 0), (-1, -1), "BOTTOM")]))
        return table

    def _title_block(self) -> None:
        """neutral：标题放在第一页顶部，不单独占一页。"""
        if self.meta.get("kicker"):
            size = self.sz("kicker")
            self.story.append(_Para(self.text(self.meta["kicker"], "label", color="accent"),
                                        self.style("label", size, "accent", 1.4, after=4)))
        self.story.append(self._title())
        if self.meta.get("subtitle"):
            size = self.sz("subtitle")
            self.story.append(_Para(self.text(self.meta["subtitle"], "body", False),
                                        self.style("body", size, "muted", 1.6, before=6)))
        info = " · ".join(str(self.meta[k]) for k in self.t.chrome_get("doc.cover_meta", []) if self.meta.get(k))
        small = self.sz("small")
        if info:
            self.story.append(_Para(self.text(info, "label", False), self.style("label", small, "muted", 1.4,
                                                                                     before=8, after=6)))
        self.story.append(HRFlowable(width="100%", thickness=0.75, color=self.color("rule_strong"),
                                     spaceBefore=0 if info else 8, spaceAfter=18))
        self._fresh_page = False

    def _cover_standard(self) -> None:
        """slate：品牌行、强调短线、类型、标题、副标题；日期 / 编制 / 版本固定在页底（见 _cover_meta）。"""
        values = self.values
        self.story.append(self._two_sides(
            self.text(values.get("brand", ""), "heading", size=self.sz("brand")),
            self.text(values.get("classification", ""), "label", False, color="muted",
                      size=self.sz("classification"))))
        self.story.append(self._bar(30, 3, "accent", before=self.space("cover_title_top"), after=10))
        if self.meta.get("kicker"):
            size = self.sz("kicker")
            self.story.append(_Para(self.text(self.meta["kicker"], "label", color="accent"),
                                        self.style("label", size, "accent", 1.3, after=10)))
        self.story.append(self._title(after=10))
        if self.meta.get("subtitle"):
            size = self.sz("subtitle")
            self.story.append(_Para(self.text(self.meta["subtitle"], "body", False),
                                        self.style("body", size, "muted", 1.6)))

    def _cover_meta(self, canvas) -> None:
        rows = [(self.t.label(k), str(self.meta[k])) for k in self.t.chrome_get("doc.cover_meta", [])
                if self.meta.get(k)]
        if not rows:
            return
        label = self.style("label", self.sz("meta_label"), "muted", 1.3, bold=False)
        value = self.style("body", self.sz("meta_value"), leading=1.3)
        data = [[_Para(self.text(k, "label", False), label), _Para(self.text(v, "body", False), value)]
                for k, v in rows]
        table = Table(data, colWidths=[54, self.text_width - 54])
        commands = [("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                    ("LINEABOVE", (0, 0), (-1, 0), 0.75, self.color("rule_strong")),
                    ("LINEBELOW", (0, 0), (-1, -1), 0.75, self.color("rule"))]
        table.setStyle(TableStyle(commands))
        table.wrapOn(canvas, self.text_width, self.page_h)
        table.drawOn(canvas, self.left, self.bottom)

    def _cover_issue(self) -> None:
        """folio：刊名与季节、超大期号、标题、导语、朱砂短线；本期目录固定在页底（见 _cover_toc）。"""
        values = self.values
        self.story.append(self._two_sides(
            self.text(values.get("publication") or values.get("brand", ""), "label", size=self.sz("brand")),
            self.text(str(self.meta.get("season", "")).upper(), "display", True, color="muted",
                      size=self.sz("classification"))))
        if self.meta.get("issue"):
            number = self.sz("issue_number")
            markup = (self.text(self.t.label("issue") + " ", "display", False, True, "muted", self.sz("issue_no"))
                      + self.text(str(self.meta["issue"]), "display", False, color="accent", size=number))
            self.story.append(_Para(markup, self.style("display", number, "accent", 0.86, before=30,
                                                           align="right", bold=False)))
        before = 30 if self.meta.get("issue") else self.space("cover_title_top")
        self.story.append(self._title(before=before))
        lede = self.meta.get("lede") or self.meta.get("subtitle")
        if lede:
            size = self.sz("subtitle")
            self.story.append(_Para(self.text(lede, "body", False),
                                        self.style("body", size, "muted", self.lead("lede"), before=18,
                                                   right=max(0, self.text_width - 330))))
        self.story.append(self._bar(180, 1.5, "accent", before=24))

    def _cover_toc(self, canvas) -> None:
        chapters = self.chapters
        if not chapters:
            return
        rows = []
        for n, chapter in enumerate(chapters, start=1):
            title = self.text(chapter["title"], "heading", size=self.sz("toc_title"))
            if chapter["desc"]:
                title += self.text("　" + chapter["desc"], "body", False, color="muted", size=self.sz("toc_desc"))
            page = str(self.toc_pages.get(n, ""))
            rows.append([_Para(self.text(chapter["number"], "display", False, color="accent"),
                                   self.style("display", self.sz("toc_number"), "accent", 1.2, bold=False)),
                         _Para(title, self.style("heading", self.sz("toc_title"), leading=1.4)),
                         _Para(self.text(page, "display", False),
                                   self.style("display", self.sz("toc_title"), leading=1.4, align="right",
                                              bold=False))])
        table = Table(rows, colWidths=[33, self.text_width - 63, 30])
        table.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                   ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                                   ("VALIGN", (0, 0), (-1, -1), "BASELINE"),
                                   ("LINEABOVE", (0, 0), (-1, -1), 0.75, self.color("rule")),
                                   ("LINEBELOW", (0, -1), (-1, -1), 0.75, self.color("rule"))]))
        _, height = table.wrapOn(canvas, self.text_width, self.page_h)
        table.drawOn(canvas, self.left, self.bottom)
        size = self.sz("toc_label")
        self.draw_text(canvas, self.left, self.bottom + height + 9, self.t.label("toc"), "label", size, "muted", False)

    # ------------------------------------------------------------ 正文

    @property
    def chapters(self) -> List[Dict[str, str]]:
        return self._chapter_list

    def render(self, blocks: List[Any]) -> List[Flowable]:
        self._reset()
        self._chapter_list = self._collect_chapters(blocks) if self.layout["chapter"] == "opener" else []
        if self.cover == "title-block":
            self.story.append(NextPageTemplate("body"))
            self._title_block()
        else:
            if self.cover == "issue":
                self._cover_issue()
            else:
                self._cover_standard()
            self.story += [NextPageTemplate("body"), PageBreak()]
            self._fresh_page = True
        items = attach_captions(blocks)
        i = 0
        while i < len(items):
            block, caption = items[i]
            following = items[i + 1][0] if i + 1 < len(items) else None
            if isinstance(block, mdb.Heading) and block.level == 1 and self.layout["chapter"] == "opener":
                lede = following if isinstance(following, mdb.Quote) and not following.alert else None
                self._chapter_opener(block, lede)
                i += 2 if lede else 1
                continue
            self._block(block, 0, caption)
            i += 1
        self.flush()
        return self.story

    def _collect_chapters(self, blocks: List[Any]) -> List[Dict[str, str]]:
        chapters = []
        for block in blocks:
            if isinstance(block, mdb.Heading) and block.level == 1:
                text = mdb.plain(block.spans).replace("\n", " ")
                label = block.attrs.get("label", "")
                title, desc = (label.split(" · ")[0], text) if label else (text, "")
                chapters.append({"number": f"{len(chapters) + 1:02d}", "title": title, "desc": desc,
                                 "label": label, "text": text})
        return chapters

    def _chapter_opener(self, heading: mdb.Heading, lede: Optional[mdb.Quote]) -> None:
        """folio 的章节页：另起一页，大号编号、标签、标题、导语、细线，之后正文分栏、首字下沉。"""
        self.flush()
        self.in_columns = False
        self.chapter_no += 1
        info = self.chapters[self.chapter_no - 1]
        label = info["label"]
        if not self._fresh_page:
            self.story.append(PageBreak())
        self._fresh_page = False
        chapter_name = label.split(" · ")[0] if label else info["desc"] or info["title"]
        self.story.append(_Mark(self, chapter=chapter_name, number=self.chapter_no, outline=(info["text"], 0)))
        number = self.sz("chapter_number")
        markup = self.text(info["number"], "display", False, color="accent", size=number)
        if label:
            head, _, tail = label.partition(" · ")
            markup += self.text("　" + head, "label", color="ink", size=self.sz("chapter_label"))
            if tail:
                markup += self.text(" · " + tail, "display", True, color="muted", size=self.sz("chapter_label") * 0.92)
        opener = [_Para(markup, self.style("display", number, "accent", 0.9, bold=False, keep=True))]
        opener.append(_Para(self.spans(heading.spans, "heading", self.sz("h1"), heading=True),
                                self.style("heading", self.sz("h1"), leading=self.lead("heading"), before=14,
                                           after=self.space("h1_after"), keep=True)))
        if lede is not None:
            text = "\n".join(mdb.plain(b.spans) for b in lede.blocks if isinstance(b, (mdb.Paragraph, mdb.Heading)))
            opener.append(_Para(self.text(text, "body", False),
                                    self.style("body", self.sz("lede"), "muted", self.lead("lede"), before=9,
                                               right=max(0, self.text_width - 390))))
        opener.append(HRFlowable(width="100%", thickness=0.75, color=self.color("rule"), spaceBefore=16,
                                 spaceAfter=16))
        self.story.append(KeepTogether(opener))
        self.in_columns = self.columns > 1
        self._drop_cap_next = bool(self.layout["drop_cap"])

    def _block(self, block: Any, level: int, caption: Optional[str] = None) -> None:
        if isinstance(block, mdb.Heading):
            self._heading(block)
        elif isinstance(block, mdb.Paragraph):
            self._paragraph(block.spans)
        elif isinstance(block, mdb.ListBlock):
            self._list(block, level)
        elif isinstance(block, mdb.Table):
            if block.header or block.rows:
                self._table(block, caption)
        elif isinstance(block, mdb.Quote):
            self._quote(block)
        elif isinstance(block, mdb.Code):
            self._code(block.text)
        elif isinstance(block, mdb.Rule):
            self.add(HRFlowable(width="100%", thickness=0.75, color=self.color("rule"), spaceBefore=6, spaceAfter=6))
        elif isinstance(block, mdb.Image):
            self._image(block)
        self._fresh_page = False

    def _heading(self, block: mdb.Heading) -> None:
        level = min(block.level, 3)
        size = self.sz(f"h{level}")
        markup = self.spans(block.spans, "heading", size, heading=True)
        if self.layout["heading_numbers"] and self.layout["chapter"] == "inline":
            self.heading_no[level - 1] += 1
            for deeper in range(level, 3):
                self.heading_no[deeper] = 0
            number = ".".join(str(n) for n in self.heading_no[:level])
            gap = f'<font name="{self.f.face("number", "en").name}">&nbsp;&nbsp;</font>'
            markup = self.text(number, "number", color="accent") + gap + markup
        style = self.style("heading", size, leading=self.lead("heading"), before=self.space(f"h{level}_before"),
                           after=self.space(f"h{level}_after"), keep=True)
        self.add(_Mark(self, outline=(mdb.plain(block.spans).replace("\n", " "), level - 1)), _Para(markup, style))

    def _paragraph(self, spans: List[mdb.Span]) -> None:
        style = self.style("body", self.sz("body"), leading=self.lead("body"), after=self.space("paragraph"),
                           align=self.body_align)
        if self._drop_cap_next and spans and spans[0].text and not spans[0].code:
            self._drop_cap_next = False
            first, rest = spans[0].text[0], spans[0].text[1:]
            spans = ([mdb.Span(rest, spans[0].bold, spans[0].italic, False, spans[0].strike, spans[0].link)]
                     if rest else []) + spans[1:]
            line = self.sz("body") * self.lead("body")
            script = "cn" if ord(first) > 0x2E7F else "en"
            # 绕排的行数按区域高度向上取整：2.5 行高的区域让段落前三行排在大字右侧（与样张一致）
            cap = _DropCap(first, self.f.face("title", script).name, self.sz("drop_cap"), self.color("accent"),
                           line * 2.5)
            self.add(ImageAndFlowables(cap, [_Para(self.spans(spans), style)], imageSide="left",
                                       imageRightPadding=6, imageBottomPadding=0, imageTopPadding=0))
            return
        self.add(_Para(self.spans(spans), style))

    def _list(self, block: mdb.ListBlock, level: int) -> None:
        size = self.sz("body")
        indent = 18 * (level + 1)
        for n, item in enumerate(block.items, start=block.start):
            if item.checked is not None:
                bullet, color = ("■" if item.checked else "□"), "accent"
            elif block.ordered:
                bullet, color = _ordinal(n, level) + ".", "ink"
            else:
                bullet, color = BULLETS[min(level, 2)], "ink"
            style = self.style("body", size, color, self.lead("body"), after=3, align=self.body_align,
                               left=indent, bullet=indent - 14)
            markup = self.spans(item.spans, color="ink" if color != "ink" else None)
            self.add(_Para(markup, style, bulletText=bullet))
            for child in item.children:
                self._block(child, level + 1)

    def _caption(self, kind: str, text: str) -> Paragraph:
        if kind == "table":
            self.table_no += 1
            number = f"{self.t.label('table')} {self.table_no}"
        else:
            self.figure_no += 1
            number = f"{self.t.label('figure')} {self.figure_no}"
        size = self.sz("caption")
        if self.layout.get("caption") == "accent":
            markup = self.text(number, "display", False, True, "accent", size * 1.13)
        else:
            markup = self.text(number, "label", color="muted")
        markup += self.text("　" + text, "label", False, color="muted")
        return _Para(markup, self.style("label", size, "muted", 1.4, before=4, after=6, bold=False,
                                            keep=kind == "table"))

    def _table(self, block: mdb.Table, caption: Optional[str]) -> None:
        header = [mdb.plain(c).replace("\n", " ") for c in block.header]
        rows_text = [[mdb.plain(c).replace("\n", " ") for c in r] for r in block.rows]
        width = max([len(header)] + [len(r) for r in rows_text])
        numeric = set(numeric_columns(header, rows_text, block.aligns))
        weights = []
        for i in range(width):
            texts = ([header[i]] if i < len(header) else []) + [r[i] for r in rows_text if i < len(r)]
            weights.append(min(max([display_width(t) for t in texts] + [4]), 40))
        widths = [self.text_width * w / sum(weights) for w in weights]
        tokens = self.layout["table"]
        head_variant = tokens["head"]
        size, head_size = self.sz("table"), self.sz("table_head")
        leading = self.lead("table")
        data, commands = [], []
        has_header = bool(header)
        all_rows = ([("head", block.header)] if has_header else []) + [("row", r) for r in block.rows]
        last = len(all_rows) - 1
        for r, (kind, cells) in enumerate(all_rows):
            texts = header if kind == "head" else rows_text[r - (1 if has_header else 0)]
            total_row = kind == "row" and is_total_row(texts)
            row = []
            for c in range(width):
                align = "right" if c in numeric else "left"
                spans = cells[c] if c < len(cells) else []
                if kind == "head":
                    markup = self.text(mdb.plain(spans), "label", False, color="muted") if spans else ""
                    style = self.style("label", head_size, "muted", leading, align=align, bold=False)
                else:
                    markup = self.spans([mdb.Span(s.text, s.bold or total_row, s.italic, s.code, s.strike, s.link)
                                         for s in spans])
                    style = self.style("body", size, leading=leading, align=align)
                row.append(_Para(markup, style))
            data.append(row)
            if kind == "head":
                if head_variant == "fill":
                    commands += [("BACKGROUND", (0, r), (-1, r), self.color("surface")),
                                 ("LINEABOVE", (0, r), (-1, r), 0.75, self.color(tokens["head_rule"]))]
                else:
                    commands.append(("LINEBELOW", (0, r), (-1, r), 1.5, self.color(tokens["head_rule"])))
            elif total_row:
                bar = tokens["total"] == "bar"
                commands.append(("LINEABOVE", (0, r), (-1, r), 1.5 if bar else 0.75, self.color(tokens["total_rule"])))
                if not bar:
                    commands.append(("LINEBELOW", (0, r), (-1, r), 0.75, self.color(tokens["total_rule"])))
            elif r == last:
                closing = tokens["row_rule" if tokens["total"] == "bar" else "total_rule"]
                commands.append(("LINEBELOW", (0, r), (-1, r), 0.75, self.color(closing)))
            elif not (r + 1 <= last and is_total_row(rows_text[r + 1 - (1 if has_header else 0)])):
                commands.append(("LINEBELOW", (0, r), (-1, r), 0.75, self.color(tokens["row_rule"])))
        cell_v, cell_h = self.space("table_cell_v"), self.space("table_cell_h")
        commands += [("TOPPADDING", (0, 0), (-1, -1), cell_v), ("BOTTOMPADDING", (0, 0), (-1, -1), cell_v),
                     ("LEFTPADDING", (0, 0), (-1, -1), cell_h), ("RIGHTPADDING", (0, 0), (-1, -1), cell_h),
                     ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
        table = Table(data, colWidths=widths, repeatRows=1 if has_header else 0, hAlign="LEFT",
                      spaceAfter=self.space("block"))
        table.setStyle(TableStyle(commands))
        flowables = ([self._caption("table", caption)] if caption else []) + [table]
        self.add(*flowables, full_width=True)

    def _box(self, content: List[Flowable], fill: Optional[str], rules: bool, pad_v: float, pad_h: float) -> Table:
        table = Table([[content]], colWidths=[self.flow_width], hAlign="LEFT", spaceAfter=self.space("block"))
        commands = [("TOPPADDING", (0, 0), (-1, -1), pad_v), ("BOTTOMPADDING", (0, 0), (-1, -1), pad_v),
                    ("LEFTPADDING", (0, 0), (-1, -1), pad_h), ("RIGHTPADDING", (0, 0), (-1, -1), pad_h)]
        if fill:
            commands.append(("BACKGROUND", (0, 0), (-1, -1), self.color(fill)))
        if rules:
            commands += [("LINEABOVE", (0, 0), (-1, 0), 0.75, self.color("rule_strong")),
                         ("LINEBELOW", (0, -1), (-1, -1), 0.75, self.color("rule_strong"))]
        table.setStyle(TableStyle(commands))
        return table

    def _quote(self, block: mdb.Quote) -> None:
        paragraphs = [b for b in block.blocks if isinstance(b, (mdb.Paragraph, mdb.Heading, mdb.ListBlock))]
        if block.alert:
            self._callout(block.alert, paragraphs)
        elif self.layout["quote"] == "pull":
            self._pull_quote(paragraphs)
        else:
            size = self.sz("quote")
            style = self.style("quote", size, "muted", self.lead("quote"), after=4)
            content = [_Para(self.spans(spans, "quote", size, "muted"), style)
                       for b in paragraphs for spans in _spans_of(b)]
            if not content:
                return
            table = Table([[content]], colWidths=[self.flow_width], hAlign="LEFT", spaceAfter=self.space("paragraph"))
            table.setStyle(TableStyle([("LINEBEFORE", (0, 0), (0, -1), 2.25, self.color("rule_mid")),
                                       ("LEFTPADDING", (0, 0), (-1, -1), 12), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                       ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
            self.add(table)

    def _callout(self, alert: str, blocks: List[Any]) -> None:
        tint = self.layout["callout"] == "tint"
        size = self.sz("callout_label")
        content = [_Para(self.text(self.t.label(alert), "label", True, color="accent"),
                             self.style("label", size, "accent", 1.3, after=3, bold=True))]
        text_style = self.style("body", self.sz("callout"), leading=self.lead("callout"), align=self.body_align)
        content += [_Para(self.spans(spans, "body", self.sz("callout")), text_style)
                    for b in blocks for spans in _spans_of(b)]
        pad_v, pad_h = self.space("callout_pad_v"), self.space("callout_pad_h")
        self.add(self._box(content, "surface" if tint else None, rules=not tint, pad_v=pad_v,
                           pad_h=pad_h if tint else 0))

    def _pull_quote(self, blocks: List[Any]) -> None:
        size = self.sz("pull_quote")
        mark_width = size * 1.7
        mark = _Para(self.text("“", "display", False, color="accent", size=size * 1.8),
                         self.style("display", size * 1.8, "accent", 0.9, bold=False))
        style = self.style("quote", size, leading=self.lead("pull_quote"))
        text = [_Para(self.spans(spans, "quote", size), style) for b in blocks for spans in _spans_of(b)]
        if not text:
            return
        table = Table([[mark, text]], colWidths=[mark_width, self.flow_width - mark_width], hAlign="LEFT",
                      spaceAfter=self.space("block"))
        table.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                   ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("LINEABOVE", (0, 0), (-1, 0), 0.75, self.color("rule_strong")),
                                   ("LINEBELOW", (0, -1), (-1, -1), 0.75, self.color("rule_strong"))]))
        self.add(table)

    def _code(self, text: str) -> None:
        size = self.sz("code")
        style = self.style("code", size, leading=1.45, bold=False)
        limit = max(20, int((self.text_width - 24) / (size * 0.6)))  # 等宽字体约 0.6 em；中文算两格
        lines = []
        for line in text.split("\n"):
            lines += _wrap_code(line, limit) or [""]
        markup = "\n".join(self.f.markup(line, "code", False, breaks=False) for line in lines)
        self.in_columns, columns = False, self.in_columns  # 代码块通栏
        box = self._box([XPreformatted(markup, style)], "surface", rules=False, pad_v=8, pad_h=10)
        self.in_columns = columns
        self.add(box, full_width=True)

    def _image(self, block: mdb.Image) -> None:
        path = block.src if os.path.isabs(block.src) else os.path.join(self.base_dir, block.src)
        if not os.path.isfile(path):
            self.warnings.append({"code": "IMAGE_NOT_FOUND", "message": f"图片不存在：{block.src}"})
            self._paragraph([mdb.Span(f"[图片：{block.alt or block.src}]")])
            return
        try:
            image = Image(path)
        except Exception as exc:
            self.warnings.append({"code": "IMAGE_UNREADABLE", "message": f"图片无法读取：{block.src}（{exc}）"})
            return
        scale = min(1.0, self.text_width / image.imageWidth)
        image.drawWidth, image.drawHeight = image.imageWidth * scale, image.imageHeight * scale
        image.hAlign = "LEFT"
        flowables: List[Flowable] = [image]
        if block.alt:
            flowables.append(self._caption("figure", block.alt))
        self.add(KeepTogether(flowables), full_width=True)

    # ------------------------------------------------------------ 输出

    def build(self, blocks: List[Any], output_path: str) -> int:
        """排两遍：第一遍拿到总页数和各章页码，第二遍写文件。返回页数。"""
        info = dict(title=str(self.meta.get("title", "")).replace("\n", " "), author=str(self.meta.get("author", "")),
                    subject=str(self.meta.get("subtitle", "")), creator="acks-office")
        for target in (BytesIO(), output_path):
            doc = BaseDocTemplate(target, pagesize=(self.page_w, self.page_h), leftMargin=self.left,
                                  rightMargin=self.right, topMargin=self.top, bottomMargin=self.bottom, **info)
            doc.addPageTemplates(self._templates())
            doc.build(self.render(blocks))
            self.total_pages = doc.page
            self.toc_pages = dict(self.chapter_pages)
        return self.total_pages


def _ordinal(n: int, level: int) -> str:
    if level % 3 == 1:
        return _letters(n)
    if level % 3 == 2:
        return _roman(n)
    return str(n)


def _letters(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(97 + rem) + out
    return out or "a"


def _roman(n: int) -> str:
    table = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
             (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"))
    out = ""
    for value, letters in table:
        while n >= value:
            out += letters
            n -= value
    return out or "i"


def _wrap_code(line: str, limit: int) -> List[str]:
    """代码按显示宽度折行（中文算两格），保留行首空格。"""
    out, current, used = [], "", 0
    for ch in line:
        w = 2 if ord(ch) > 0x2E7F else 1
        if used + w > limit and current:
            out.append(current)
            current, used = "", 0
        current += ch
        used += w
    if current or not out:
        out.append(current)
    return out


def _spans_of(block: Any) -> List[List[mdb.Span]]:
    if isinstance(block, (mdb.Paragraph, mdb.Heading)):
        return [block.spans]
    if isinstance(block, mdb.ListBlock):
        return [[mdb.Span("· ")] + item.spans for item in block.items]
    return []


def render_pdf(title: Optional[str], content: str, output_path: str, theme: Any = None,
               base_dir: Optional[str] = None, footer_label: Optional[str] = None,
               font: Optional[str] = None, bold_font: Optional[str] = None, **meta: Any) -> Dict[str, Any]:
    """按主题生成 PDF。meta 可传 subtitle、kicker、brand、author、date、version、issue、lede 等，
    会覆盖正文开头 front matter 里的同名字段。font / bold_font 为代替主题中文字体的 TrueType 文件
    （也可用环境变量 ACKS_OFFICE_PDF_FONT / ACKS_OFFICE_PDF_FONT_BOLD）。"""
    theme = load_theme(theme)
    meta = {k: v for k, v in meta.items() if k not in ("theme", "font_policy")}
    info, body = document_meta(title, content, meta)
    renderer = PdfRenderer(theme, info, base_dir or os.getcwd(), footer_label)
    regular = font or os.environ.get("ACKS_OFFICE_PDF_FONT")
    if regular:
        renderer.f.override_cn(regular, bold_font or os.environ.get("ACKS_OFFICE_PDF_FONT_BOLD"))
    pages = renderer.build(mdb.parse(body), output_path)
    result: Dict[str, Any] = {"output_path": output_path, "file_size": os.path.getsize(output_path),
                              "pages": pages, "theme": theme.name, "fonts": renderer.f.embedded_fonts()}
    warnings = renderer.f.warnings + _dedupe(renderer.warnings)
    if warnings:
        result["warnings"] = warnings
    return result


def _dedupe(items: List[Dict[str, str]]) -> List[Dict[str, str]]:
    seen, out = set(), []
    for item in items:
        key = (item.get("code"), item.get("message"))
        if key not in seen:
            seen.add(key)
            out.append(item)
    return out
