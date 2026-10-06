import io
import os
import re
from copy import deepcopy
from typing import Optional, List, Dict, Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table as DocxTable
from docx.text.run import Run

from . import markdown_blocks as mdb

DEFAULT_BRAND = "ACKS Studio"
_NUMERIC = re.compile(r"^[\s(（]*[+\-−]?[¥$€£]?[\d,]+(\.\d+)?\s*%?[)）]*$")
_ALERT_LABELS = {"note": "Note · 注释", "tip": "Tip · 提示", "important": "Important · 重要",
                 "warning": "Warning · 警告", "caution": "Caution · 注意"}


def _is_cjk(text: str) -> bool:
    """Return True if text contains CJK characters."""
    return any('一' <= ch <= '鿿' for ch in text)


LEGACY_THEMES = ("acks",)  # 2.x 的 ACKS 样式；"default" 为 neutral 的别名


def create_word(title: Optional[str], content: str, output_path: str,
                theme: Any = None,
                brand_name: Optional[str] = None,
                footer_label: Optional[str] = None,
                **kwargs) -> Dict[str, Any]:
    """
    创建 Word 文档。

    Args:
        title: 文档标题；为空时用正文 front matter 里的 title
        content: 文档内容，Markdown：标题、粗体/斜体、链接、列表与任务清单、表格、引文与提示块、代码块、图片；
            开头可以写 front matter（kicker、subtitle、author、date、version、issue、lede 等）
        output_path: 输出文件路径
        theme: 主题名称（neutral、slate、folio 或已安装的主题）、主题目录或 Theme 对象；
            "acks"、"default" 是 2.x 的内置样式
        brand_name: 品牌名，传 "" 去品牌化
        footer_label: 页脚左侧文字，默认按主题模板生成
        subtitle 等元数据: 见 front matter
        base_dir: 图片相对路径的基准目录，默认当前目录
        font_policy: "local"（缺主题字体时改用本机字体，默认）或 "theme"（总是写主题字体名）
    """
    if isinstance(theme, str) and theme in LEGACY_THEMES:
        blocks = mdb.parse(content)
        base_dir = kwargs.get("base_dir") or os.getcwd()
        brand = DEFAULT_BRAND if brand_name is None else brand_name
        return _create_word_acks(title or "", blocks, output_path, brand, footer_label, base_dir,
                                 kwargs.get("subtitle", ""))
    from .render.common import META_KEYS
    from .render.word import render_word
    meta = {key: kwargs[key] for key in kwargs if key in META_KEYS or key not in _RENDER_OPTIONS}
    if brand_name is not None:
        meta["brand"] = brand_name
    return render_word(title, content, output_path, theme=theme, base_dir=kwargs.get("base_dir"),
                       font_policy=kwargs.get("font_policy", "local"), footer_label=footer_label, **meta)


_RENDER_OPTIONS = ("base_dir", "font_policy", "theme", "create_chart", "font", "bold_font")


def _footer_label(title: str, brand_name: str, footer_label: Optional[str]) -> str:
    if footer_label is not None:
        return footer_label
    if brand_name == DEFAULT_BRAND:
        return f"{brand_name} · 文档设计规范 v2"
    return brand_name or title


def _result(doc, output_path: str, theme: str, warnings: List[Dict]) -> Dict[str, Any]:
    result = {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "theme": theme,
    }
    if warnings:
        result["warnings"] = warnings
    return result


def _create_word_acks(title: str, blocks, output_path: str, brand_name: str,
                      footer_label: Optional[str], base_dir: str, subtitle: str = "") -> Dict[str, Any]:
    """用 ACKS 设计规范（design_system.tokens）生成 Word 文档。"""
    from .design_system import tokens as acks

    doc = Document()
    acks.apply_default_styles(doc)
    acks.set_a4_page_margins(doc)
    acks.add_page_footer(doc, label=_footer_label(title, brand_name, footer_label))
    _add_cover_acks(doc, title, brand_name, acks, subtitle)
    warnings = _AcksRenderer(doc, base_dir, acks).render(blocks)
    doc.save(output_path)
    return _result(doc, output_path, "acks", warnings)


def _add_cover_acks(doc, title: str, brand_name: str, acks, subtitle: str = "") -> None:
    """根据标题语言映射到 ACKS 封面组件；标题只出现一次，副标题放在第二行。"""
    if _is_cjk(title):
        acks.add_cover_page(
            doc,
            title_top="", title_em=title, zh_title=subtitle,
            doctype="", brand_name=brand_name,
        )
    else:
        acks.add_cover_page(
            doc,
            title_top=title, title_em="", zh_title=subtitle,
            doctype="", brand_name=brand_name,
        )


def _add_hyperlink(par, url: str) -> Run:
    r_id = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    link.append(r)
    par._p.append(link)
    return Run(r, par)


def _add_spans(par, spans: List[mdb.Span], style_run) -> None:
    """style_run(run, span) 设字体与颜色；这里处理粗体、斜体、删除线和链接。"""
    for span in spans:
        run = _add_hyperlink(par, span.link) if span.link else par.add_run()
        run.text = span.text  # "\n" 会转成换行
        style_run(run, span)
        if span.bold:
            run.bold = True
        if span.italic:
            run.italic = True
        if span.strike:
            run.font.strike = True


def _set_shading(par, hex_color: str) -> None:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    par._p.get_or_add_pPr().append(shd)


def _set_bottom_border(par, color: str) -> None:
    pPr = par._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", color)):
        bottom.set(qn(key), value)
    borders.append(bottom)
    pPr.append(borders)


def _table_text(block: mdb.Table):
    width = max([len(block.header)] + [len(r) for r in block.rows])
    headers = [mdb.plain(c).replace("\n", " ") for c in block.header]
    headers += [""] * (width - len(headers))
    rows = [[mdb.plain(c).replace("\n", " ") for c in r] + [""] * (width - len(r)) for r in block.rows]
    numeric = []
    for i in range(width):
        values = [r[i].strip() for r in rows if r[i].strip()]
        right = i < len(block.aligns) and block.aligns[i] == "right"
        if right or (values and all(_NUMERIC.match(v) for v in values)):
            numeric.append(i)
    return headers, rows, numeric


def _quote_text(block: mdb.Quote) -> str:
    lines = []
    for inner in block.blocks:
        if isinstance(inner, (mdb.Paragraph, mdb.Heading)):
            lines.append(mdb.plain(inner.spans))
        elif isinstance(inner, mdb.ListBlock):
            lines.extend("· " + mdb.plain(item.spans) for item in inner.items)
        elif isinstance(inner, mdb.Code):
            lines.append(inner.text)
    return "\n".join(lines)


class _Renderer:
    def __init__(self, doc, base_dir: str):
        self.doc = doc
        self.base_dir = base_dir
        self.warnings: List[Dict] = []

    def render(self, blocks) -> List[Dict]:
        for block in blocks:
            self.block(block, 0)
        return self.warnings

    def block(self, block, level: int) -> None:
        if isinstance(block, mdb.Heading):
            self.heading(block.level, mdb.plain(block.spans).replace("\n", " "))
        elif isinstance(block, mdb.Paragraph):
            self.paragraph(block.spans)
        elif isinstance(block, mdb.ListBlock):
            self.list_start(block, level)
            for n, item in enumerate(block.items, start=block.start):
                self.list_item(item, block.ordered, n, level)
                for child in item.children:
                    self.block(child, level + 1)
        elif isinstance(block, mdb.Table):
            if block.header or block.rows:
                self.table(block)
        elif isinstance(block, mdb.Quote):
            self.quote(_quote_text(block), block.alert)
        elif isinstance(block, mdb.Code):
            self.code(block.text)
        elif isinstance(block, mdb.Rule):
            self.rule()
        elif isinstance(block, mdb.Image):
            self.image(block)

    def list_start(self, block: mdb.ListBlock, level: int) -> None:
        pass

    def image(self, block: mdb.Image) -> None:
        path = block.src if os.path.isabs(block.src) else os.path.join(self.base_dir, block.src)
        if not os.path.isfile(path):
            self.warnings.append({"code": "IMAGE_NOT_FOUND", "message": f"图片不存在：{block.src}"})
            self.paragraph([mdb.Span(f"[图片：{block.alt or block.src}]")])
            return
        section = self.doc.sections[-1]
        max_width = section.page_width - section.left_margin - section.right_margin
        shape = self.doc.add_picture(path)  # 原尺寸；比版心宽才按比例缩小
        if shape.width > max_width:
            shape.height = int(shape.height * max_width / shape.width)
            shape.width = max_width
        if block.alt:
            self.caption(block.alt)


class _AcksRenderer(_Renderer):
    """ACKS 主题：复用 design_system.tokens 的组件，行内格式沿用正文字体设置。"""

    def __init__(self, doc, base_dir: str, acks):
        super().__init__(doc, base_dir)
        self.a = acks

    def heading(self, level: int, text: str) -> None:
        if level == 1:
            self.a.add_section_heading(self.doc, en=text)
        elif level == 2:
            self.a.add_sub_heading(self.doc, en=text)
        else:
            self.a.add_label(self.doc, text)

    def _style_run(self, run, span: mdb.Span) -> None:
        a = self.a
        if span.code:
            a._set_run_font(run, mono=True, size=Pt(9.5), color=a.RGB_PRIMARY_DEEP)
        else:
            a._set_run_font(run, font=a.FONT_BODY_EN, font_cn=a.FONT_BODY_CN,
                            size=a.FONT_BODY_SIZE, color=a.RGB_INK)
        if span.link:
            run.font.color.rgb = a.RGB_PRIMARY_DEEP
            run.underline = True

    def paragraph(self, spans) -> None:
        a = self.a
        par = self.doc.add_paragraph()
        _add_spans(par, spans, self._style_run)
        a._set_paragraph_spacing(par, before=a.PARAGRAPH_SPACE_BEFORE, after=a.PARAGRAPH_SPACE_AFTER,
                                 line_height=a.LINE_HEIGHT_CN, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    def list_item(self, item: mdb.ListItem, ordered: bool, n: int, level: int) -> None:
        a = self.a
        par = self.doc.add_paragraph()
        if item.checked is not None:
            marker = par.add_run(("☑" if item.checked else a.GLYPH_CHECK_EMPTY) + "  ")
            a._set_run_font(marker, size=Pt(11), color=a.RGB_PRIMARY)
            indent = Pt(18)
        elif ordered:
            marker = par.add_run(f"{n:02d}  ")
            a._set_run_font(marker, mono=True, size=Pt(10), color=a.RGB_PRIMARY_DEEP, bold=True)
            indent = Pt(20)
        else:
            marker = par.add_run(a.GLYPH_BULLET + " ")
            a._set_run_font(marker, mono=True, size=Pt(10.5), color=a.RGB_PRIMARY, bold=True)
            indent = Pt(14)
        _add_spans(par, item.spans, self._style_run)
        a._set_paragraph_spacing(par, before=Pt(0), after=Pt(4), line_height=a.LINE_HEIGHT_CN)
        par.paragraph_format.left_indent = indent + Pt(14 * level)
        par.paragraph_format.first_line_indent = -indent

    def table(self, block: mdb.Table) -> None:
        headers, rows, numeric = _table_text(block)
        self.a.add_data_table(self.doc, headers, rows, numeric_cols=numeric)
        self.doc.add_paragraph()

    def quote(self, text: str, alert: Optional[str]) -> None:
        if alert in ("note", "tip"):
            self.a.add_callout_note(self.doc, text, _ALERT_LABELS[alert])
        elif alert:
            self.a.add_callout_key(self.doc, text, _ALERT_LABELS[alert])
        else:
            self.a.add_quote(self.doc, text)

    def code(self, text: str) -> None:
        self.a.add_code_block(self.doc, text)

    def rule(self) -> None:
        _set_bottom_border(self.doc.add_paragraph(), self.a.COLOR_RULE)

    def caption(self, text: str) -> None:
        self.a.add_caption(self.doc, "Fig", text)


class _DefaultRenderer(_Renderer):
    """简单样式：使用 Word 内置样式。"""

    _HEADING_COLORS = {1: RGBColor(0, 102, 204), 2: RGBColor(51, 153, 255)}

    def __init__(self, doc, base_dir: str):
        super().__init__(doc, base_dir)
        self._num_ids: Dict[int, int] = {}

    def heading(self, level: int, text: str) -> None:
        level = min(level, 3)
        para = self.doc.add_heading(text, level=level)
        if para.runs:
            para.runs[0].font.bold = True
            if level in self._HEADING_COLORS:
                para.runs[0].font.color.rgb = self._HEADING_COLORS[level]

    def _style_run(self, run, span: mdb.Span) -> None:
        if span.code:
            run.font.name = "Consolas"
        if span.link:
            run.font.color.rgb = RGBColor(0, 102, 204)
            run.underline = True

    def paragraph(self, spans) -> None:
        _add_spans(self.doc.add_paragraph(), spans, self._style_run)

    @staticmethod
    def _list_style(base: str, level: int) -> str:
        return f"{base} {min(level + 1, 3)}" if level else base

    def list_start(self, block: mdb.ListBlock, level: int) -> None:
        """每个有序列表从自己的起始号重新编号；内置样式默认会接着上一个列表往下编。"""
        if not block.ordered:
            return
        style = self.doc.styles[self._list_style("List Number", level)]
        numbering = self.doc.part.numbering_part.element
        abstract_id = numbering.num_having_numId(style.element.pPr.numPr.numId.val).abstractNumId.val
        num = numbering.add_num(abstract_id)
        num.add_lvlOverride(ilvl=0).add_startOverride(block.start)
        self._num_ids[level] = num.numId

    def list_item(self, item: mdb.ListItem, ordered: bool, n: int, level: int) -> None:
        if item.checked is not None:
            # 任务项只用复选框做标记，不再叠加项目符号
            par = self.doc.add_paragraph(style="List Paragraph")
            par.paragraph_format.left_indent = Pt(18 * (level + 1))
            par.paragraph_format.first_line_indent = Pt(-18)
            par.add_run("☑ " if item.checked else "☐ ")
        else:
            par = self.doc.add_paragraph(style=self._list_style("List Number" if ordered else "List Bullet", level))
            if ordered:
                num_pr = par._p.get_or_add_pPr().get_or_add_numPr()
                num_pr.get_or_add_ilvl().val = 0
                num_pr.get_or_add_numId().val = self._num_ids[level]
        _add_spans(par, item.spans, self._style_run)

    def table(self, block: mdb.Table) -> None:
        headers, rows, numeric = _table_text(block)
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = "Table Grid"
        for i, text in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.paragraphs[0].add_run(text).bold = True
        for r, row in enumerate(rows, start=1):
            for i, text in enumerate(row):
                par = table.rows[r].cells[i].paragraphs[0]
                par.add_run(text)
                if i in numeric:
                    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self.doc.add_paragraph()

    def quote(self, text: str, alert: Optional[str]) -> None:
        if alert:
            text = f"{_ALERT_LABELS[alert]}\n{text}"
        self.doc.add_paragraph(text, style="Intense Quote" if alert else "Quote")

    def code(self, text: str) -> None:
        par = self.doc.add_paragraph()
        run = par.add_run(text)
        run.font.name = "Consolas"
        run.font.size = Pt(10)
        _set_shading(par, "F2F2F2")

    def rule(self) -> None:
        _set_bottom_border(self.doc.add_paragraph(), "BFBFBF")

    def caption(self, text: str) -> None:
        self.doc.add_paragraph(text, style="Caption")


# ---------------------------------------------------------------- 水印、提取、合并

def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None,
                  font_size: int = 48, color: tuple = (200, 200, 200),
                  italic: bool = True, **kwargs) -> Dict[str, Any]:
    """
    给 Word 文档添加水印（页眉斜体灰色大字，简化版）。

    Args:
        input_path: 输入文件路径
        watermark_text: 水印文本
        output_path: 输出文件路径，默认为 <原名>_watermarked.docx，不覆盖原文件
        font_size: 字号
        color: RGB 颜色元组
        italic: 是否斜体
    """
    if not output_path:
        from .api import watermarked_path
        output_path = watermarked_path(input_path)

    doc = Document(input_path)

    section = doc.sections[0]
    header = section.header

    watermark_para = header.add_paragraph()
    watermark_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = watermark_para.add_run(watermark_text)
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(*color)
    run.bold = True
    run.italic = italic

    doc.save(output_path)

    return {"output_path": output_path}


def extract_text(input_path: str, **kwargs) -> str:
    """
    提取 Word 文档中的文本：段落与表格按原文顺序输出，表格每行一行、单元格用 " | " 分隔
    """
    doc = Document(input_path)
    lines: List[str] = []
    for item in doc.iter_inner_content():
        if isinstance(item, DocxTable):
            lines.extend(_table_lines(item))
        else:
            lines.append(item.text)
    return '\n'.join(lines)


def _table_lines(table: DocxTable) -> List[str]:
    lines = []
    for row in table.rows:
        seen, cells = [], []
        for cell in row.cells:
            if any(cell._tc is tc for tc in seen):  # 合并单元格只取一次
                continue
            seen.append(cell._tc)
            cells.append(cell.text.strip().replace("\n", " "))
        lines.append(" | ".join(cells))
    return lines


_REL_ATTRS = (qn("r:embed"), qn("r:link"), qn("r:id"))


def _copy_relationships(element, src_part, dst_part) -> List[Dict]:
    """复制元素引用的图片与外部链接关系，并把元素里的关系 ID 改成目标文档里的新 ID。"""
    warnings, mapping = [], {}
    for el in element.iter():
        for attr in _REL_ATTRS:
            rid = el.get(attr)
            if not rid or rid not in src_part.rels:
                continue
            if rid not in mapping:
                rel = src_part.rels[rid]
                if rel.is_external:
                    mapping[rid] = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    mapping[rid], _ = dst_part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                else:
                    mapping[rid] = None
                    warnings.append({"code": "MERGE_PARTIAL",
                                     "message": f"未能合并内嵌对象：{rel.reltype.rsplit('/', 1)[-1]}"})
            if mapping[rid]:
                el.set(attr, mapping[rid])
    if element.find(".//" + qn("w:numPr")) is not None:
        warnings.append({"code": "MERGE_NUMBERING", "message": "列表编号沿用第一份文档的编号定义，格式可能与原文档不同"})
    return warnings


def merge_documents(input_paths: List[str], output_path: str, **kwargs) -> Dict[str, Any]:
    """
    合并多个 Word 文档：以第一份为底（沿用其样式与页面设置），后续文档另起一页接在后面。
    任一文件不存在时报错，不会只合并一部分。
    """
    missing = [p for p in input_paths if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError(f"要合并的文件不存在：{', '.join(missing)}")
    merged = None
    merged_count = 0
    warnings: List[Dict] = []

    for path in input_paths:
        if merged is None:
            merged = Document(path)
        else:
            source = Document(path)
            body = merged.element.body
            anchor = body.sectPr
            page_break = OxmlElement("w:p")
            run = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run.append(br)
            page_break.append(run)
            for element in [page_break] + [e for e in source.element.body if e.tag != qn("w:sectPr")]:
                if element is not page_break:
                    element = deepcopy(element)
                    warnings.extend(_copy_relationships(element, source.part, merged.part))
                if anchor is not None:
                    anchor.addprevious(element)
                else:
                    body.append(element)
        merged_count += 1

    if merged is None:
        merged = Document()
    merged.save(output_path)

    result: Dict[str, Any] = {"output_path": output_path, "merged_count": merged_count}
    if warnings:
        result["warnings"] = [dict(t) for t in {tuple(w.items()) for w in warnings}]
    return result
