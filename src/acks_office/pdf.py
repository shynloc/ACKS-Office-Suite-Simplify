
import os
from io import BytesIO
from typing import Any, Dict, List, Optional
from xml.sax.saxutils import escape

from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                Preformatted, Image as RLImage)
from reportlab.platypus.flowables import HRFlowable

from . import fonts
from . import markdown_blocks as mdb

_MARGIN = 72
_RULE = colors.HexColor("#CCCCCC")
_ALERT_LABELS = {"note": "Note · 注释", "tip": "Tip · 提示", "important": "Important · 重要",
                 "warning": "Warning · 警告", "caution": "Caution · 注意"}


LEGACY_THEMES = ("acks", "default")


def create_pdf(title: Optional[str], content: str, output_path: str,
               font: Optional[str] = None, bold_font: Optional[str] = None, theme: Any = "acks",
               brand_name: Optional[str] = None, footer_label: Optional[str] = None, **kwargs) -> Dict[str, Any]:
    """
    创建PDF文档
    Args:
        title: 标题；为空时用正文 front matter 里的 title
        content: Markdown 内容（标题、粗体/斜体、链接、列表、表格、引文与提示块、代码块、图片）；
            开头可以写 front matter（kicker、subtitle、author、date、version、issue、lede 等）
        output_path: 输出路径
        font / bold_font: 中文字体文件路径（TrueType），代替主题的中文字体；默认按主题自动选择
        theme: 主题名称（neutral、slate、folio 或已安装的主题）、主题目录或 Theme 对象；
            "acks"、"default" 是 2.x 的内置样式
        brand_name: 品牌名；footer_label: 页脚左侧文字，默认按主题模板生成
        base_dir: 图片相对路径的基准目录，默认当前目录
    """
    if not (isinstance(theme, str) and theme in LEGACY_THEMES):
        from .render.pdf import render_pdf
        meta = {k: v for k, v in kwargs.items() if k not in ("base_dir", "create_chart")}
        if brand_name is not None:
            meta["brand"] = brand_name
        return render_pdf(title, content, output_path, theme=theme, base_dir=kwargs.get("base_dir"),
                          footer_label=footer_label, font=font, bold_font=bold_font, **meta)
    title = title or ""
    pdf_fonts = fonts.pdf_fonts(font, bold_font)
    styles = _styles(pdf_fonts)
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=_MARGIN, leftMargin=_MARGIN,
                            topMargin=_MARGIN, bottomMargin=_MARGIN, title=title)
    renderer = _PdfRenderer(styles, kwargs.get("base_dir") or os.getcwd(), A4[0] - 2 * _MARGIN)
    story = [Paragraph(_escape(title), styles["title"]), Spacer(1, 24)] + renderer.render(mdb.parse(content))
    doc.build(story)

    result: Dict[str, Any] = {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "pages": doc.page,
        "font": pdf_fonts.describe(),
    }
    warnings = renderer.warnings
    if not pdf_fonts.embedded:
        warnings.append({"code": "FONT_FALLBACK",
                         "message": "未找到可嵌入的中文字体，使用了由阅读器替换显示的 STSong-Light；"
                                    "可运行 acks-office fonts install noto-sans-sc 安装开源字体"})
    if warnings:
        result["warnings"] = warnings
    return result


def _styles(f: "fonts.PdfFonts") -> Dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    body = ParagraphStyle("AcksBody", parent=base["Normal"], fontName=f.regular, fontSize=11,
                          leading=19, wordWrap="CJK", textColor=colors.HexColor("#1A1A1A"))
    return {
        "title": ParagraphStyle("AcksTitle", parent=body, fontName=f.bold, fontSize=24, leading=32,
                                textColor=colors.HexColor("#003366"), alignment=1),
        "h1": ParagraphStyle("AcksH1", parent=body, fontName=f.bold, fontSize=18, leading=26,
                             textColor=colors.HexColor("#006699"), spaceBefore=12, spaceAfter=8),
        "h2": ParagraphStyle("AcksH2", parent=body, fontName=f.bold, fontSize=14, leading=21,
                             spaceBefore=10, spaceAfter=6),
        "h3": ParagraphStyle("AcksH3", parent=body, fontName=f.bold, fontSize=12, leading=18,
                             spaceBefore=8, spaceAfter=4),
        "body": ParagraphStyle("AcksP", parent=body, spaceAfter=6),
        "cell": ParagraphStyle("AcksCell", parent=body, fontSize=10, leading=15),
        "cell_right": ParagraphStyle("AcksCellR", parent=body, fontSize=10, leading=15, alignment=2),
        "cell_head": ParagraphStyle("AcksCellH", parent=body, fontName=f.bold, fontSize=10, leading=15),
        "quote": ParagraphStyle("AcksQuote", parent=body, textColor=colors.HexColor("#555555")),
        "code": ParagraphStyle("AcksCode", parent=body, fontSize=9.5, leading=14),
        "caption": ParagraphStyle("AcksCaption", parent=body, fontSize=9, leading=13,
                                  textColor=colors.HexColor("#666666"), alignment=1),
    }


def _escape(text: str) -> str:
    """reportlab 的段落会解析标记；用户文字先转义，再把换行变成 <br/>。"""
    return escape(text).replace("\n", "<br/>")


def _markup(spans: List[mdb.Span]) -> str:
    parts = []
    for s in spans:
        text = _escape(s.text)
        if s.code:
            text = f'<font color="#8B1A1A">{text}</font>'
        if s.bold:
            text = f"<b>{text}</b>"
        if s.italic:
            text = f"<i>{text}</i>"
        if s.strike:
            text = f"<strike>{text}</strike>"
        if s.link:
            text = f'<a href="{escape(s.link, {chr(34): "&quot;"})}" color="#1F5FAE"><u>{text}</u></a>'
        parts.append(text)
    return "".join(parts)


class _PdfRenderer:
    def __init__(self, styles: Dict[str, ParagraphStyle], base_dir: str, width: float):
        self.s = styles
        self.base_dir = base_dir
        self.width = width
        self.story: List = []
        self.warnings: List[Dict] = []

    def render(self, blocks) -> List:
        for block in blocks:
            self.block(block, 0)
        return self.story

    def block(self, block, level: int) -> None:
        s = self.s
        if isinstance(block, mdb.Heading):
            self.story.append(Paragraph(_markup(block.spans), s[f"h{min(block.level, 3)}"]))
        elif isinstance(block, mdb.Paragraph):
            self.story.append(Paragraph(_markup(block.spans), s["body"]))
        elif isinstance(block, mdb.ListBlock):
            style = ParagraphStyle(f"AcksLi{level}", parent=s["body"], leftIndent=18 + 16 * level,
                                   bulletIndent=4 + 16 * level, spaceAfter=2)
            # Helvetica 里没有 ■ □，任务标记改用正文的中文字体
            task = ParagraphStyle(f"AcksTask{level}", parent=style, bulletFontName=style.fontName)
            for n, item in enumerate(block.items, start=block.start):
                if item.checked is not None:
                    para = Paragraph(_markup(item.spans), task, bulletText="■" if item.checked else "□")
                else:
                    para = Paragraph(_markup(item.spans), style, bulletText=f"{n}." if block.ordered else "•")
                self.story.append(para)
                for child in item.children:
                    self.block(child, level + 1)
        elif isinstance(block, mdb.Table):
            if block.header or block.rows:
                self.table(block)
        elif isinstance(block, mdb.Quote):
            self.quote(block)
        elif isinstance(block, mdb.Code):
            box = Table([[Preformatted(block.text, s["code"], maxLineLength=80)]], colWidths=[self.width])
            box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F5F5")),
                                     ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8)]))
            self.story += [box, Spacer(1, 8)]
        elif isinstance(block, mdb.Rule):
            self.story.append(HRFlowable(width="100%", thickness=0.5, color=_RULE, spaceBefore=6, spaceAfter=6))
        elif isinstance(block, mdb.Image):
            self.image(block)

    def table(self, block: mdb.Table) -> None:
        s = self.s
        width = max([len(block.header)] + [len(r) for r in block.rows])
        right = {i for i, a in enumerate(block.aligns) if a == "right"}
        data = [[Paragraph(_markup(c), s["cell_head"]) for c in block.header] + [""] * (width - len(block.header))]
        for row in block.rows:
            cells = [Paragraph(_markup(c), s["cell_right"] if i in right else s["cell"]) for i, c in enumerate(row)]
            data.append(cells + [""] * (width - len(cells)))
        table = Table(data, colWidths=[self.width / width] * width, repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, _RULE),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F4F7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        self.story += [table, Spacer(1, 10)]

    def quote(self, block: mdb.Quote) -> None:
        s = self.s
        inner = []
        if block.alert:
            inner.append(Paragraph(f"<b>{_escape(_ALERT_LABELS[block.alert])}</b>", s["quote"]))
        for b in block.blocks:
            if isinstance(b, (mdb.Paragraph, mdb.Heading)):
                inner.append(Paragraph(_markup(b.spans), s["quote"]))
            elif isinstance(b, mdb.ListBlock):
                inner += [Paragraph("· " + _markup(i.spans), s["quote"]) for i in b.items]
        if not inner:
            return
        box = Table([[inner]], colWidths=[self.width])
        box.setStyle(TableStyle([("LINEBEFORE", (0, 0), (0, -1), 2, colors.HexColor("#BBBBBB")),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 12)]))
        self.story += [box, Spacer(1, 8)]

    def image(self, block: mdb.Image) -> None:
        path = block.src if os.path.isabs(block.src) else os.path.join(self.base_dir, block.src)
        if not os.path.isfile(path):
            self.warnings.append({"code": "IMAGE_NOT_FOUND", "message": f"图片不存在：{block.src}"})
            self.story.append(Paragraph(_escape(f"[图片：{block.alt or block.src}]"), self.s["body"]))
            return
        img = RLImage(path)
        scale = min(1.0, self.width / img.imageWidth)
        img.drawWidth, img.drawHeight = img.imageWidth * scale, img.imageHeight * scale
        self.story.append(img)
        if block.alt:
            self.story.append(Paragraph(_escape(block.alt), self.s["caption"]))


def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None,
                  opacity: float = 0.15, position: str = "diagonal", **kwargs) -> Dict[str, Any]:
    """
    给 PDF 添加水印（半透明大字，按每页实际尺寸居中）。

    Args:
        input_path: 输入文件路径
        watermark_text: 水印文本（支持中文）
        output_path: 输出文件路径（默认覆盖输入）
        opacity: 透明度 0~1，默认 0.15
        position: 水印位置 "diagonal"（沿对角线斜排）或 "center"
    """
    if not output_path:
        output_path = input_path

    writer = PdfWriter(clone_from=input_path)
    stamps: Dict[tuple, Any] = {}
    for page in writer.pages:
        box = page.mediabox
        size = (float(box.width), float(box.height))
        if size not in stamps:  # 同尺寸的页面共用一张水印
            stamps[size] = PdfReader(_build_watermark_page(watermark_text, opacity, position, size)).pages[0]
        page.merge_transformed_page(stamps[size], Transformation().translate(float(box.left), float(box.bottom)))

    with open(output_path, 'wb') as f:
        writer.write(f)

    return {"output_path": output_path, "watermarked_pages": len(writer.pages)}


def _build_watermark_page(text: str, opacity: float, position: str, size=A4):
    """用 reportlab 生成与页面同尺寸的单页水印 PDF，返回 BytesIO。"""
    import math
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    w, h = size
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=size)
    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.setFillAlpha(max(0.0, min(1.0, opacity)))
    # 含中文用可嵌入的中文字体，否则用 Helvetica
    font = fonts.pdf_fonts().bold if any('一' <= ch <= '鿿' for ch in text) else 'Helvetica-Bold'
    angle = math.degrees(math.atan2(h, w)) if position == "diagonal" else 0
    room = 0.8 * (math.hypot(w, h) if angle else w)
    font_size = min(min(w, h) / 10, room / max(stringWidth(text, font, 1), 1e-6))
    c.setFont(font, font_size)

    c.translate(w / 2, h / 2)
    c.rotate(angle)
    c.drawCentredString(0, -font_size / 3, text)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def extract_text(input_path: str, **kwargs) -> str:
    """
    提取PDF文本内容
    """
    reader = PdfReader(input_path)
    return '\n'.join(page.extract_text() or "" for page in reader.pages)

def merge_pdfs(input_paths: List[str], output_path: str, **kwargs) -> Dict[str, Any]:
    """
    合并多个PDF文件；不存在的文件会列在结果的 skipped 里
    """
    writer = PdfWriter()
    skipped = []

    for path in input_paths:
        if not os.path.exists(path):
            skipped.append(path)
            continue
        for page in PdfReader(path).pages:
            writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)

    result: Dict[str, Any] = {"output_path": output_path, "merged_pages": len(writer.pages)}
    if skipped:
        result["skipped"] = skipped
    return result

def split_pdf(input_path: str, output_dir: str, split_by: str = "page", **kwargs) -> Dict[str, Any]:
    """
    拆分PDF文件：每页一个文件
    """
    os.makedirs(output_dir, exist_ok=True)
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)

    for page_num in range(total_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[page_num])

        output_path = os.path.join(output_dir, f"page_{page_num + 1}.pdf")
        with open(output_path, 'wb') as f:
            writer.write(f)

    return {"output_dir": output_dir, "total_pages": total_pages}
