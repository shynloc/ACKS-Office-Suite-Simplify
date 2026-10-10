
import os
from io import BytesIO
from typing import Any, Dict, List, Optional

from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.lib.pagesizes import A4

from . import fonts


def create_pdf(title: Optional[str], content: str, output_path: str,
               font: Optional[str] = None, bold_font: Optional[str] = None, theme: Any = None,
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
            "default" 是 neutral 的别名
        brand_name: 品牌名；footer_label: 页脚左侧文字，默认按主题模板生成
        base_dir: 图片相对路径的基准目录，默认当前目录
    """
    from .render.pdf import render_pdf
    meta = {k: v for k, v in kwargs.items() if k not in ("base_dir", "create_chart")}
    if brand_name is not None:
        meta["brand"] = brand_name
    return render_pdf(title, content, output_path, theme=theme, base_dir=kwargs.get("base_dir"),
                      footer_label=footer_label, font=font, bold_font=bold_font, **meta)


def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None,
                  opacity: float = 0.15, position: str = "diagonal", **kwargs) -> Dict[str, Any]:
    """
    给 PDF 添加水印（半透明大字，按每页实际尺寸居中）。

    Args:
        input_path: 输入文件路径
        watermark_text: 水印文本（支持中文）
        output_path: 输出文件路径，默认为 <原名>_watermarked.pdf，不覆盖原文件
        opacity: 透明度 0~1，默认 0.15
        position: 水印位置 "diagonal"（沿对角线斜排）或 "center"
    """
    if not output_path:
        from .api import watermarked_path
        output_path = watermarked_path(input_path)

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
    合并多个 PDF 文件。任一文件不存在时报错，不会只合并一部分。
    """
    missing = [p for p in input_paths if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError(f"要合并的文件不存在：{', '.join(missing)}")
    writer = PdfWriter()
    for path in input_paths:
        for page in PdfReader(path).pages:
            writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)

    return {"output_path": output_path, "merged_pages": len(writer.pages)}

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
