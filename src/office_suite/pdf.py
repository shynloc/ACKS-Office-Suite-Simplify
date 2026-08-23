
import os
from typing import Optional, List, Dict, Any
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter

def create_pdf(title: str, content: str, output_path: str, **kwargs) -> Dict[str, Any]:
    """
    创建PDF文档
    """
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    styles = getSampleStyleSheet()
    
    # 自定义样式
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        textColor=colors.HexColor('#003366'),
        alignment=1  # 居中
    )
    
    heading1_style = ParagraphStyle(
        'CustomHeading1',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#006699')
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=12,
        leading=18
    )
    
    # 构建内容
    story = []
    
    # 添加标题
    story.append(Paragraph(title, title_style))
    story.append(Spacer(1, 30))
    
    # 解析内容
    lines = content.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
            continue
            
        if line.startswith('# '):
            story.append(Paragraph(line[2:], heading1_style))
            story.append(Spacer(1, 12))
        elif line.startswith('## '):
            story.append(Paragraph(line[3:], styles['Heading2']))
            story.append(Spacer(1, 10))
        elif line.startswith('### '):
            story.append(Paragraph(line[4:], styles['Heading3']))
            story.append(Spacer(1, 8))
        elif line.startswith('- ') or line.startswith('* '):
            story.append(Paragraph(f'• {line[2:]}', normal_style))
        else:
            story.append(Paragraph(line, normal_style))
    
    # 生成PDF
    doc.build(story)
    
    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "pages": len(story) // 20 + 1  # 估算页数
    }

def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None,
                  opacity: float = 0.15, position: str = "diagonal", **kwargs) -> Dict[str, Any]:
    """
    给 PDF 添加水印（半透明斜体，reportlab 生成水印页 + merge_page 合并）。

    Args:
        input_path: 输入文件路径
        watermark_text: 水印文本（支持中文）
        output_path: 输出文件路径（默认覆盖输入）
        opacity: 透明度 0~1，默认 0.15
        position: 水印位置 "diagonal"（45° 斜排）或 "center"
    """
    if not output_path:
        output_path = input_path

    reader = PdfReader(input_path)
    writer = PdfWriter()

    watermark_stream = _build_watermark_page(watermark_text, opacity, position)
    watermark_reader = PdfReader(watermark_stream)
    watermark_page = watermark_reader.pages[0]

    for page in reader.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)

    return {"output_path": output_path, "watermarked_pages": len(writer.pages)}


def _build_watermark_page(text: str, opacity: float, position: str):
    """用 reportlab 生成单页 A4 水印 PDF，返回 BytesIO。"""
    from io import BytesIO
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.setFillAlpha(max(0.0, min(1.0, opacity)))

    font_size = 60
    # 含中文用 CID 宋体，否则用 Helvetica
    if any('\u4e00' <= ch <= '\u9fff' for ch in text):
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        c.setFont('STSong-Light', font_size)
    else:
        c.setFont('Helvetica-Bold', font_size)

    c.saveState()
    c.translate(w / 2, h / 2)
    if position == "diagonal":
        c.rotate(45)
    c.drawCentredString(0, 0, text)
    c.restoreState()

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def extract_text(input_path: str, **kwargs) -> str:
    """
    提取PDF文本内容
    """
    reader = PdfReader(input_path)
    full_text = []
    
    for page in reader.pages:
        full_text.append(page.extract_text())
        
    return '\n'.join(full_text)

def merge_pdfs(input_paths: List[str], output_path: str, **kwargs) -> Dict[str, Any]:
    """
    合并多个PDF文件
    """
    writer = PdfWriter()
    
    for path in input_paths:
        if not os.path.exists(path):
            continue
            
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)
            
    with open(output_path, 'wb') as f:
        writer.write(f)
        
    return {"output_path": output_path, "merged_pages": len(writer.pages)}

def split_pdf(input_path: str, output_dir: str, split_by: str = "page", **kwargs) -> Dict[str, Any]:
    """
    拆分PDF文件
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
