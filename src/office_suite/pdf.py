
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

def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None, **kwargs) -> Dict[str, Any]:
    """
    给PDF添加水印
    注：完整的水印功能需要使用reportlab生成水印页面然后合并，这里是简化版本
    """
    if not output_path:
        output_path = input_path
        
    # TODO: 实现完整水印功能
    # 临时方案：直接复制文件
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    for page in reader.pages:
        writer.add_page(page)
        
    with open(output_path, 'wb') as f:
        writer.write(f)
        
    return {"output_path": output_path}

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
