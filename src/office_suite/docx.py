
import os
from typing import Optional, Dict, Any
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION

def create_word(title: str, content: str, output_path: str, **kwargs) -> Dict[str, Any]:
    """
    创建Word文档
    Args:
        title: 文档标题
        content: 文档内容，支持Markdown格式的标题、列表
        output_path: 输出文件路径
    """
    # 创建文档
    doc = Document()
    
    # 设置默认字体
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal'].font.size = Pt(12)
    
    # 添加标题
    title_para = doc.add_heading(title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.runs[0].font.size = Pt(24)
    title_para.runs[0].font.bold = True
    title_para.runs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    doc.add_paragraph()  # 空行
    
    # 解析内容并添加
    lines = content.split('\n')
    for line in lines:
        line = line.rstrip()
        if not line:
            doc.add_paragraph()
            continue
            
        # 标题
        if line.startswith('# '):
            para = doc.add_heading(line[2:], level=1)
            para.runs[0].font.bold = True
            para.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        elif line.startswith('## '):
            para = doc.add_heading(line[3:], level=2)
            para.runs[0].font.bold = True
            para.runs[0].font.color.rgb = RGBColor(51, 153, 255)
        elif line.startswith('### '):
            para = doc.add_heading(line[4:], level=3)
            para.runs[0].font.bold = True
        # 列表
        elif line.startswith('- ') or line.startswith('* '):
            para = doc.add_paragraph(style='List Bullet')
            para.add_run(line[2:])
        elif line.startswith('1. ') or line.startswith('2. ') or line.startswith('3. '):
            para = doc.add_paragraph(style='List Number')
            para.add_run(line)
        # 普通段落
        else:
            para = doc.add_paragraph(line)
    
    # 保存文档
    doc.save(output_path)
    
    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "pages": int(len(doc.paragraphs) / 30) + 1  # 估算页数
    }

def add_watermark(input_path: str, watermark_text: str, output_path: Optional[str] = None, **kwargs) -> Dict[str, Any]:
    """
    给Word文档添加水印
    """
    if not output_path:
        output_path = input_path
        
    doc = Document(input_path)
    
    # 添加水印（通过页眉）
    section = doc.sections[0]
    header = section.header
    
    # 创建水印段落
    watermark_para = header.add_paragraph()
    watermark_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run = watermark_para.add_run(watermark_text)
    run.font.size = Pt(48)
    run.font.color.rgb = RGBColor(200, 200, 200)  # 灰色
    run.bold = True
    
    # 保存
    doc.save(output_path)
    
    return {"output_path": output_path}

def extract_text(input_path: str, **kwargs) -> str:
    """
    提取Word文档中的文本
    """
    doc = Document(input_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def merge_documents(input_paths: List[str], output_path: str, **kwargs) -> Dict[str, Any]:
    """
    合并多个Word文档
    """
    merged_doc = Document()
    
    for i, path in enumerate(input_paths):
        if not os.path.exists(path):
            continue
            
        # 加入分页符（第一个文档除外）
        if i > 0:
            merged_doc.add_section(WD_SECTION.NEW_PAGE)
            
        # 读取源文档
        source_doc = Document(path)
        
        # 复制所有元素
        for element in source_doc.element.body:
            merged_doc.element.body.append(element)
    
    merged_doc.save(output_path)
    return {"output_path": output_path, "merged_count": len(input_paths)}
