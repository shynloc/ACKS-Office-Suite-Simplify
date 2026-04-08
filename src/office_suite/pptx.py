
import os
from typing import Optional, List, Dict, Any
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_pptx(title: str, slides: List[Dict[str, Any]], output_path: str, **kwargs) -> Dict[str, Any]:
    """
    创建PowerPoint演示文稿
    Args:
        title: 演示文稿标题
        slides: 幻灯片列表，每个字典包含title, content, layout属性
        output_path: 输出路径
    """
    prs = Presentation()
    
    # 设置幻灯片大小为宽屏16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    for slide_data in slides:
        layout_type = slide_data.get("layout", "content")
        
        # 选择布局
        if layout_type == "title":
            slide_layout = prs.slide_layouts[0]  # 标题幻灯片
        elif layout_type == "title_only":
            slide_layout = prs.slide_layouts[5]  # 仅标题
        else:
            slide_layout = prs.slide_layouts[1]  # 标题和内容
        
        slide = prs.slides.add_slide(slide_layout)
        
        # 设置标题
        if "title" in slide_data and slide.shapes.title:
            title_shape = slide.shapes.title
            title_shape.text = slide_data["title"]
            # 设置标题样式
            for para in title_shape.text_frame.paragraphs:
                para.font.size = Pt(36 if layout_type == "title" else 28)
                para.font.bold = True
                para.font.color.rgb = RGBColor(0, 51, 102)
                para.alignment = PP_ALIGN.CENTER if layout_type == "title" else PP_ALIGN.LEFT
        
        # 设置内容
        if "content" in slide_data and len(slide.placeholders) >= 2:
            content_placeholder = slide.placeholders[1]
            tf = content_placeholder.text_frame
            tf.text = ""  # 清空默认内容
            
            lines = slide_data["content"].split('\n')
            for i, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue
                    
                p = tf.add_paragraph()
                p.text = line
                p.font.size = Pt(24)
                p.font.color.rgb = RGBColor(0, 0, 0)
                if line.startswith('•') or line.startswith('-'):
                    p.level = 1
                else:
                    p.level = 0
    
    # 保存文件
    prs.save(output_path)
    
    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "slides_count": len(prs.slides)
    }

def add_transition_effects(input_path: str, output_path: Optional[str] = None, effect: str = "fade", duration: float = 0.5, **kwargs) -> Dict[str, Any]:
    """
    给幻灯片添加过渡效果
    """
    if not output_path:
        output_path = input_path
        
    prs = Presentation(input_path)
    
    # TODO: 实现完整的过渡效果功能
    # 目前版本只实现基本的效果设置
    for slide in prs.slides:
        # 简单设置淡入淡出效果
        if hasattr(slide, 'transition'):
            slide.transition.type = effect
            slide.transition.duration = duration
    
    prs.save(output_path)
    return {"output_path": output_path}

def extract_text(input_path: str, **kwargs) -> Dict[int, str]:
    """
    提取PPT中的文本，返回字典，key是幻灯片编号，value是该页文本
    """
    prs = Presentation(input_path)
    result = {}
    
    for i, slide in enumerate(prs.slides, 1):
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_text.append(shape.text)
        result[i] = '\n'.join(slide_text)
    
    return result
