import os
from typing import Optional, List, Dict, Any
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


def create_pptx(title: str, slides: List[Dict[str, Any]], output_path: str,
                theme: str = "acks",
                brand_name: str = "ACKS Studio",
                **kwargs) -> Dict[str, Any]:
    """
    创建 PowerPoint 演示文稿。

    Args:
        title: 演示文稿标题
        slides: 幻灯片列表，每个字典包含 title, content, layout 属性
        output_path: 输出路径
        theme: "acks"（符合 ACKS 设计规范，默认）或 "default"（简单样式）
        brand_name: 内容页 eyebrow 品牌名（theme="acks" 生效）
    """
    if theme == "acks":
        return _create_pptx_acks(title, slides, output_path, brand_name, **kwargs)
    return _create_pptx_default(title, slides, output_path, **kwargs)


def _create_pptx_acks(title: str, slides: List[Dict[str, Any]], output_path: str,
                      brand_name: str, **kwargs) -> Dict[str, Any]:
    """用 ACKS 设计规范（design_system.slides）生成 PPT。"""
    from .design_system import slides as acks
    from .docx import _is_cjk

    prs = acks.init_presentation()

    for idx, slide_data in enumerate(slides, start=1):
        layout = slide_data.get("layout", "content")
        s_title = slide_data.get("title", "") or title
        s_content = slide_data.get("content", "")

        if layout == "title":
            if _is_cjk(s_title):
                acks.add_title_slide(prs, doctype="", title_top="",
                                     title_em=s_title, zh_sub=s_title)
            else:
                acks.add_title_slide(prs, doctype="", title_top=s_title,
                                     title_em="", zh_sub="")
        else:
            paragraphs = [ln.strip() for ln in s_content.split('\n') if ln.strip()]
            acks.add_content_slide(
                prs,
                eyebrow=brand_name,
                title_en=s_title,
                title_zh=s_title if _is_cjk(s_title) else "",
                paragraphs=paragraphs,
                page_no=idx,
            )

    acks.finalize_footers(prs)
    prs.save(output_path)

    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "slides_count": len(prs.slides),
        "theme": "acks",
    }


def _create_pptx_default(title: str, slides: List[Dict[str, Any]], output_path: str,
                         **kwargs) -> Dict[str, Any]:
    """原简单样式（内置布局 + 蓝色标题）。"""
    prs = Presentation()

    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for slide_data in slides:
        layout_type = slide_data.get("layout", "content")

        if layout_type == "title":
            slide_layout = prs.slide_layouts[0]
        elif layout_type == "title_only":
            slide_layout = prs.slide_layouts[5]
        else:
            slide_layout = prs.slide_layouts[1]

        slide = prs.slides.add_slide(slide_layout)

        if "title" in slide_data and slide.shapes.title:
            title_shape = slide.shapes.title
            title_shape.text = slide_data["title"]
            for para in title_shape.text_frame.paragraphs:
                para.font.size = Pt(36 if layout_type == "title" else 28)
                para.font.bold = True
                para.font.color.rgb = RGBColor(0, 51, 102)
                para.alignment = PP_ALIGN.CENTER if layout_type == "title" else PP_ALIGN.LEFT

        if "content" in slide_data and len(slide.placeholders) >= 2:
            content_placeholder = slide.placeholders[1]
            tf = content_placeholder.text_frame
            tf.text = ""

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

    prs.save(output_path)

    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "slides_count": len(prs.slides),
        "theme": "default",
    }


def add_transition_effects(input_path: str, output_path: Optional[str] = None, effect: str = "fade", duration: float = 0.5, **kwargs) -> Dict[str, Any]:
    """
    给幻灯片添加过渡效果
    """
    if not output_path:
        output_path = input_path

    prs = Presentation(input_path)

    for slide in prs.slides:
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
