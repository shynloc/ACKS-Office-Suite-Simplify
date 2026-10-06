import os
from typing import Optional, List, Dict, Any
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

DEFAULT_BRAND = "ACKS Studio"
TRANSITIONS = ("fade", "push", "wipe", "split", "cover", "pull", "dissolve", "cut", "zoom", "random")


def create_pptx(title: str, slides: List[Dict[str, Any]], output_path: str,
                theme: Any = "acks",
                brand_name: Optional[str] = None,
                **kwargs) -> Dict[str, Any]:
    """
    创建 PowerPoint 演示文稿。

    Args:
        title: 演示文稿标题
        slides: 幻灯片列表。每页一个字典，layout 为 title（封面）、section（章节页）、content（内容页，
            content 或 bullets）、data（关键数字 kpis + 条形图 chart）、table（表格）、number（大数字）、
            quote（引文）、image（图片），其余字段见命令参考；acks / default 只认 title 与 content
        output_path: 输出路径
        theme: 主题名称（neutral、slate、folio 或已安装的主题）、主题目录或 Theme 对象；
            "acks"、"default" 是 2.x 的内置样式
        brand_name: 品牌名，传 "" 去品牌化
        footer_label: 页脚左侧文字，默认按主题模板生成
        meta_right: 标题页右上角文字（theme="acks" 生效）
        其余关键字参数（kicker、classification、publication、issue、season、date、author 等）作为元数据
    """
    if isinstance(theme, str) and theme in ("acks", "default"):
        brand = DEFAULT_BRAND if brand_name is None else brand_name
        if theme == "acks":
            return _create_pptx_acks(title, slides, output_path, brand, **kwargs)
        return _create_pptx_default(title, slides, output_path, **kwargs)
    from .render.slides import render_slides
    options = ("base_dir", "font_policy", "footer_label")
    meta = {k: v for k, v in kwargs.items() if k not in options}
    if brand_name is not None:
        meta["brand"] = brand_name
    return render_slides(title, slides, output_path, theme=theme, base_dir=kwargs.get("base_dir"),
                         font_policy=kwargs.get("font_policy", "local"), footer_label=kwargs.get("footer_label"),
                         **meta)


def _create_pptx_acks(title: str, slides: List[Dict[str, Any]], output_path: str,
                      brand_name: str, **kwargs) -> Dict[str, Any]:
    """用 ACKS 设计规范（design_system.slides）生成 PPT。"""
    from .design_system import slides as acks
    from .docx import _is_cjk

    prs = acks.init_presentation()
    acks.set_brand(prs, name=brand_name, footer_label=kwargs.get("footer_label"))
    meta_right = kwargs.get("meta_right", "CONFIDENTIAL · 2026" if brand_name == DEFAULT_BRAND else "")

    for idx, slide_data in enumerate(slides, start=1):
        layout = slide_data.get("layout", "content")
        s_title = slide_data.get("title", "") or title
        s_content = slide_data.get("content", "")
        subtitle = slide_data.get("subtitle", "")

        if layout == "title":
            if _is_cjk(s_title):
                acks.add_title_slide(prs, doctype="", title_top="", title_em=s_title,
                                     zh_sub=subtitle, meta_right=meta_right)
            else:
                acks.add_title_slide(prs, doctype="", title_top=s_title, title_em="",
                                     zh_sub=subtitle, meta_right=meta_right)
        else:
            paragraphs = [ln.strip() for ln in s_content.split('\n') if ln.strip()] or [""]
            acks.add_content_slide(
                prs,
                eyebrow=brand_name,
                title_en=s_title,
                title_zh=subtitle,
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
            for line in lines:
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


def add_transition_effects(input_path: str, output_path: Optional[str] = None, effect: str = "fade",
                           duration: float = 0.5, **kwargs) -> Dict[str, Any]:
    """
    给每张幻灯片添加切换效果
    Args:
        effect: fade / push / wipe / split / cover / pull / dissolve / cut / zoom / random
        duration: 秒，映射为 PowerPoint 的快（≤0.5）/ 中（≤1）/ 慢 三档
        output_path: 输出路径（默认覆盖输入）
    """
    if effect not in TRANSITIONS:
        raise ValueError(f"不支持的切换效果：{effect}；可选：{', '.join(TRANSITIONS)}")
    if not output_path:
        output_path = input_path

    speed = "fast" if duration <= 0.5 else "med" if duration <= 1.0 else "slow"
    prs = Presentation(input_path)
    for slide in prs.slides:
        sld = slide._element
        for old in sld.findall(qn("p:transition")):
            sld.remove(old)
        transition = parse_xml(f'<p:transition {nsdecls("p")} spd="{speed}"><p:{effect}/></p:transition>')
        # 元素顺序：cSld、clrMapOvr、transition、timing
        timing = sld.find(qn("p:timing"))
        if timing is not None:
            timing.addprevious(transition)
        else:
            anchor = sld.find(qn("p:clrMapOvr"))
            (anchor if anchor is not None else sld.find(qn("p:cSld"))).addnext(transition)

    prs.save(output_path)
    return {"output_path": output_path, "slides": len(prs.slides), "effect": effect}


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
