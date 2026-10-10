"""函数式接口：生成、读取、转换、加水印、合并。

出错时抛出异常（FileNotFoundError、ValueError、ThemeError 等），不再像 OfficeSuite 那样返回
{"success": False} 字典。所有写文件的操作都写到新文件，不覆盖原文件。
"""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional

KINDS = {"word": "word", "docx": "word", "excel": "excel", "xlsx": "excel", "pdf": "pdf",
         "pptx": "pptx", "ppt": "pptx", "powerpoint": "pptx"}


def create(kind: str, output_path: str, **kwargs: Any) -> Dict[str, Any]:
    """生成文档。kind：word / excel / pdf / pptx（也可写 docx、xlsx、ppt）。

    其余参数见 docx.create_word、xlsx.create_excel、pdf.create_pdf、pptx.create_pptx：
    常用的有 title、content（Word / PDF 的 Markdown）、data 或 sheets（Excel）、slides（PPT）、
    theme（主题名称或目录，默认 neutral）以及 subtitle、author、date 等元数据。
    """
    target = KINDS.get(str(kind).lower())
    if target is None:
        raise ValueError(f"不支持的文档类型：{kind}；可选：word、excel、pdf、pptx")
    if target == "word":
        from .docx import create_word as build
    elif target == "excel":
        from .xlsx import create_excel as build
    elif target == "pdf":
        from .pdf import create_pdf as build
    else:
        from .pptx import create_pptx as build
    kwargs.setdefault("title", None)
    return build(output_path=output_path, **kwargs)


def _require(path: str) -> None:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"文件不存在：{path}")


def extract(input_path: str, **kwargs: Any) -> Any:
    """读取文档：Excel 返回行字典列表；Word、PDF 返回文本（表格按行输出）；PPT 返回各页文字。"""
    _require(input_path)
    ext = Path(input_path).suffix.lower()
    if ext in (".xlsx", ".xls"):
        from .xlsx import extract_data
        return extract_data(input_path, **kwargs)
    if ext == ".docx":
        from .docx import extract_text
        return extract_text(input_path, **kwargs)
    if ext == ".pdf":
        from .pdf import extract_text
        return extract_text(input_path, **kwargs)
    if ext == ".pptx":
        from .pptx import extract_text
        return extract_text(input_path, **kwargs)
    raise ValueError(f"不支持读取 {ext or '无扩展名'} 文件；旧版 .doc、.ppt 可先用 convert 转换")


def convert(input_path: str, to: str, output_path: Optional[str] = None, **kwargs: Any) -> Dict[str, Any]:
    """格式转换（需要 LibreOffice）：to 为目标格式，如 pdf、docx、xlsx、pptx。默认输出到原文件旁边。"""
    _require(input_path)
    from .utils import convert_with_libreoffice
    target = to.lower().split(":")[0]
    output_path = output_path or str(Path(input_path).with_suffix("." + target))
    if os.path.abspath(output_path) == os.path.abspath(input_path):
        raise ValueError("输出文件和原文件相同，请指定 output_path")
    result = convert_with_libreoffice(input_path, to.lower(), output_path, **kwargs)
    return {"output_path": output_path, **result}


def watermarked_path(input_path: str) -> str:
    source = Path(input_path)
    return str(source.with_name(f"{source.stem}_watermarked{source.suffix}"))


def add_watermark(input_path: str, text: str, output_path: Optional[str] = None, **kwargs: Any) -> Dict[str, Any]:
    """给 PDF 或 Word 加水印。默认写到新文件 <原名>_watermarked.<扩展名>，不覆盖原文件。"""
    _require(input_path)
    output_path = output_path or watermarked_path(input_path)
    ext = Path(input_path).suffix.lower()
    if ext == ".pdf":
        from .pdf import add_watermark as stamp
    elif ext == ".docx":
        from .docx import add_watermark as stamp
    else:
        raise ValueError(f"不支持给 {ext or '无扩展名'} 文件加水印，支持 .pdf、.docx")
    return stamp(input_path, text, output_path, **kwargs)


def merge(input_paths: List[str], output_path: str) -> Dict[str, Any]:
    """合并多个 PDF 或多个 Word。任一文件不存在时报错，不会只合并一部分。"""
    missing = [p for p in input_paths if not os.path.isfile(p)]
    if missing:
        raise FileNotFoundError(f"要合并的文件不存在：{', '.join(missing)}")
    suffixes = {Path(p).suffix.lower() for p in input_paths}
    if suffixes == {".pdf"}:
        from .pdf import merge_pdfs
        return merge_pdfs(input_paths, output_path)
    if suffixes == {".docx"}:
        from .docx import merge_documents
        return merge_documents(input_paths, output_path)
    raise ValueError("只能合并同一种格式：全部 .pdf 或全部 .docx")
