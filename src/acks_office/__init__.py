"""ACKS Office Suite Simplify：按主题生成 Word、Excel、PPT、PDF，读取、转换办公文档。

    import acks_office
    acks_office.create("word", "报告.docx", content="# 标题\n\n正文", theme="slate")
    rows = acks_office.extract("数据.xlsx")
"""

__version__ = "2.1.2"
__author__ = "Thom Jing (shynloc)"

__all__ = ["create", "extract", "convert", "add_watermark", "merge", "OfficeSuite"]

_API = ("create", "extract", "convert", "add_watermark", "merge")


def __getattr__(name):
    # 延迟导入：只运行 CLI 或 doctor 时，不要求 python-docx、pandas 等依赖已经装齐
    if name in _API:
        from . import api
        return getattr(api, name)
    if name == "OfficeSuite":
        from .core import OfficeSuite
        return OfficeSuite
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
