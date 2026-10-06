"""ACKS Office Suite Simplify：按文档设计规范生成、读取、转换 Word / Excel / PPT / PDF，不依赖本机 Office。"""

__version__ = "2.1.0"
__author__ = "Thom Jing (shynloc)"

__all__ = ["OfficeSuite"]


def __getattr__(name):
    # 延迟导入：只运行 CLI 或 doctor 时，不要求 python-docx、pandas 等依赖已经装齐
    if name == "OfficeSuite":
        from .core import OfficeSuite
        return OfficeSuite
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
