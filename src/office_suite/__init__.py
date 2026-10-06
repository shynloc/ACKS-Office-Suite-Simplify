"""兼容层：包名已改为 acks_office，原来的 `import office_suite` 写法继续可用（3.0 移除）。"""

import importlib
import sys
import warnings

import acks_office
from acks_office import OfficeSuite

warnings.warn("office_suite 已更名为 acks_office，请改用 `import acks_office`；旧名将在 3.0 移除",
              DeprecationWarning, stacklevel=2)

__version__ = acks_office.__version__
__author__ = acks_office.__author__
__all__ = ["OfficeSuite"]

# 让 office_suite.xlsx、office_suite.design_system.tokens 等旧路径指向同一个模块对象
for _name in ("core", "docx", "xlsx", "pdf", "pptx", "email", "utils",
              "design_system", "design_system.tokens", "design_system.slides",
              "design_system.xlsx", "design_system.tokens_constants",
              "design_system.slides_constants", "design_system.xlsx_constants"):
    _module = importlib.import_module(f"acks_office.{_name}")
    sys.modules[f"{__name__}.{_name}"] = _module
    if "." not in _name:
        globals()[_name] = _module
