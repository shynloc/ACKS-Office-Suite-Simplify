"""兼容层：包名已改为 acks_office，原来的 `import office_suite` 写法暂时可用（4.0 移除）。"""

import importlib
import sys
import warnings

import acks_office
from acks_office import OfficeSuite

warnings.warn("office_suite 已更名为 acks_office，请改用 `import acks_office`；旧名将在 4.0 移除",
              DeprecationWarning, stacklevel=2)

__version__ = acks_office.__version__
__author__ = acks_office.__author__
__all__ = ["OfficeSuite"]

# 让 office_suite.xlsx 等旧路径指向同一个模块对象；2.x 的 office_suite.design_system 已在 3.0 移除
for _name in ("core", "docx", "xlsx", "pdf", "pptx", "email", "utils"):
    _module = importlib.import_module(f"acks_office.{_name}")
    sys.modules[f"{__name__}.{_name}"] = _module
    if "." not in _name:
        globals()[_name] = _module
