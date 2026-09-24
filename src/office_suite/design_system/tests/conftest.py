"""pytest configuration — adds src/ to sys.path so `import office_suite.design_system` works."""

import sys
from pathlib import Path

# 模块已改为包内相对 import（见 SOURCE.md「本地改动」），测试需经由 office_suite.design_system 导入
SRC = Path(__file__).resolve().parents[3]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
