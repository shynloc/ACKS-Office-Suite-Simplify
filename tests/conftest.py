"""pytest configuration — adds src/ to sys.path so `import acks_office` works without installing."""

import string
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

# 测试里用到的中文，都要在 cjk_font 生成的字体里
CJK_TEXT = "中文测试报告第一二三四行段落粗体斜体删除代码链接区域营收合计华东华南完成待办步骤■□"


@pytest.fixture(scope="session")
def cjk_font(tmp_path_factory):
    """现场生成一个含中文字形的 TrueType 字体，PDF 测试不依赖本机字体。"""
    pytest.importorskip("fontTools")
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    def box():  # 每个字形都画成同一个方块
        pen = TTGlyphPen(None)
        pen.moveTo((100, 0))
        pen.lineTo((100, 700))
        pen.lineTo((900, 700))
        pen.lineTo((900, 0))
        pen.closePath()
        return pen.glyph()

    chars = sorted(set(CJK_TEXT + string.ascii_letters + string.digits + string.punctuation + " •"))
    names = [".notdef"] + [f"uni{ord(c):04X}" for c in chars]
    fb = FontBuilder(1000, isTTF=True)
    fb.setupGlyphOrder(names)
    fb.setupCharacterMap({ord(c): f"uni{ord(c):04X}" for c in chars})
    fb.setupGlyf({name: box() for name in names})
    fb.setupHorizontalMetrics({name: (1000, 100) for name in names})
    fb.setupHorizontalHeader(ascent=880, descent=-120)
    fb.setupNameTable({"familyName": "AcksTest", "styleName": "Regular", "psName": "AcksTest-Regular"})
    fb.setupOS2(sTypoAscender=880, sTypoDescender=-120, usWinAscent=880, usWinDescent=120, fsType=0)
    fb.setupPost()
    path = tmp_path_factory.mktemp("fonts") / "AcksTest.ttf"
    fb.save(str(path))
    return str(path)


@pytest.fixture
def png(tmp_path):
    """生成 PNG：png(宽, 高, 文件名) -> 路径。"""
    from PIL import Image

    def make(width=120, height=60, name="img.png"):
        path = tmp_path / name
        Image.new("RGB", (width, height), (200, 50, 31)).save(path)
        return path

    return make
