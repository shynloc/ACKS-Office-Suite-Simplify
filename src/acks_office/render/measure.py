"""文字宽度测量：用本机实际选用的字体文件（Pillow / FreeType）测量，测不了时按字符估算。

中文用中文字体量，其余字符用西文字体量，和 Office 的中西文字体分开设置一致。
"""

from functools import lru_cache
from typing import List, Optional

from .. import fonts
from ..themes import FontChoice

_REF = 100  # 以 100 号字测量，再按比例换算


def _is_cjk(ch: str) -> bool:
    return ord(ch) > 0x2E7F


@lru_cache(maxsize=64)
def _font(family: str, bold: bool):
    face = fonts.best_face(family, 700 if bold else 400)
    if face is None:
        return None
    try:
        from PIL import ImageFont
        return ImageFont.truetype(face.path, _REF, index=face.index)
    except Exception:
        return None


def _segments(text: str) -> List[str]:
    parts: List[str] = []
    for ch in text:
        if parts and _is_cjk(parts[-1][-1]) == _is_cjk(ch):
            parts[-1] += ch
        else:
            parts.append(ch)
    return parts


def text_width(text: str, size: float, choice: FontChoice, bold: Optional[bool] = None) -> float:
    """一行文字的宽度（磅）。"""
    bold = choice.bold if bold is None else bold
    total = 0.0
    for part in _segments(text):
        cjk = _is_cjk(part[0])
        font = _font(choice.cn if cjk else choice.en, bold)
        if font is not None:
            total += font.getlength(part) * size / _REF
        else:
            total += size * (1.0 if cjk else 0.56) * len(part) * (1.04 if bold else 1.0)
    return total


def wrap_count(text: str, size: float, width: float, choice: FontChoice, bold: Optional[bool] = None) -> int:
    """按宽度折行后的行数（中文可在任意字之间断行，西文按空格断行）。"""
    lines = 0
    for paragraph in text.split("\n"):
        lines += 1
        current = 0.0
        for token in _tokens(paragraph):
            w = text_width(token, size, choice, bold)
            if current and current + w > width:
                lines += 1
                current = text_width(token.lstrip(), size, choice, bold)
            else:
                current += w
    return lines


def _tokens(text: str) -> List[str]:
    tokens: List[str] = []
    word = ""
    for ch in text:
        if _is_cjk(ch):
            if word:
                tokens.append(word)
                word = ""
            tokens.append(ch)
        elif ch == " ":
            tokens.append(word + " ")
            word = ""
        else:
            word += ch
    if word:
        tokens.append(word)
    return tokens


def fit_size(text: str, size: float, width: float, choice: FontChoice, max_lines: int = 1,
             minimum: Optional[float] = None, bold: Optional[bool] = None) -> float:
    """在最多 max_lines 行内放得下的最大字号（不小于 minimum，默认为原字号的一半）。"""
    floor = minimum if minimum is not None else size * 0.5
    current = size
    while current > floor and wrap_count(text, current, width, choice, bold) > max_lines:
        current = max(floor, current * 0.94)
    return round(current, 1)
