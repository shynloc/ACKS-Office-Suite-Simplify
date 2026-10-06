"""PDF 用的字体：把主题的字体角色落到能嵌入的字体文件上，并把文字按中西文切成 reportlab 段落标记。

PDF 要嵌入真实的字体文件，而 reportlab 只支持 TrueType 轮廓：苹方、Noto Sans CJK 这类 CFF 轮廓的字体不能用。
每个角色分中文、西文两套，依次尝试：主题字体及其备选（fonts.json）→ 各系统常见的 TrueType 字体 →
中文退回 reportlab 内置的 STSong-Light（不嵌入，由阅读器替换显示），西文退回 PDF 标准字体。
可变字体按需要的字重生成静态实例（需要 fontTools），缓存在用户数据目录。
"""

import hashlib
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

from .. import fonts
from ..themes import Theme

# 各系统常见、能嵌入的 TrueType 中文字体，按风格分组
SYSTEM_CN = {
    "sans": ["Heiti SC", "STHeiti", "Microsoft YaHei", "DengXian", "SimHei", "WenQuanYi Zen Hei",
             "WenQuanYi Micro Hei", "Droid Sans Fallback", "Noto Sans SC"],
    "serif": ["Songti SC", "STSong", "SimSun", "NSimSun", "Noto Serif SC", "AR PL UMing CN"],
    "kai": ["Kaiti SC", "STKaiti", "KaiTi", "AR PL UKai CN", "LXGW WenKai"],
}
# PDF 标准字体（阅读器自带，不需要嵌入）：常规、粗体、斜体、粗斜体
STANDARD_EN = {"sans": ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"),
               "serif": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
               "mono": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique")}
CID_FONT = "STSong-Light"
_KAI = ("kai", "楷", "wenkai")
_SERIF = ("serif", "song", "宋", "mincho", "georgia", "times", "fraunces", "garamond")
_MONO = ("mono", "menlo", "consolas", "courier", "code")
# 跟随前一个字的标点：前面是中文就用中文字体（中文排版里的引号、破折号、省略号是全角的）
_NEUTRAL = "“”‘’—…·"
# 两端对齐时一行最多撑开的宽度（字号的倍数），超过就保持左对齐
JUSTIFY_LIMIT = 2.5
# 中文排版不能出现在行首的标点；reportlab 自带的列表按日文整理，缺了全角逗号、分号等
CANNOT_START = "，；：？！》〉”’…—·）】』」〕"


def is_cjk(ch: str) -> bool:
    return ord(ch) > 0x2E7F


def _join_split_lines(blPara, start, stop):
    """拆分中文折行的段落时，组装前后两段的行。

    前半段（start 为 0）沿用 reportlab 的做法：行尾补的空格落在行末、不显示，重新折行时行数不变。
    后半段会按新的宽度重新折行（下一页、下一栏或首字下沉下方），原样拼接各行：中文折行的每一行
    都是原文的连续片段（在西文空格处断行时，空格留在上一行末尾），拼起来就是原文；
    若按 reportlab 的做法在行尾补空格，中文里就会多出空格。
    """
    if start == 0:
        return _REPORTLAB_SPLIT(blPara, start, stop)
    joined = []
    for line in blPara.lines[start:stop]:
        joined.extend(line.words)
    return joined


def _justify_line(tx, offset, line, last=0):
    """两端对齐：含中文的行把多出的宽度平均分到字与字之间（字距），不是只加在少数几个空格上，
    否则一行里只有一两个空格时会被撑出很大的空隙。纯西文的行照旧加大词距。"""
    from reportlab.platypus import paragraph
    words = [w for w in getattr(line, "words", []) if not hasattr(w, "cbDefn")]
    text = "".join(getattr(w, "text", "") for w in words)
    extra = getattr(line, "extraSpace", 0)
    if last or getattr(line, "lineBreak", False) or extra <= 1e-8 or not any(is_cjk(ch) for ch in text):
        return _REPORTLAB_JUSTIFY(tx, offset, line, last)
    trailing = len(text) - len(text.rstrip(" "))
    if trailing and words:  # 行尾的空格不显示：它的宽度也分给字距
        tail = words[-1]
        extra += tx._canvas.stringWidth(" " * trailing, tail.fontName, tail.fontSize)
    gaps = len(text) - trailing - 1
    size = max((getattr(w, "fontSize", 0) for w in words), default=0)
    # 空得太多（段落末行、被长单词挤短的行）就不撑开：字距过大比行尾参差更难看
    if gaps <= 0 or extra > size * JUSTIFY_LIMIT:
        return _REPORTLAB_JUSTIFY(tx, offset, line, last=1)
    tx._x_offset = offset
    paragraph.setXPos(tx, offset)
    tx.setCharSpace(extra / gaps)
    paragraph._putFragLine(offset, tx, line, last, "justify")
    tx.setCharSpace(0)
    paragraph.setXPos(tx, -offset)


def _install_cjk_fixes() -> None:
    """补上 reportlab 中文排版的几处缺口：行首禁则缺全角标点；拆分段落时多出空格；两端对齐只加词距。"""
    global _REPORTLAB_SPLIT, _REPORTLAB_JUSTIFY
    from reportlab.platypus import paragraph
    missing = "".join(ch for ch in CANNOT_START if ch not in paragraph.ALL_CANNOT_START)
    if missing:
        paragraph.ALL_CANNOT_START += missing
    if paragraph._split_blParaHard is not _join_split_lines:
        _REPORTLAB_SPLIT = paragraph._split_blParaHard
        paragraph._split_blParaHard = _join_split_lines
    if paragraph._justifyDrawParaLineX is not _justify_line:
        _REPORTLAB_JUSTIFY = paragraph._justifyDrawParaLineX
        paragraph._justifyDrawParaLineX = _justify_line


_REPORTLAB_SPLIT = _REPORTLAB_JUSTIFY = None
_install_cjk_fixes()


def script_runs(text: str) -> List[Tuple[str, str]]:
    """按书写系统切分：[("cn" | "en", 文字)]。引号、破折号等跟随前一个字（在开头时跟随后一个字）。"""
    runs: List[Tuple[str, str]] = []
    for i, ch in enumerate(text):
        if ch in _NEUTRAL:
            following = next((c for c in text[i + 1:] if c not in _NEUTRAL), "中")
            kind = runs[-1][0] if runs else ("cn" if is_cjk(following) else "en")
        else:
            kind = "cn" if is_cjk(ch) else "en"
        if runs and runs[-1][0] == kind:
            runs[-1] = (kind, runs[-1][1] + ch)
        else:
            runs.append((kind, ch))
    return runs


def _category(names: List[str], script: str) -> str:
    text = " ".join(names).lower()
    if any(k in text for k in _KAI):
        return "kai"
    if script == "en" and any(k in text for k in _MONO):
        return "mono"
    if any(k in text for k in _SERIF):
        return "serif"
    return "sans"


@dataclass
class PdfFace:
    name: str       # reportlab 里的字体名
    family: str     # 实际用到的字体家族
    embedded: bool  # 是否嵌入 PDF


class PdfFontBook:
    def __init__(self, theme: Theme):
        self.t = theme
        self.warnings: List[Dict[str, str]] = []
        self._faces: Dict[Tuple[str, str, int, bool], PdfFace] = {}
        self._noted: set = set()
        self._override: Optional[Tuple[PdfFace, PdfFace]] = None

    def override_cn(self, regular: str, bold: Optional[str] = None) -> None:
        """所有角色的中文都改用指定的 TrueType 字体文件（粗体可另给一个文件）。"""
        names = []
        for path in (regular, bold or regular):
            name = fonts._register_ttf(path)
            if not name:
                raise ValueError(f"字体无法用于 PDF：{path}（需要包含中文字形的 TrueType .ttf/.ttc 字体）")
            names.append(PdfFace(name, os.path.splitext(os.path.basename(path))[0], True))
        self._override = (names[0], names[1])
        self._faces = {key: face for key, face in self._faces.items() if key[1] != "cn"}

    # ------------------------------------------------------------ 选字体

    def weight(self, role: str, bold: Optional[bool] = None) -> int:
        weight = int(self.t.get(f"font.roles.{role}").get("weight", 400))
        if bold is True:
            return max(weight, 700)
        if bold is False and weight >= 600:
            return 400
        return weight

    def face(self, role: str, script: str, bold: Optional[bool] = None, italic: bool = False) -> PdfFace:
        slot = self.t.get(f"font.roles.{role}")["family"]
        return self._resolve(slot, script, self.weight(role, bold), italic and script == "en")

    def _resolve(self, slot: str, script: str, weight: int, italic: bool) -> PdfFace:
        key = (slot, script, weight, italic)
        if key in self._faces:
            return self._faces[key]
        if script == "cn" and self._override:
            return self._override[1 if weight >= 600 else 0]
        names = self.t.family_candidates(slot, script)
        category = _category(names, script)
        tried = list(names)
        if script == "cn":
            for group in (category, "sans", "serif", "kai"):  # 同风格的系统字体优先
                tried += [n for n in SYSTEM_CN[group] if n not in tried]
        result = None
        for family in tried:
            face = fonts.best_face(family, weight, italic, truetype=True)
            source = self._static(face, weight) if face else None
            name = fonts._register_ttf(*source, require_cjk=script == "cn") if source else None
            if name:
                result = PdfFace(name, family, True)
                break
        if result is None:
            result = self._fallback(script, category, weight, italic)
        self._note(names[0], script, result)
        self._faces[key] = result
        return result

    def _static(self, face: "fonts.FontFace", weight: int) -> Optional[Tuple[str, int]]:
        """可变字体按字重生成静态实例（reportlab 只会用默认实例）；没有 fontTools 时跳过这个字体。"""
        if not face.variable:
            return face.path, face.index
        try:
            from fontTools.ttLib import TTFont
            from fontTools.varLib import instancer
        except ImportError:
            return None
        stat = os.stat(face.path)
        digest = hashlib.sha1(f"{face.path}#{face.index}#{stat.st_size}#{stat.st_mtime_ns}#{weight}"
                              .encode()).hexdigest()[:16]
        target = fonts.data_dir() / "cache" / "pdf-fonts" / f"{digest}.ttf"
        if not target.is_file():
            try:
                font = TTFont(face.path, fontNumber=face.index)
                location = {}
                for axis in font["fvar"].axes:  # 字重取最接近的值，其余轴取默认值，得到完全静态的字体
                    value = weight if axis.axisTag == "wght" else axis.defaultValue
                    location[axis.axisTag] = min(max(value, axis.minValue), axis.maxValue)
                static = instancer.instantiateVariableFont(font, location)
                target.parent.mkdir(parents=True, exist_ok=True)
                tmp = target.with_suffix(".tmp")
                static.save(str(tmp))
                os.replace(tmp, target)
            except Exception:
                return None
        return str(target), 0

    def _fallback(self, script: str, category: str, weight: int, italic: bool) -> PdfFace:
        from reportlab.pdfbase import pdfmetrics
        if script == "cn":
            if CID_FONT not in pdfmetrics.getRegisteredFontNames():
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont
                pdfmetrics.registerFont(UnicodeCIDFont(CID_FONT))
            return PdfFace(CID_FONT, CID_FONT, False)
        names = STANDARD_EN["serif" if category in ("serif", "kai") else category]
        return PdfFace(names[(2 if italic else 0) + (1 if weight >= 600 else 0)], names[0], False)

    def _note(self, primary: str, script: str, face: PdfFace) -> None:
        if face.family == primary or (primary, script) in self._noted:
            return
        self._noted.add((primary, script))
        install = self.t.fonts.get("families", {}).get(primary, {}).get("install", "")
        if script == "cn" and not face.embedded:
            message = (f"找不到能嵌入 PDF 的中文字体（{primary} 未安装，或不是 TrueType 轮廓），"
                       f"中文改用阅读器自带的 {CID_FONT}，不同设备上显示可能不同")
            note = {"code": "FONT_NOT_EMBEDDED", "family": primary, "used": face.family, "message": message}
        else:
            note = {"code": "FONT_SUBSTITUTED", "family": primary, "used": face.family,
                    "message": f"主题字体 {primary} 不能用于 PDF（未安装，或不是 TrueType 轮廓），已改用 {face.family}"}
        if install and install in fonts.CATALOG:
            note["install"] = install
            note["message"] += f"；可运行 fonts install {install} 安装"
        self.warnings.append(note)

    # ------------------------------------------------------------ 段落标记

    def markup(self, text: str, role: str, bold: Optional[bool] = None, italic: bool = False,
               color: Optional[str] = None, size: Optional[float] = None, breaks: bool = True) -> str:
        """一段文字 → reportlab 段落标记：中文、西文各用各的字体；color 为 #RRGGBB。
        breaks=False 时保留原样的换行（给保留空白的 XPreformatted 用）。"""
        attrs = (f' color="{color}"' if color else "") + (f' size="{size:g}"' if size else "")
        parts = []
        for script, chunk in script_runs(text):
            face = self.face(role, script, bold, italic)
            body = escape(chunk)
            if breaks:
                body = body.replace("\n", "<br/>")
            parts.append(f'<font name="{face.name}"{attrs}>{body}</font>')
        return "".join(parts)

    def width(self, text: str, role: str, size: float, bold: Optional[bool] = None) -> float:
        """一行文字的宽度（磅）。"""
        from reportlab.pdfbase.pdfmetrics import stringWidth
        return sum(stringWidth(chunk, self.face(role, script, bold).name, size)
                   for script, chunk in script_runs(text))

    def embedded_fonts(self) -> List[str]:
        return sorted({face.family for face in self._faces.values() if face.embedded})

