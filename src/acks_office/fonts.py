"""字体：为 PDF 找到能嵌入的中文字体，检查主题字体是否已安装，并按需下载开源字体。

这个模块只依赖标准库；reportlab、fontTools 在用到时才导入，所以 doctor 在依赖不全时也能运行。
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from .utils import data_dir

CID_FALLBACK = "STSong-Light"
_FONT_EXTENSIONS = (".ttf", ".ttc", ".otf", ".otc")

_GOOGLE_FONTS = "https://raw.githubusercontent.com/google/fonts/2894aab31764f10f29c421bdfd2340d3b382d384/ofl/"
_LXGW = "https://github.com/lxgw/LxgwWenKai/releases/download/v1.522/"
_OFL = "SIL Open Font License 1.1"

# 可以按需下载的开源字体（都是 SIL OFL 1.1，可免费商用）。地址固定到具体提交或发布版本，下载后校验哈希：
# sha256，或该文件在 Git 仓库里的 blob SHA-1（git_sha1，随固定的提交一起确定内容）。
# 可变字体按 instances 生成固定字重的 TrueType 文件（需要 fontTools）；static 为直接使用的静态字体。
CATALOG: Dict[str, Dict] = {
    "noto-sans-sc": {
        "family": "Noto Sans SC", "license": _OFL, "file_stem": "NotoSansSC",
        "license_url": _GOOGLE_FONTS + "notosanssc/OFL.txt",
        "sources": [{"url": _GOOGLE_FONTS + "notosanssc/NotoSansSC%5Bwght%5D.ttf", "size": 17772300,
                     "sha256": "a3041811a78c361b1de50f953c805e0244951c21c5bd412f7232ef0d899af0da",
                     "instances": {"Regular": {"wght": 400}, "Medium": {"wght": 500}, "Bold": {"wght": 700}}}],
    },
    "noto-serif-sc": {
        "family": "Noto Serif SC", "license": _OFL, "file_stem": "NotoSerifSC",
        "license_url": _GOOGLE_FONTS + "notoserifsc/OFL.txt",
        "sources": [{"url": _GOOGLE_FONTS + "notoserifsc/NotoSerifSC%5Bwght%5D.ttf", "size": 59925648,
                     "git_sha1": "e41956d22ba6f7b7126277a5eee4dfa2129075bc",
                     "instances": {"Regular": {"wght": 400}, "Bold": {"wght": 700}, "Black": {"wght": 900}}}],
    },
    "source-sans-3": {
        "family": "Source Sans 3", "license": _OFL, "file_stem": "SourceSans3",
        "license_url": _GOOGLE_FONTS + "sourcesans3/OFL.txt",
        "sources": [{"url": _GOOGLE_FONTS + "sourcesans3/SourceSans3%5Bwght%5D.ttf", "size": 652632,
                     "git_sha1": "c578acd4067fc58acfa8ec09a79e2ae2a8d933f8",
                     "instances": {"Regular": {"wght": 400}, "Medium": {"wght": 500}, "SemiBold": {"wght": 600},
                                   "Bold": {"wght": 700}}},
                    {"url": _GOOGLE_FONTS + "sourcesans3/SourceSans3-Italic%5Bwght%5D.ttf", "size": 383592,
                     "git_sha1": "30896f2e951284e92583e55948493a4067d6baf0",
                     "instances": {"Italic": {"wght": 400}, "BoldItalic": {"wght": 700}}}],
    },
    "source-serif-4": {
        "family": "Source Serif 4", "license": _OFL, "file_stem": "SourceSerif4",
        "license_url": "https://raw.githubusercontent.com/adobe-fonts/source-serif/4.005R/LICENSE.md",
        "sources": [{"url": _GOOGLE_FONTS + "sourceserif4/SourceSerif4%5Bopsz%2Cwght%5D.ttf", "size": 1209508,
                     "git_sha1": "40bbb58cb58e893ad0470f0bb291e096c73d3554",
                     "instances": {"Regular": {"wght": 400, "opsz": 12}, "Bold": {"wght": 700, "opsz": 12}}},
                    {"url": _GOOGLE_FONTS + "sourceserif4/SourceSerif4-Italic%5Bopsz%2Cwght%5D.ttf", "size": 855432,
                     "git_sha1": "ccb969aa6847e1169ff9c8c8bdc28daff6966861",
                     "instances": {"Italic": {"wght": 400, "opsz": 12}}}],
    },
    "fraunces": {  # 用在标题和大号数字：光学尺寸取最大
        "family": "Fraunces", "license": _OFL, "file_stem": "Fraunces",
        "license_url": _GOOGLE_FONTS + "fraunces/OFL.txt",
        "sources": [{"url": _GOOGLE_FONTS + "fraunces/Fraunces%5BSOFT%2CWONK%2Copsz%2Cwght%5D.ttf", "size": 360440,
                     "git_sha1": "8210f9488d3c732359a9292dd09aca3f2bae830e",
                     "instances": {"Light": {"wght": 300, "opsz": 144}, "Regular": {"wght": 400, "opsz": 144},
                                   "Medium": {"wght": 500, "opsz": 144}, "Bold": {"wght": 700, "opsz": 144}}},
                    {"url": _GOOGLE_FONTS + "fraunces/Fraunces-Italic%5BSOFT%2CWONK%2Copsz%2Cwght%5D.ttf",
                     "size": 414904, "git_sha1": "2ddf59a11b89a3d6b2816ea6ae7965026c08143a",
                     "instances": {"Italic": {"wght": 400, "opsz": 144}}}],
    },
    "lxgw-wenkai": {
        "family": "LXGW WenKai", "license": _OFL, "file_stem": "LXGWWenKai",
        "license_url": "https://raw.githubusercontent.com/lxgw/LxgwWenKai/v1.522/OFL.txt",
        "sources": [{"url": _LXGW + "LXGWWenKai-Regular.ttf", "size": 25575676, "static": "Regular",
                     "sha256": "39ad71264b588165b469e35e6afb162a378dacd1f95348160240ba9038ac3009"},
                    {"url": _LXGW + "LXGWWenKai-Medium.ttf", "size": 25379848, "static": "Medium",
                     "sha256": "d4bdeb38a39151d74d084cba5090f8cb7d20bf83eedb78c35939ae70b9f4e3f6"}],
    },
    "jetbrains-mono": {
        "family": "JetBrains Mono", "license": _OFL, "file_stem": "JetBrainsMono",
        "license_url": _GOOGLE_FONTS + "jetbrainsmono/OFL.txt",
        "sources": [{"url": _GOOGLE_FONTS + "jetbrainsmono/JetBrainsMono%5Bwght%5D.ttf", "size": 187208,
                     "git_sha1": "aa310be8b717fe3774f9444dd89d5f4101cc6d10",
                     "instances": {"Regular": {"wght": 400}, "Bold": {"wght": 700}}}],
    },
}


@dataclass
class PdfFonts:
    """PDF 用的中文字体：reportlab 注册名、来源，以及是否会嵌入 PDF。"""
    regular: str
    bold: str
    family: str
    path: Optional[str]
    embedded: bool
    source: str

    def describe(self) -> dict:
        return asdict(self)


def font_dirs() -> List[Path]:
    """系统与用户字体目录，外加本工具下载字体的目录。"""
    home = Path.home()
    if sys.platform == "darwin":
        dirs = [Path("/System/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
                Path("/Library/Fonts"), home / "Library" / "Fonts"]
        # 苹方等系统字体在新版 macOS 里按需下载，放在资源目录里
        dirs += sorted(Path("/System/Library/AssetsV2").glob("com_apple_MobileAsset_Font*"))
    elif os.name == "nt":
        dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
        if os.environ.get("LOCALAPPDATA"):
            dirs.append(Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts")
    else:
        dirs = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                home / ".local" / "share" / "fonts", home / ".fonts"]
    return dirs + [data_dir() / "fonts"]


def _font_files(limit: int = 20000) -> List[Path]:
    files: List[Path] = []
    for root in font_dirs():
        if not root.is_dir():
            continue
        for dirpath, _, names in os.walk(root):
            for name in names:
                if name.lower().endswith(_FONT_EXTENSIONS):
                    files.append(Path(dirpath) / name)
                    if len(files) >= limit:
                        return files
    return files


# ---------------------------------------------------------------- 读取字体内部名称

@dataclass(frozen=True)
class FontFace:
    """字体文件里的一款字形：家族名（含本地化名称）、字重、是否斜体、轮廓类型。"""
    path: str
    index: int
    families: Tuple[str, ...]
    style: str
    weight: int
    italic: bool
    truetype: bool  # glyf 轮廓才能被 reportlab 嵌入；CFF 轮廓不能
    variable: bool
    localized: Tuple[Tuple[str, str], ...] = ()  # (语言, 家族名)：("en", "PingFang SC")、("zh-Hans", "苹方-简")


def _u16(data: bytes, offset: int) -> int:
    return int.from_bytes(data[offset:offset + 2], "big")


def _u32(data: bytes, offset: int) -> int:
    return int.from_bytes(data[offset:offset + 4], "big")


_CONTROL = re.compile(r"[\x00-\x1f\x7f\ufffe\uffff\ud800-\udfff]")
# name 表的语言编号 → 语言标签（只关心中日韩的本地化家族名）
_WIN_LANGS = {0x0804: "zh-Hans", 0x1004: "zh-Hans", 0x0404: "zh-Hant", 0x0C04: "zh-Hant", 0x1404: "zh-Hant",
              0x0411: "ja", 0x0412: "ko"}
_MAC_LANGS = {33: "zh-Hans", 19: "zh-Hant", 11: "ja", 23: "ko"}


def _name_strings(table: bytes, wanted: Tuple[int, ...]) -> Dict[int, List[Tuple[str, str]]]:
    """name 表里指定编号的全部字符串 [(语言标签, 文字)]，英文排在前面；语言不关心时标签为空。"""
    found: Dict[int, List[Tuple[int, str, str]]] = {i: [] for i in wanted}
    count, base = _u16(table, 2), _u16(table, 4)
    for i in range(count):
        rec = 6 + 12 * i
        platform, encoding, language, name_id = (_u16(table, rec), _u16(table, rec + 2),
                                                 _u16(table, rec + 4), _u16(table, rec + 6))
        if name_id not in found:
            continue
        raw = table[base + _u16(table, rec + 10): base + _u16(table, rec + 10) + _u16(table, rec + 8)]
        if platform in (0, 3):
            text = raw.decode("utf-16-be", "ignore")
        elif platform == 1 and encoding == 0:
            text = raw.decode("mac_roman", "ignore")
        else:
            continue
        text = _CONTROL.sub("", text)  # 有些字体的名称记录里混有 NUL 等控制字符
        english = (platform == 3 and language == 0x409) or (platform == 1 and language == 0)
        tag = "en" if english else (_WIN_LANGS.get(language, "") if platform == 3
                                     else _MAC_LANGS.get(language, "") if platform == 1 else "")
        if text.strip():
            found[name_id].append((0 if english else 1, tag, text.strip()))
    return {k: list(dict.fromkeys((tag, t) for _, tag, t in sorted(v, key=lambda x: x[0])))
            for k, v in found.items()}


def _read_face(f, offset: int, path: str, index: int) -> Optional[FontFace]:
    f.seek(offset)
    head = f.read(12)
    if len(head) < 12:
        return None
    count = _u16(head, 4)
    records = f.read(16 * count)
    tables = {records[i * 16:i * 16 + 4]: (_u32(records, i * 16 + 8), _u32(records, i * 16 + 12))
              for i in range(count)}
    if b"name" not in tables:
        return None
    f.seek(tables[b"name"][0])
    names = _name_strings(f.read(tables[b"name"][1]), (1, 2, 16, 17))
    weight, italic = 400, False
    if b"OS/2" in tables:
        f.seek(tables[b"OS/2"][0])
        os2 = f.read(64)
        if len(os2) >= 64:
            weight, italic = _u16(os2, 4), bool(_u16(os2, 62) & 1)
    family_names = names[16] + names[1]
    families = tuple(dict.fromkeys(t for _, t in family_names))
    if not families:
        return None
    localized = tuple(dict.fromkeys((tag, t) for tag, t in family_names if tag))
    style = (names[17] or names[2] or [("", "Regular")])[0][1]
    return FontFace(path, index, families, style, weight, italic or "italic" in style.lower(),
                    b"glyf" in tables, b"fvar" in tables, localized)


def read_font_faces(path: str) -> List[FontFace]:
    """读取一个字体文件（.ttf / .otf / .ttc）里的全部字形信息；读不了的文件返回空列表。"""
    try:
        with open(path, "rb") as f:
            head = f.read(12)
            if head[:4] == b"ttcf":
                offsets = f.read(4 * _u32(head, 8))
                return [face for i in range(len(offsets) // 4)
                        if (face := _read_face(f, _u32(offsets, 4 * i), path, i))]
            face = _read_face(f, 0, path, 0)
            return [face] if face else []
    except (OSError, ValueError, IndexError):
        return []


_FACES: Optional[List[FontFace]] = None
_CACHE_VERSION = 4


def _cache_path() -> Path:
    return data_dir() / "cache" / "font-faces.json"


def font_faces(refresh: bool = False, write_cache: bool = True) -> List[FontFace]:
    """本机全部字体（系统、用户目录与本工具下载的），同一进程内只扫描一次。

    读过的文件按路径、大小、修改时间缓存在用户数据目录；write_cache=False 时只读缓存
    （doctor 用，保证只读）。
    """
    global _FACES
    if _FACES is not None and not refresh:
        return _FACES
    try:
        cached = json.loads(_cache_path().read_text(encoding="utf-8"))
        if cached.get("version") != _CACHE_VERSION:
            cached = {}
    except (OSError, ValueError):
        cached = {}
    entries, faces, changed = {}, [], False
    for path in _font_files():
        try:
            stat = path.stat()
        except OSError:
            continue
        stamp = [stat.st_size, int(stat.st_mtime)]
        hit = cached.get("files", {}).get(str(path))
        if hit and hit["stamp"] == stamp:
            found = [FontFace(**dict(f, families=tuple(f["families"]),
                                     localized=tuple(tuple(x) for x in f["localized"]))) for f in hit["faces"]]
        else:
            found, changed = read_font_faces(str(path)), True
        entries[str(path)] = {"stamp": stamp, "faces": [asdict(f) for f in found]}
        faces.extend(found)
    if write_cache and (changed or len(entries) != len(cached.get("files", {}))):
        try:
            _cache_path().parent.mkdir(parents=True, exist_ok=True)
            _cache_path().write_text(json.dumps({"version": _CACHE_VERSION, "files": entries},
                                                ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass  # 缓存只是加速，写不了不影响结果
    _FACES = faces
    return faces


def _key(name: str) -> str:
    return re.sub(r"[\s_-]+", "", name).lower()


def faces_of(family: str) -> List[FontFace]:
    """某个字体家族在本机的全部字形（名称不区分大小写与空格，本地化名称也算）。"""
    key = _key(family)
    return [face for face in font_faces() if any(_key(name) == key for name in face.families)]


def family_installed(family: str) -> bool:
    return bool(faces_of(family))


def best_face(family: str, weight: int = 400, italic: bool = False,
              truetype: bool = False) -> Optional[FontFace]:
    """在某个家族里挑最接近要求的字形；truetype=True 时只要 reportlab 能嵌入的。"""
    faces = [face for face in faces_of(family) if face.truetype or not truetype]
    if not faces:
        return None
    return min(faces, key=lambda face: (face.italic != italic, face.variable,
                                        abs(face.weight - weight), face.path))


def _cmap_codepoints(table: bytes) -> set:
    """cmap 表里能映射到字形的码位（只读 format 4 与 12，覆盖常见字体）。"""
    best = None
    for i in range(_u16(table, 2)):
        rec = 4 + 8 * i
        platform, encoding, offset = _u16(table, rec), _u16(table, rec + 2), _u32(table, rec + 4)
        fmt = _u16(table, offset)
        # 与 fontTools 的 getBestCmap 一致：完整 Unicode 优先，同类时 Windows 子表优先
        windows = platform == 3 and encoding in (1, 10)
        if fmt not in (4, 12) or not (windows or platform == 0):
            continue
        rank = (0 if fmt == 12 else 2) + (0 if windows else 1)
        if best is None or rank < best[0]:
            best = (rank, offset, fmt)
    points: set = set()
    if best is None:
        return points
    _, offset, fmt = best
    if fmt == 4:
        segs = _u16(table, offset + 6) // 2
        ends, starts = offset + 14, offset + 16 + 2 * segs
        deltas, ranges = starts + 2 * segs, starts + 4 * segs
        for s in range(segs):
            end, start = _u16(table, ends + 2 * s), _u16(table, starts + 2 * s)
            delta, range_offset = _u16(table, deltas + 2 * s), _u16(table, ranges + 2 * s)
            for cp in range(start, min(end, 0xFFFE) + 1):
                if range_offset == 0:
                    glyph = (cp + delta) & 0xFFFF
                else:
                    at = ranges + 2 * s + range_offset + 2 * (cp - start)
                    glyph = _u16(table, at)
                    glyph = (glyph + delta) & 0xFFFF if glyph else 0
                if glyph:
                    points.add(cp)
    elif fmt == 12:
        for g in range(_u32(table, offset + 12)):
            rec = offset + 16 + 12 * g
            start = _u32(table, rec) + (1 if _u32(table, rec + 8) == 0 else 0)  # 跳过映射到 .notdef 的码位
            points.update(range(start, _u32(table, rec + 4) + 1))
    return points


def face_codepoints(face: FontFace) -> set:
    """某款字形支持的字符码位。读不了时返回空集合。"""
    try:
        with open(face.path, "rb") as f:
            head = f.read(12)
            offset = 0
            if head[:4] == b"ttcf":
                f.seek(12 + 4 * face.index)
                offset = _u32(f.read(4), 0)
            f.seek(offset)
            count = _u16(f.read(12), 4)
            records = f.read(16 * count)
            for i in range(count):
                if records[i * 16:i * 16 + 4] == b"cmap":
                    f.seek(_u32(records, i * 16 + 8))
                    return _cmap_codepoints(f.read(_u32(records, i * 16 + 12)))
    except (OSError, ValueError, IndexError):
        pass
    return set()


def common_hanzi() -> List[str]:
    """GB2312 一级字（3755 个常用汉字），用来检查中文字体的覆盖率。"""
    chars = []
    for hi in range(0xB0, 0xD8):
        for lo in range(0xA1, 0xFF):
            try:
                chars.append(bytes([hi, lo]).decode("gb2312"))
            except UnicodeDecodeError:
                continue
    return chars


def mac_ui_language() -> str:
    """macOS 的首选语言（如 zh-Hans-CN）；读不到或不是 macOS 时返回空字符串。"""
    if sys.platform != "darwin":
        return ""
    try:
        import plistlib
        with open(Path.home() / "Library" / "Preferences" / ".GlobalPreferences.plist", "rb") as f:
            languages = plistlib.load(f).get("AppleLanguages") or []
        return str(languages[0]) if languages else ""
    except Exception:
        return ""


def _language_tag(ui_language: str) -> str:
    lang = ui_language.lower()
    if lang.startswith(("zh-hant", "zh-tw", "zh-hk", "zh-mo")):
        return "zh-Hant"
    if lang.startswith("zh"):
        return "zh-Hans"
    return {"ja": "ja", "ko": "ko"}.get(lang[:2], "")


def localized_family_names(ui_language: Optional[str] = None) -> Dict[str, str]:
    """英文家族名 → 当前系统语言下的本地化家族名（如 PingFang SC → 苹方-简）。

    macOS 上的 LibreOffice 只认本地化后的名字，转换前用它生成字体替换表。
    """
    tag = _language_tag(mac_ui_language() if ui_language is None else ui_language)
    if not tag:
        return {}
    pairs: Dict[str, str] = {}
    for face in font_faces(write_cache=False):
        local = next((name for lang, name in face.localized if lang == tag), None)
        if not local:
            continue
        for lang, name in face.localized:
            if lang == "en" and name.isascii() and name != local and name not in pairs:
                pairs[name] = local
    return pairs


def theme_font_status(theme: Any = None) -> Dict[str, Any]:
    """某个主题用到的字体在本机的情况（不传则为默认主题）。

    present：主题字体已安装；substituted：没装主题字体，用的是备选字体 {主题字体: 备选}；
    missing：主题字体和备选都没装；install：可以用 fonts install 下载的字体。
    """
    from .themes import load_theme
    t = load_theme(theme)
    families = t.fonts.get("families", {})
    present, substituted, missing, install = set(), {}, set(), set()
    for slot in sorted({spec["family"] for spec in t.get("font.roles").values()}):
        for script in ("cn", "en"):
            names = t.family_candidates(slot, script)
            primary = names[0]
            if family_installed(primary):
                present.add(primary)
                continue
            used = next((n for n in names[1:] if family_installed(n)), None)
            if used:
                substituted[primary] = used
            else:
                missing.add(primary)
            key = families.get(primary, {}).get("install")
            if key in CATALOG:
                install.add(key)
    return {"theme": t.name, "present": sorted(present), "substituted": substituted,
            "missing": sorted(missing), "install": sorted(install)}


# ---------------------------------------------------------------- PDF 中文字体

def _linux_cjk_candidates() -> List[Tuple[Tuple[str, int], Optional[Tuple[str, int]], str]]:
    wanted = {"wqy-microhei.ttc": "WenQuanYi Micro Hei", "wqy-zenhei.ttc": "WenQuanYi Zen Hei",
              "droidsansfallbackfull.ttf": "Droid Sans Fallback"}
    found = []
    for path in _font_files():
        family = wanted.get(path.name.lower())
        if family:
            found.append(((str(path), 0), None, family))
    return found


def _system_cjk_candidates() -> List[Tuple[Tuple[str, int], Optional[Tuple[str, int]], str]]:
    """各系统自带、且是 TrueType 轮廓（reportlab 能嵌入）的中文字体：(常规, 粗体, 名称)。"""
    if sys.platform == "darwin":
        return [
            (("/System/Library/Fonts/STHeiti Light.ttc", 1), ("/System/Library/Fonts/STHeiti Medium.ttc", 1), "STHeiti SC"),
            (("/Library/Fonts/Arial Unicode.ttf", 0), None, "Arial Unicode MS"),
            (("/System/Library/Fonts/Supplemental/Arial Unicode.ttf", 0), None, "Arial Unicode MS"),
        ]
    if os.name == "nt":
        fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        return [
            ((os.path.join(fonts, "msyh.ttc"), 0), (os.path.join(fonts, "msyhbd.ttc"), 0), "Microsoft YaHei"),
            ((os.path.join(fonts, "Deng.ttf"), 0), (os.path.join(fonts, "Dengb.ttf"), 0), "DengXian"),
            ((os.path.join(fonts, "simhei.ttf"), 0), None, "SimHei"),
            ((os.path.join(fonts, "simsun.ttc"), 0), None, "SimSun"),
        ]
    return _linux_cjk_candidates()


_REGISTERED: Dict[Tuple[str, int], Tuple[Optional[str], bool]] = {}


def _register_ttf(path: str, index: int = 0, require_cjk: bool = True) -> Optional[str]:
    """把一个 TrueType 字体注册给 reportlab，返回注册名；不能用（CFF 轮廓、文件损坏，
    或 require_cjk 时缺中文字形）返回 None。"""
    key = (os.path.abspath(path), index)
    if key not in _REGISTERED:
        name, has_cjk = None, False
        if os.path.isfile(path):
            try:
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                name = "acks-" + hashlib.sha1(f"{key[0]}#{index}".encode()).hexdigest()[:10]
                font = TTFont(name, path, subfontIndex=index)
                has_cjk = {0x4E2D, 0x6587} <= set(font.face.charToGlyph)  # 「中」「文」
                pdfmetrics.registerFont(font)
            except Exception:
                name = None
        _REGISTERED[key] = (name, has_cjk)
    name, has_cjk = _REGISTERED[key]
    return name if (has_cjk or not require_cjk) else None


def _family(regular: str, bold: str) -> None:
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily(regular, normal=regular, bold=bold, italic=regular, boldItalic=bold)


def _installed_catalog_fonts() -> List[Tuple[Tuple[str, int], Optional[Tuple[str, int]], str]]:
    found = []
    base = data_dir() / "fonts"
    for spec in CATALOG.values():
        regular = base / f"{spec['file_stem']}-Regular.ttf"
        bold = base / f"{spec['file_stem']}-Bold.ttf"
        if regular.is_file():
            found.append(((str(regular), 0), (str(bold), 0) if bold.is_file() else None, spec["family"]))
    return found


def pdf_fonts(font_path: Optional[str] = None, bold_font_path: Optional[str] = None) -> PdfFonts:
    """选出 PDF 用的中文字体。

    顺序：参数指定 → 环境变量 ACKS_OFFICE_PDF_FONT(_BOLD) → `acks-office fonts install` 下载的字体
    → 系统自带的中文 TrueType 字体 → reportlab 内置的 STSong-Light（不嵌入，由阅读器替换显示）。
    """
    explicit = font_path or os.environ.get("ACKS_OFFICE_PDF_FONT")
    if explicit:
        regular = _register_ttf(explicit)
        if not regular:
            raise ValueError(f"字体无法用于 PDF：{explicit}（需要包含中文字形的 TrueType .ttf/.ttc 字体）")
        bold_path = bold_font_path or os.environ.get("ACKS_OFFICE_PDF_FONT_BOLD")
        bold = _register_ttf(bold_path) if bold_path else None
        _family(regular, bold or regular)
        return PdfFonts(regular, bold or regular, Path(explicit).stem, explicit, True,
                        "argument" if font_path else "env")

    for source, candidates in (("installed", _installed_catalog_fonts()), ("system", _system_cjk_candidates())):
        for (path, index), bold_spec, family in candidates:
            regular = _register_ttf(path, index)
            if not regular:
                continue
            bold = _register_ttf(*bold_spec) if bold_spec else None
            _family(regular, bold or regular)
            return PdfFonts(regular, bold or regular, family, path, True, source)

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    if CID_FALLBACK not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(CID_FALLBACK))
    _family(CID_FALLBACK, CID_FALLBACK)
    return PdfFonts(CID_FALLBACK, CID_FALLBACK, CID_FALLBACK, None, False, "builtin")


# ---------------------------------------------------------------- 下载开源字体

def _download(url: str, dest: Path, progress: Optional[Callable[[str], None]] = None,
              size: Optional[int] = None) -> Dict[str, str]:
    """下载到 dest，返回内容的 sha256，以及按 Git blob 算的 SHA-1（给出 size 时）。"""
    mirror = os.environ.get("ACKS_OFFICE_DOWNLOAD_MIRROR")
    if mirror:
        url = url.replace("https://raw.githubusercontent.com", mirror.rstrip("/"))
    request = urllib.request.Request(url, headers={"User-Agent": "acks-office"})
    sha256 = hashlib.sha256()
    git = hashlib.sha1(b"blob %d\x00" % size) if size is not None else None
    with urllib.request.urlopen(request, timeout=60) as response, open(dest, "wb") as out:
        total = int(response.headers.get("Content-Length") or 0) or (size or 0)
        done = 0
        while True:
            chunk = response.read(1 << 20)
            if not chunk:
                break
            out.write(chunk)
            sha256.update(chunk)
            if git is not None:
                git.update(chunk)
            done += len(chunk)
            if progress and total:
                progress(f"下载 {done * 100 // total}%")
    return {"sha256": sha256.hexdigest(), "git_sha1": git.hexdigest() if git is not None else ""}


def _verified(digest: Dict[str, str], source: Dict) -> bool:
    if source.get("sha256"):
        return digest["sha256"] == source["sha256"]
    return bool(source.get("git_sha1")) and digest["git_sha1"] == source["git_sha1"]


_STYLE_NAMES = {"BoldItalic": "Bold Italic"}


def _static_instance(path: Path, location: Dict[str, float], family: str, style: str):
    """可变字体 → 某个字重的静态字体：所有轴都固定（没给的取默认值），并改好名称。"""
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer
    font = TTFont(str(path))
    pinned = {}
    for axis in font["fvar"].axes:
        value = location.get(axis.axisTag, axis.defaultValue)
        pinned[axis.axisTag] = min(max(value, axis.minValue), axis.maxValue)
    static = instancer.instantiateVariableFont(font, pinned)
    _rename(static, family, style, int(pinned.get("wght", 400)))
    return static


def _rename(font, family: str, style: str, weight: int) -> None:
    """静态实例的名称与字重。Word、PowerPoint 按「家族 + 样式」区分字体，几个字重同名会互相覆盖：
    常规、粗体、斜体、粗斜体共用家族名；其余字重（Medium、Black 等）另起家族名，并写上排版用的家族与样式。"""
    italic = style.endswith("Italic")
    bold = style in ("Bold", "BoldItalic")
    display = _STYLE_NAMES.get(style, style)
    if style in ("Regular", "Bold", "Italic", "BoldItalic"):
        legacy_family, legacy_style = family, display
    else:
        base = display.replace("Italic", "").strip()
        legacy_family, legacy_style = f"{family} {base}", "Italic" if italic else "Regular"
    postscript = f"{family.replace(' ', '')}-{style}"
    names = {1: legacy_family, 2: legacy_style, 3: f"acks-office:{postscript}", 4: f"{family} {display}",
             6: postscript, 16: family, 17: display}
    table = font["name"]
    table.names = [record for record in table.names if record.nameID not in (1, 2, 3, 4, 6, 16, 17, 25)]
    for name_id, text in names.items():
        table.setName(text, name_id, 3, 1, 0x409)
        table.setName(text, name_id, 1, 0, 0)
    os2 = font["OS/2"]
    os2.usWeightClass = weight
    selection = os2.fsSelection & ~(1 | 1 << 5 | 1 << 6)  # 斜体、粗体、常规三个标志位按样式重设
    os2.fsSelection = selection | (1 if italic else 0) | (1 << 5 if bold else 0) | (0 if bold or italic else 1 << 6)
    font["head"].macStyle = (1 if bold else 0) | (2 if italic else 0)


def system_font_dir() -> Path:
    """当前用户的字体目录：装在这里，Word、PowerPoint、WPS 等软件都能用，不需要管理员权限。"""
    home = Path.home()
    if sys.platform == "darwin":
        return home / "Library" / "Fonts"
    if os.name == "nt":
        return Path(os.environ.get("LOCALAPPDATA") or home / "AppData" / "Local") / "Microsoft" / "Windows" / "Fonts"
    return Path(os.environ.get("XDG_DATA_HOME") or home / ".local" / "share") / "fonts"


def _install_system(files: List[Path]) -> List[str]:
    target = system_font_dir()
    target.mkdir(parents=True, exist_ok=True)
    done = []
    for path in files:
        dest = target / path.name
        shutil.copy2(path, dest)
        if os.name == "nt":
            _register_windows_font(dest)
        done.append(str(dest))
    if sys.platform.startswith("linux") and shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f", str(target)], capture_output=True, timeout=300, check=False)
    return done


def _register_windows_font(path: Path) -> None:
    """Windows 的用户字体登记到注册表（HKCU）后，其他程序才能看到。"""
    import winreg
    faces = read_font_faces(str(path))
    name = f"{faces[0].families[0]} {faces[0].style} (TrueType)" if faces else path.stem
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts",
                            0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(path))
    try:  # 通知正在运行的程序字体有变化
        import ctypes
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001D, 0, 0, 0x0002, 1000, None)
    except Exception:
        pass


def install_font(key: str, progress: Optional[Callable[[str], None]] = None, system: bool = False) -> Dict:
    """下载目录中的开源字体，校验后生成固定字重的 TrueType 文件，放到用户数据目录（生成 PDF 时嵌入）；
    system=True 时再复制到当前用户的字体目录，Word、PowerPoint 等软件也能用。"""
    global _FACES
    if key not in CATALOG:
        raise ValueError(f"未知字体：{key}；可选：{', '.join(CATALOG)}")
    spec = CATALOG[key]
    if any("instances" in source for source in spec["sources"]):
        try:
            import fontTools.varLib.instancer  # noqa: F401
        except ImportError:
            raise RuntimeError('安装字体需要 fonttools：pip install "fonttools>=4.40"') from None
    target = data_dir() / "fonts"
    target.mkdir(parents=True, exist_ok=True)
    stem = spec["file_stem"]
    files: List[Path] = []
    for n, source in enumerate(spec["sources"]):
        download = target / f"{stem}.{n}.download"
        try:
            if progress:
                progress(f"下载 {spec['family']}（{source['size'] / 1e6:.1f} MB）")
            if not _verified(_download(source["url"], download, progress, source.get("size")), source):
                raise RuntimeError("字体文件校验失败（哈希不一致），已放弃安装")
            if "static" in source:
                out = target / f"{stem}-{source['static']}.ttf"
                os.replace(download, out)
                files.append(out)
                continue
            for style, location in source["instances"].items():
                if progress:
                    progress(f"生成 {style} 字重")
                out = target / f"{stem}-{style}.ttf"
                _static_instance(download, location, spec["family"], style).save(str(out))
                files.append(out)
        finally:
            if download.exists():
                download.unlink()
    license_file = target / f"{stem}-OFL.txt"
    _download(spec["license_url"], license_file)
    _FACES = None  # 下次查找字体时重新扫描
    _REGISTERED.clear()
    result = {"key": key, "family": spec["family"], "license": spec["license"],
              "files": [str(f) for f in files] + [str(license_file)]}
    if system:
        result["system_files"] = _install_system(files)
    return result


def catalog_installed(key: str) -> bool:
    """目录里的字体是否已经装进用户数据目录（看第一个字重的文件）。"""
    spec = CATALOG[key]
    first = spec["sources"][0]
    style = first.get("static") or next(iter(first["instances"]))
    return (data_dir() / "fonts" / f"{spec['file_stem']}-{style}.ttf").is_file()


def fonts_report() -> Dict:
    """字体情况汇总：PDF 将使用的中文字体、可下载的字体、各内置主题的字体情况。"""
    try:
        pdf = pdf_fonts().describe()
    except Exception as exc:  # reportlab 未安装等
        pdf = {"error": str(exc)}
    catalog = [{"key": key, "family": spec["family"], "license": spec["license"], "installed": catalog_installed(key),
                "download_mb": round(sum(source["size"] for source in spec["sources"]) / 1e6, 1)}
               for key, spec in CATALOG.items()]
    return {"pdf_cjk": pdf, "catalog": catalog, "system_dir": str(system_font_dir()),
            "themes": {name: theme_font_status(name) for name in builtin_themes()}}


def builtin_themes() -> List[str]:
    from .themes import list_themes
    return [t["name"] for t in list_themes() if t["source"] == "builtin" and not t.get("shadowed")]
