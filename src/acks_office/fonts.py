"""字体：为 PDF 找到能嵌入的中文字体，检查主题字体是否已安装，并按需下载开源字体。

这个模块只依赖标准库；reportlab、fontTools 在用到时才导入，所以 doctor 在依赖不全时也能运行。
"""

import hashlib
import json
import os
import re
import sys
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from .utils import data_dir

CID_FALLBACK = "STSong-Light"
_FONT_EXTENSIONS = (".ttf", ".ttc", ".otf", ".otc")

# 可以按需下载的开源字体：地址固定到具体提交，下载后校验 SHA-256
CATALOG: Dict[str, Dict] = {
    "noto-sans-sc": {
        "family": "Noto Sans SC",
        "license": "SIL Open Font License 1.1",
        "url": "https://raw.githubusercontent.com/google/fonts/2894aab31764f10f29c421bdfd2340d3b382d384/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf",
        "sha256": "a3041811a78c361b1de50f953c805e0244951c21c5bd412f7232ef0d899af0da",
        "license_url": "https://raw.githubusercontent.com/google/fonts/2894aab31764f10f29c421bdfd2340d3b382d384/ofl/notosanssc/OFL.txt",
        "file_stem": "NotoSansSC",
        "instances": {"Regular": 400, "Bold": 700},
    },
}

# 主题字体的常见文件名写法（思源黑体也以 Noto Sans SC、Source Han Sans 等名称发布）
_ALIASES = {
    "noto sans sc": ["notosanssc", "notosanscjksc", "sourcehansanssc", "sourcehansanscn"],
    "noto serif sc": ["notoserifsc", "notoserifcjksc", "sourcehanserifsc", "sourcehanserifcn"],
    "pingfang sc": ["pingfang"],
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


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def font_installed(family: str, _stems: Optional[List[str]] = None) -> bool:
    """按文件名判断某个字体家族是否已安装（快速检查，不读取字体内部名称）。"""
    stems = _stems if _stems is not None else [_norm(p.stem) for p in _font_files()]
    keys = _ALIASES.get(family.lower(), [_norm(family)])
    return any(key in stem for key in keys for stem in stems)


def theme_fonts() -> Dict[str, str]:
    """ACKS 主题声明的字体（来自 tokens.json）。"""
    tokens = Path(__file__).parent / "design_system" / "tokens.json"
    return json.loads(tokens.read_text(encoding="utf-8"))["font"]["family"]


def theme_font_status() -> Dict[str, List[str]]:
    """主题字体在本机的安装情况：present / missing。回退字体和表格字体不计入缺失。"""
    stems = [_norm(p.stem) for p in _font_files()]
    present, missing = [], []
    for role, family in theme_fonts().items():
        if role in ("FALLBACK_CN", "XLSX"):
            continue
        (present if font_installed(family, stems) else missing).append(family)
    return {"present": sorted(set(present)), "missing": sorted(set(missing))}


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

def _download(url: str, dest: Path, progress: Optional[Callable[[str], None]] = None) -> str:
    mirror = os.environ.get("ACKS_OFFICE_DOWNLOAD_MIRROR")
    if mirror:
        url = url.replace("https://raw.githubusercontent.com", mirror.rstrip("/"))
    request = urllib.request.Request(url, headers={"User-Agent": "acks-office"})
    digest = hashlib.sha256()
    with urllib.request.urlopen(request, timeout=60) as response, open(dest, "wb") as out:
        total = int(response.headers.get("Content-Length") or 0)
        done = 0
        while True:
            chunk = response.read(1 << 20)
            if not chunk:
                break
            out.write(chunk)
            digest.update(chunk)
            done += len(chunk)
            if progress and total:
                progress(f"下载 {done * 100 // total}%")
    return digest.hexdigest()


def install_font(key: str, progress: Optional[Callable[[str], None]] = None) -> Dict:
    """下载目录中的开源字体，校验后生成固定字重的 TrueType 文件，放到用户数据目录。"""
    if key not in CATALOG:
        raise ValueError(f"未知字体：{key}；可选：{', '.join(CATALOG)}")
    try:
        from fontTools.ttLib import TTFont
        from fontTools.varLib import instancer
    except ImportError:
        raise RuntimeError('安装字体需要 fonttools：pip install "fonttools>=4.40"') from None

    spec = CATALOG[key]
    target = data_dir() / "fonts"
    target.mkdir(parents=True, exist_ok=True)
    download = target / f"{spec['file_stem']}.download"
    try:
        if _download(spec["url"], download, progress) != spec["sha256"]:
            raise RuntimeError("字体文件校验失败（SHA-256 不一致），已放弃安装")
        files = []
        for style, weight in spec["instances"].items():
            if progress:
                progress(f"生成 {style} 字重")
            static = instancer.instantiateVariableFont(TTFont(download), {"wght": weight})
            out = target / f"{spec['file_stem']}-{style}.ttf"
            static.save(out)
            files.append(str(out))
        license_file = target / f"{spec['file_stem']}-OFL.txt"
        _download(spec["license_url"], license_file)
    finally:
        if download.exists():
            download.unlink()
    _REGISTERED.clear()
    return {"key": key, "family": spec["family"], "license": spec["license"],
            "files": files + [str(license_file)]}


def fonts_report() -> Dict:
    """字体情况汇总：PDF 将使用的中文字体、可下载字体、主题字体安装情况。"""
    try:
        pdf = pdf_fonts().describe()
    except Exception as exc:  # reportlab 未安装等
        pdf = {"error": str(exc)}
    base = data_dir() / "fonts"
    catalog = [{"key": k, "family": s["family"], "license": s["license"],
                "installed": (base / f"{s['file_stem']}-Regular.ttf").is_file()}
               for k, s in CATALOG.items()]
    return {"pdf_cjk": pdf, "catalog": catalog, "theme": theme_font_status()}
