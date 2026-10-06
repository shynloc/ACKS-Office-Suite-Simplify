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


_REGISTERED: Dict[Tuple[str, int], Optional[str]] = {}


def _register_ttf(path: str, index: int = 0) -> Optional[str]:
    """把一个 TrueType 字体注册给 reportlab；不能用（CFF 轮廓、缺中文字形等）返回 None。"""
    key = (os.path.abspath(path), index)
    if key in _REGISTERED:
        return _REGISTERED[key]
    name = None
    if os.path.isfile(path):
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            name = "acks-" + hashlib.sha1(f"{key[0]}#{index}".encode()).hexdigest()[:10]
            font = TTFont(name, path, subfontIndex=index)
            if not {0x4E2D, 0x6587} <= set(font.face.charToGlyph):  # 「中」「文」
                name = None
            else:
                pdfmetrics.registerFont(font)
        except Exception:
            name = None
    _REGISTERED[key] = name
    return name


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
