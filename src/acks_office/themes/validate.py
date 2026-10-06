"""主题校验（theme validate）：字段与取值、对比度、字体安装与授权、常用汉字覆盖、模板与资源。

错误（errors）会让主题无法正确出品；提醒（warnings）不影响出品，但值得处理，比如字体没装。
"""

from typing import Any, Dict, List, Optional, Tuple

from .. import fonts
from .loader import BASE, BUILTIN_DIR, load_theme
from .model import FIELDS, HEX, PLACEHOLDER, PLACEHOLDERS, Theme, ThemeRef

VARIANTS = {
    "layout.doc.cover": ("title-block", "standard", "issue"),
    "layout.doc.chapter": ("inline", "opener"),
    "layout.doc.callout": ("tint", "rule"),
    "layout.doc.quote": ("bar", "pull"),
    "layout.doc.caption": ("plain", "accent"),
    "layout.doc.table.head": ("fill", "label"),
    "layout.doc.page.size": ("A4", "A5", "Letter"),
    "layout.slide.cover": ("simple", "split", "issue"),
    "layout.slide.section": ("light", "dark"),
    "layout.slide.content": ("stacked", "columns"),
    "layout.sheet.head": ("fill", "label"),
}
COLOR_REFS = ("layout.doc.table.head_rule", "layout.doc.table.row_rule", "layout.doc.table.total_rule",
              "layout.sheet.head_rule", "layout.sheet.row_rule", "layout.sheet.total_rule",
              "layout.sheet.negative", "layout.sheet.tab_color")
RANGES = (("type.", 4, 400), ("leading.", 1.0, 3.0), ("space.", 0, 400), ("tracking.", 0, 1),
          ("layout.doc.page.margin_cm.", 0.5, 6), ("layout.doc.page.header_cm", 0.2, 4),
          ("layout.doc.page.footer_cm", 0.2, 4), ("layout.doc.columns", 1, 3),
          ("layout.doc.column_gap_cm", 0.2, 3), ("layout.slide.margin.", 0, 200))
# (前景, 背景, 最低对比度, 用途)：正文 4.5:1；大字、线条和标签 3:1（WCAG AA）
CONTRAST = (("ink", "paper", 4.5, "正文"), ("muted", "paper", 4.5, "次要文字"),
            ("ink", "surface", 4.5, "浅底上的文字"), ("muted", "surface", 4.5, "浅底上的次要文字"),
            ("accent", "paper", 3.0, "强调色（大字、编号与线条）"), ("negative", "paper", 4.5, "负数"),
            ("ink", "slide_paper", 4.5, "幻灯片正文"), ("muted", "slide_paper", 4.5, "幻灯片次要文字"),
            ("accent", "slide_paper", 3.0, "幻灯片强调色"), ("on_dark", "dark", 4.5, "深色页标题"),
            ("on_dark_soft", "dark", 4.5, "深色页正文"), ("on_dark_muted", "dark", 3.0, "深色页标签"),
            ("accent_on_dark", "dark", 3.0, "深色页强调色"))
OPEN_LICENSES = ("OFL-1.1", "Apache-2.0", "MIT", "Ubuntu-1.0", "CC0-1.0", "IPA", "GPL-FE")
MIN_HANZI_COVERAGE = 0.99


def luminance(color: str) -> float:
    channels = []
    for i in (1, 3, 5):
        c = int(color[i:i + 2], 16) / 255
        channels.append(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4)
    r, g, b = channels
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(fg: str, bg: str) -> float:
    high, low = sorted((luminance(fg), luminance(bg)), reverse=True)
    return (high + 0.05) / (low + 0.05)


def _walk(node: Any, prefix: str = ""):
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _walk(value, f"{prefix}{key}.")
    else:
        yield prefix[:-1], node


def _lookup(node: Any, path: str) -> Tuple[bool, Any]:
    for key in path.split("."):
        if not isinstance(node, dict) or key not in node:
            return False, None
        node = node[key]
    return True, node


def _same_type(a: Any, b: Any) -> bool:
    number = (int, float)
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool)
    if isinstance(a, number) and isinstance(b, number):
        return True
    if isinstance(b, str) and isinstance(a, list):  # 字体名可以写成候选列表
        return all(isinstance(x, str) for x in a)
    return type(a) is type(b)


def _kind(value: Any) -> str:
    if isinstance(value, bool):
        return "布尔值"
    if isinstance(value, (int, float)):
        return "数字"
    return {str: "文本", list: "列表", dict: "对象"}.get(type(value), type(value).__name__)


class _Report:
    def __init__(self) -> None:
        self.errors: List[Dict[str, str]] = []
        self.warnings: List[Dict[str, str]] = []

    def error(self, code: str, path: str, message: str) -> None:
        self.errors.append({"code": code, "path": path, "message": message})

    def warn(self, code: str, path: str, message: str) -> None:
        self.warnings.append({"code": code, "path": path, "message": message})


def _check_fields(theme: Theme, base: Theme, report: _Report) -> None:
    own_tokens = theme.own.get("tokens", {})
    for path, value in _walk(own_tokens):
        if path == "schema":
            continue
        found, expected = _lookup(base.tokens, path)
        if not found:
            # 自定义颜色可以随意命名，供版式里的颜色引用使用
            if not path.startswith("color.") and not path.startswith("font.families."):
                report.warn("UNKNOWN_FIELD", path, f"基础主题里没有这个字段，可能是拼写错误：{path}")
            continue
        if not _same_type(value, expected):
            report.error("WRONG_TYPE", path, f"{path} 应为{_kind(expected)}，实际是{_kind(value)}：{value!r}")
    for path, _value in _walk(theme.own.get("chrome", {})):
        if path in ("schema",) or path.startswith("labels."):
            continue
        found, _ = _lookup(base.chrome, path)
        if not found and path not in ("brand", "classification", "publication"):
            report.warn("UNKNOWN_FIELD", f"chrome.{path}", f"品牌外壳里没有这个字段：{path}")


def _check_values(theme: Theme, report: _Report) -> None:
    for key, value in theme.tokens.get("color", {}).items():
        values = value if isinstance(value, list) else [value]
        if not values or not all(isinstance(v, str) and HEX.match(v) for v in values):
            report.error("BAD_COLOR", f"color.{key}", f"颜色应写成 #RRGGBB：{value}")
    for path, allowed in VARIANTS.items():
        value = theme.get(path, None)
        if value not in allowed:
            report.error("BAD_VARIANT", path, f"{path} 只能是 {' / '.join(allowed)}，实际是 {value}")
    for path in COLOR_REFS:
        value = theme.get(path, None)
        if not isinstance(value, str) or not (HEX.match(value) or value in theme.tokens.get("color", {})):
            report.error("BAD_COLOR_REF", path, f"{path} 应为颜色名或 #RRGGBB：{value}")
    for path, value in _walk(theme.tokens):
        for prefix, low, high in RANGES:
            if path.startswith(prefix) and isinstance(value, (int, float)) and not isinstance(value, bool):
                if not low <= value <= high:
                    report.error("OUT_OF_RANGE", path, f"{path} = {value}，应在 {low} 到 {high} 之间")
    families = theme.get("font.families", {})
    for role, spec in theme.get("font.roles", {}).items():
        if not isinstance(spec, dict) or spec.get("family") not in families:
            report.error("BAD_FONT_ROLE", f"font.roles.{role}", f"字体角色 {role} 引用了不存在的字体槽位")
        elif not 100 <= int(spec.get("weight", 400)) <= 900:
            report.error("OUT_OF_RANGE", f"font.roles.{role}.weight", "字重应在 100 到 900 之间")


def _check_contrast(theme: Theme, report: _Report) -> List[Dict[str, Any]]:
    results = []
    for fg, bg, minimum, use in CONTRAST:
        try:
            ratio = contrast_ratio(theme.color(fg), theme.color(bg))
        except Exception:
            continue
        ok = ratio >= minimum
        results.append({"foreground": fg, "background": bg, "use": use,
                        "ratio": round(ratio, 2), "minimum": minimum, "ok": ok})
        if not ok:
            report.error("LOW_CONTRAST", f"color.{fg}",
                         f"{use}：{fg} 在 {bg} 上对比度 {ratio:.2f}:1，低于 {minimum}:1")
    return results


def _check_fonts(theme: Theme, report: _Report, check_installed: bool) -> List[Dict[str, Any]]:
    results = []
    registry = theme.fonts.get("families", {})
    hanzi: Optional[List[str]] = None
    seen = set()
    for slot, scripts in theme.get("font.families", {}).items():
        for script, spec in scripts.items():
            primary = spec if isinstance(spec, str) else spec[0]
            if (primary, script) in seen:
                continue
            seen.add((primary, script))
            meta = registry.get(primary, {})
            license_ = meta.get("license", "")
            entry = {"family": primary, "slot": slot, "script": script, "license": license_ or "未登记"}
            if not license_:
                report.warn("FONT_LICENSE_UNKNOWN", f"font.families.{slot}.{script}",
                            f"{primary} 没有在 fonts.json 登记授权；请确认可以使用")
            elif license_ not in OPEN_LICENSES and license_ != "system":
                report.warn("FONT_COMMERCIAL", f"font.families.{slot}.{script}",
                            f"{primary} 的授权为 {license_}：只引用、不随主题分发，请确认用户有使用授权")
            if check_installed:
                face = fonts.best_face(primary)
                entry["installed"] = face is not None
                if face is None:
                    hint = f"，可运行 fonts install {meta['install']}" if meta.get("install") else ""
                    used = next((n for n in theme.family_candidates(slot, script)[1:]
                                 if fonts.family_installed(n)), None)
                    entry["substitute"] = used
                    report.warn("FONT_NOT_INSTALLED", f"font.families.{slot}.{script}",
                                f"{primary} 未安装{('，生成时改用 ' + used) if used else ''}{hint}")
                elif script == "cn":
                    if hanzi is None:
                        hanzi = fonts.common_hanzi()
                    points = fonts.face_codepoints(face)
                    coverage = sum(ord(c) in points for c in hanzi) / len(hanzi)
                    entry["hanzi_coverage"] = round(coverage, 4)
                    if coverage < MIN_HANZI_COVERAGE:
                        report.error("FONT_COVERAGE", f"font.families.{slot}.{script}",
                                     f"{primary} 只覆盖 {coverage:.1%} 的常用汉字（GB2312 一级字）")
            results.append(entry)
    return results


def _check_chrome(theme: Theme, report: _Report) -> None:
    for path, value in _walk(theme.chrome):
        if not isinstance(value, str):
            continue
        for name in PLACEHOLDER.findall(value):
            if name not in PLACEHOLDERS:
                report.error("BAD_PLACEHOLDER", f"chrome.{path}",
                             f"未知占位符 {{{name}}}，可用：{', '.join(PLACEHOLDERS)}")
    for key, value in (theme.chrome.get("logo") or {}).items():
        if value and theme.asset(value) is None:
            report.error("ASSET_MISSING", f"chrome.logo.{key}", f"找不到 logo 文件：{value}")
    for key in theme.chrome_get("doc.cover_meta", []):
        if key not in PLACEHOLDERS or key in FIELDS:
            report.error("BAD_PLACEHOLDER", "chrome.doc.cover_meta", f"封面信息项只能用元数据字段：{key}")


def validate_theme(ref: ThemeRef, check_installed: bool = True) -> Dict[str, Any]:
    """校验主题。返回 {theme, ok, errors, warnings, contrast, fonts}；ok 表示没有错误。"""
    report = _Report()
    theme = load_theme(ref, use_cache=False)
    base = load_theme(str(BUILTIN_DIR / BASE), use_cache=False)
    meta = theme.meta
    if not meta.get("title"):
        report.warn("MISSING_FIELD", "theme.title", "theme.json 缺少 title（展示名称）")
    for name in ("version", "license"):
        if not meta.get(name):
            report.warn("MISSING_FIELD", f"theme.{name}", f"theme.json 缺少 {name}")
    _check_fields(theme, base, report)
    _check_values(theme, report)
    contrast = _check_contrast(theme, report)
    font_results = _check_fonts(theme, report, check_installed)
    _check_chrome(theme, report)
    return {"theme": theme.name, "title": theme.title, "path": str(theme.path), "extends": theme.chain[1:],
            "ok": not report.errors, "errors": report.errors, "warnings": report.warnings,
            "contrast": contrast, "fonts": font_results}


__all__ = ["validate_theme", "contrast_ratio"]
