"""主题对象：合并后的设计令牌、字体与品牌外壳，以及取值、选字体、填写页眉页脚模板的方法。"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from .. import fonts

HEX = re.compile(r"^#[0-9A-Fa-f]{6}$")
PLACEHOLDER = re.compile(r"\{(\w+)\}")
# 页眉页脚、封面模板里可用的占位符；page、pages 由渲染器写成页码域
PLACEHOLDERS = ("title", "short_title", "subtitle", "brand", "publication", "issue", "season",
                "date", "author", "version", "classification", "chapter", "page", "pages")
FIELDS = ("page", "pages")
SEPARATOR = " · "
_MISSING = object()


class ThemeError(Exception):
    """主题找不到或内容不合规。code 与命令行的错误码一致。"""

    def __init__(self, code: str, message: str, hint: Optional[str] = None):
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint


@dataclass
class FontChoice:
    """某个字体角色实际写进文档的字体。substitutions 记录主题字体没装、改用替代字体的情况。"""
    cn: str
    en: str
    weight: int
    bold: bool
    substitutions: List[Dict[str, str]] = field(default_factory=list)


def render_template(template: str, values: Dict[str, Any]) -> List[Tuple[str, str]]:
    """把页眉页脚模板填成 [(kind, value)]：kind 为 "text" 或 "field"（page / pages）。

    模板按「 · 」分段，一段里的占位符都没有值时整段省略，避免出现「 · 经营回顾」这类残缺。
    """
    kept = []
    for part in template.split(SEPARATOR):
        names = PLACEHOLDER.findall(part)
        if names and not any(n in FIELDS or str(values.get(n) or "").strip() for n in names):
            continue
        if part.strip():
            kept.append(part)
    segments: List[Tuple[str, str]] = []
    for i, part in enumerate(kept):
        if i:
            segments.append(("text", SEPARATOR))
        pos = 0
        for m in PLACEHOLDER.finditer(part):
            segments.append(("text", part[pos:m.start()]))
            name = m.group(1)
            segments.append(("field", name) if name in FIELDS else ("text", str(values.get(name) or "")))
            pos = m.end()
        segments.append(("text", part[pos:]))
    merged: List[Tuple[str, str]] = []
    for kind, value in segments:
        if kind == "text" and not value:
            continue
        if kind == "text" and merged and merged[-1][0] == "text":
            merged[-1] = ("text", merged[-1][1] + value)
        else:
            merged.append((kind, value))
    return merged


def plain_template(template: str, values: Dict[str, Any]) -> str:
    """模板的纯文本形式（页码域写成占位文字），用于预览与测试。"""
    return "".join(v if k == "text" else "{" + v + "}" for k, v in render_template(template, values))


class Theme:
    """一套主题：meta（theme.json）、tokens、fonts、chrome 都已按继承关系合并。"""

    def __init__(self, name: str, path: Path, source: str, chain: List[str],
                 data: Dict[str, Dict], own: Dict[str, Dict], dirs: List[Path]):
        self.name = name
        self.path = path
        self.source = source
        self.chain = chain          # 自身在前，基础主题在后
        self.meta = data["theme"]
        self.tokens = data["tokens"]
        self.fonts = data["fonts"]
        self.chrome = data["chrome"]
        self.own = own              # 主题目录里自己写的内容（未合并），供校验
        self.dirs = dirs            # 自身与各级父主题的目录，用于查找 assets
        self._fonts: Dict[Tuple[str, str], Tuple[str, List[Dict[str, str]]]] = {}

    # ------------------------------------------------------------ 基本信息

    @property
    def title(self) -> str:
        return self.meta.get("title") or self.name

    def __repr__(self) -> str:
        return f"<Theme {self.name} ({self.source})>"

    def get(self, path: str, default: Any = _MISSING) -> Any:
        """按点号路径取令牌，如 get("type.doc.h1")、get("layout.doc.cover")。"""
        node: Any = self.tokens
        for key in path.split("."):
            if isinstance(node, dict) and key in node:
                node = node[key]
            elif default is not _MISSING:
                return default
            else:
                raise ThemeError("THEME_INVALID", f"主题 {self.name} 缺少令牌：{path}")
        return node

    def size(self, path: str) -> float:
        return float(self.get(path))

    # ------------------------------------------------------------ 颜色

    def color(self, ref: str) -> str:
        """颜色名（如 accent）或 #RRGGBB → #RRGGBB。"""
        if HEX.match(ref):
            return ref.upper()
        value = self.tokens.get("color", {}).get(ref)
        if isinstance(value, str) and HEX.match(value):
            return value.upper()
        raise ThemeError("THEME_INVALID", f"主题 {self.name} 没有颜色：{ref}")

    def hex(self, ref: str) -> str:
        """去掉 # 的 RRGGBB，写 Word / PPT / Excel 时用。"""
        return self.color(ref)[1:]

    def rgb(self, ref: str) -> Tuple[int, int, int]:
        h = self.hex(ref)
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    def chart_colors(self) -> List[str]:
        return [c.upper() for c in self.get("color.chart")]

    # ------------------------------------------------------------ 字体

    def family_candidates(self, family_key: str, script: str) -> List[str]:
        """某个字体家族槽位（sans / serif / display / quote / mono）的候选字体，主题字体在前。"""
        spec = self.get(f"font.families.{family_key}.{script}")
        names = [spec] if isinstance(spec, str) else list(spec)
        for name in list(names):
            for alt in self.fonts.get("families", {}).get(name, {}).get("fallback", []):
                if alt not in names:
                    names.append(alt)
        return names

    def family_name(self, family_key: str, script: str, policy: str = "local") -> Tuple[str, List[Dict[str, str]]]:
        """选出写进文档的字体名：policy="local" 时取本机已安装的第一个候选；
        "theme" 时总是写主题字体名（由打开文件的软件替换缺失字体）。"""
        cache_key = (f"{family_key}.{script}", policy)
        if cache_key in self._fonts:
            return self._fonts[cache_key]
        names = self.family_candidates(family_key, script)
        primary, notes = names[0], []
        chosen = primary
        if not fonts.family_installed(primary):
            installed = next((n for n in names[1:] if fonts.family_installed(n)), None)
            meta = self.fonts.get("families", {}).get(primary, {})
            note = {"code": "FONT_SUBSTITUTED" if installed and policy == "local" else "FONT_MISSING",
                    "family": primary, "install": meta.get("install", "")}
            if installed and policy == "local":
                chosen = installed
                note["used"] = installed
                note["message"] = f"主题字体 {primary} 未安装，已改用 {installed}"
            else:
                note["message"] = f"主题字体 {primary} 未安装，打开文件的软件会用其他字体替代"
            if meta.get("install"):
                note["message"] += f"；可运行 fonts install {meta['install']} 安装"
            notes.append(note)
        self._fonts[cache_key] = (chosen, notes)
        return chosen, notes

    def font(self, role: str, policy: str = "local") -> FontChoice:
        """字体角色（body / title / heading / label / number / display / quote / code）→ 字体选择。"""
        spec = self.get(f"font.roles.{role}")
        family_key, weight = spec["family"], int(spec.get("weight", 400))
        cn, cn_notes = self.family_name(family_key, "cn", policy)
        en, en_notes = self.family_name(family_key, "en", policy)
        return FontChoice(cn, en, weight, weight >= 600, cn_notes + en_notes)

    # ------------------------------------------------------------ 品牌外壳

    def chrome_get(self, path: str, default: Any = "") -> Any:
        node: Any = self.chrome
        for key in path.split("."):
            if not isinstance(node, dict) or key not in node:
                return default
            node = node[key]
        return node

    def label(self, key: str) -> str:
        return self.chrome_get(f"labels.{key}", key)

    def chrome_values(self, meta: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """模板可用的值：主题外壳里的默认品牌信息，被文档元数据覆盖。"""
        values = {k: self.chrome.get(k, "") for k in ("brand", "classification", "publication")}
        for key, value in (meta or {}).items():
            if value is not None:
                values[key] = value
        if not values.get("short_title"):
            values["short_title"] = values.get("title", "")
        return values

    def template(self, path: str, meta: Optional[Dict[str, Any]] = None) -> List[Tuple[str, str]]:
        return render_template(self.chrome_get(path, ""), self.chrome_values(meta))

    def asset(self, relative: str) -> Optional[Path]:
        """在主题目录及各级父主题目录里查找资源文件（如 logo）。"""
        if not relative:
            return None
        for directory in self.dirs:
            candidate = (directory / relative).resolve()
            if candidate.is_file() and directory.resolve() in candidate.parents:
                return candidate
        return None

    def describe(self) -> Dict[str, Any]:
        layout = self.get("layout")
        return {
            "name": self.name, "title": self.title, "description": self.meta.get("description", ""),
            "version": self.meta.get("version", ""), "source": self.source, "path": str(self.path),
            "extends": self.chain[1:],
            "colors": {k: v for k, v in self.tokens["color"].items()},
            "fonts": {k: v for k, v in self.tokens["font"]["families"].items()},
            "variants": {
                "doc": {k: layout["doc"][k] for k in ("cover", "chapter", "columns", "callout", "quote")},
                "slide": {k: layout["slide"][k] for k in ("cover", "section", "content")},
                "sheet": {"head": layout["sheet"]["head"]},
            },
        }


ThemeRef = Union[str, Path, Theme, None]
