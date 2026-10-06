"""查找与加载主题包：按名称或目录找到主题，沿 extends 合并父主题。

查找顺序：主题目录路径 → 环境变量 ACKS_OFFICE_THEMES 列出的目录 → 用户数据目录下的 themes/
→ 内置主题。同名时排在前面的生效，所以用户可以用同名主题覆盖内置主题。
"""

import copy
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple

from ..utils import data_dir
from .model import Theme, ThemeError, ThemeRef

BUILTIN_DIR = Path(__file__).parent / "builtin"
BASE = "_base"
DEFAULT_THEME = "neutral"
FILES = ("theme", "tokens", "fonts", "chrome")
NAME = re.compile(r"^_?[a-z0-9][a-z0-9_-]{0,63}$")  # 以 _ 开头的是内部主题
_CACHE: Dict[str, Theme] = {}


def user_theme_dir() -> Path:
    return data_dir() / "themes"


def search_dirs() -> List[Tuple[str, Path]]:
    dirs: List[Tuple[str, Path]] = []
    for entry in os.environ.get("ACKS_OFFICE_THEMES", "").split(os.pathsep):
        if entry.strip():
            dirs.append(("env", Path(entry.strip()).expanduser()))
    dirs.append(("user", user_theme_dir()))
    dirs.append(("builtin", BUILTIN_DIR))
    return dirs


def _available() -> List[str]:
    return sorted({t["name"] for t in list_themes()})


def _locate(ref: str, exclude: Set[Path] = frozenset()) -> Tuple[str, Path]:
    path = Path(ref).expanduser()
    if path.name == "theme.json" and path.is_file():
        path = path.parent
    if (path / "theme.json").is_file():
        return "path", path.resolve()
    if "/" in ref or "\\" in ref or ref.startswith("."):
        raise ThemeError("THEME_NOT_FOUND", f"主题目录不存在或缺少 theme.json：{ref}")
    if not NAME.match(ref):
        raise ThemeError("THEME_NOT_FOUND", f"主题名只能用小写字母、数字、- 和 _：{ref}")
    for source, base in search_dirs():
        candidate = base / ref
        if (candidate / "theme.json").is_file() and candidate.resolve() not in exclude:
            return source, candidate.resolve()
    raise ThemeError("THEME_NOT_FOUND", f"找不到主题：{ref}",
                     f"可用主题：{', '.join(_available())}；也可以传主题目录的路径")


def _read_json(path: Path) -> Dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError) as exc:
        raise ThemeError("THEME_INVALID", f"读取失败：{path}（{exc}）") from None
    except json.JSONDecodeError as exc:
        raise ThemeError("THEME_INVALID", f"{path} 不是有效的 JSON：{exc}") from None
    if not isinstance(data, dict):
        raise ThemeError("THEME_INVALID", f"{path} 的内容应为 JSON 对象")
    return data


def read_package(directory: Path) -> Dict[str, Dict[str, Any]]:
    """读取主题目录里自己写的文件（不含父主题）。"""
    return {key: _read_json(directory / f"{key}.json") for key in FILES
            if (directory / f"{key}.json").is_file()}


def deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    """字典逐层合并，列表与标量整体替换。"""
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def _load(ref: str, seen: Set[Path], exclude: Set[Path]) -> Theme:
    source, directory = _locate(ref, exclude)
    if directory in seen:
        raise ThemeError("THEME_INVALID", f"主题继承出现循环：{ref}")
    own = read_package(directory)
    meta = own.get("theme", {})
    name = str(meta.get("name") or directory.name)
    parent_ref = meta.get("extends", None if directory == (BUILTIN_DIR / BASE).resolve() else BASE)
    if parent_ref:
        parent_path = Path(str(parent_ref)).expanduser()
        if not parent_path.is_absolute() and ("/" in str(parent_ref) or str(parent_ref).startswith(".")):
            parent_ref = str((directory / parent_path).resolve())
        # 按名称找父主题时跳过继承链上已有的目录：用户的 slate 可以继承内置的 slate
        chain_dirs = seen | {directory}
        parent = _load(str(parent_ref), chain_dirs, chain_dirs)
        data = {"tokens": parent.tokens, "fonts": parent.fonts, "chrome": parent.chrome}
        chain, dirs = [name] + parent.chain, [directory] + parent.dirs
    else:
        data, chain, dirs = {"tokens": {}, "fonts": {}, "chrome": {}}, [name], [directory]
    merged = {key: deep_merge(data[key], own.get(key, {})) for key in ("tokens", "fonts", "chrome")}
    merged["theme"] = dict(meta, name=name)
    return Theme(name, directory, source, chain, merged, own, dirs)


def load_theme(ref: ThemeRef = None, use_cache: bool = True) -> Theme:
    """按名称（neutral、slate、folio 或已安装的主题）或主题目录路径加载主题。"""
    if isinstance(ref, Theme):
        return ref
    ref = str(ref) if ref else DEFAULT_THEME
    key = f"{ref}|{os.environ.get('ACKS_OFFICE_THEMES', '')}|{data_dir()}"
    if use_cache and key in _CACHE:
        return _CACHE[key]
    theme = _load(ref, set(), set())
    _CACHE[key] = theme
    return theme


def clear_cache() -> None:
    _CACHE.clear()


def list_themes() -> List[Dict[str, Any]]:
    """可用主题（不含以 _ 开头的内部主题）。同名时只列生效的那一个，被覆盖的标为 shadowed。"""
    found: List[Dict[str, Any]] = []
    names: Set[str] = set()
    for source, base in search_dirs():
        if not base.is_dir():
            continue
        for directory in sorted(p for p in base.iterdir() if (p / "theme.json").is_file()):
            if directory.name.startswith("_"):
                continue
            try:
                meta = _read_json(directory / "theme.json")
            except ThemeError as exc:
                found.append({"name": directory.name, "source": source, "path": str(directory),
                              "error": exc.message})
                continue
            name = str(meta.get("name") or directory.name)
            found.append({"name": name, "title": meta.get("title", name),
                          "description": meta.get("description", ""), "source": source,
                          "path": str(directory), "extends": meta.get("extends", BASE),
                          "shadowed": name in names})
            names.add(name)
    return found
