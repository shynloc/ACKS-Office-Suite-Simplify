"""主题引擎：主题是运行时加载的数据包，同一个程序可以同时使用多套主题。

    from acks_office import themes
    slate = themes.load_theme("slate")
    slate.color("accent"), slate.font("heading")
    themes.validate_theme("path/to/my-brand")

内置主题：neutral（默认）、slate（商务）、folio（杂志）。
"""

from .loader import (BUILTIN_DIR, DEFAULT_THEME, clear_cache, deep_merge, list_themes, load_theme,
                     search_dirs, user_theme_dir)
from .model import FontChoice, Theme, ThemeError, plain_template, render_template
from .validate import contrast_ratio, validate_theme

__all__ = ["Theme", "ThemeError", "FontChoice", "load_theme", "list_themes", "validate_theme",
           "contrast_ratio", "render_template", "plain_template", "deep_merge", "search_dirs",
           "user_theme_dir", "clear_cache", "BUILTIN_DIR", "DEFAULT_THEME"]
