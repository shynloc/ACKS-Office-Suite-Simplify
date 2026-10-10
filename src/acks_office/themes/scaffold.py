"""新建主题（theme init）：生成一个继承内置主题的主题包，只写和父主题不同的部分。"""

import json
from pathlib import Path
from typing import Any, Dict, Optional

from .loader import NAME, clear_cache, load_theme, user_theme_dir
from .model import HEX, ThemeError
from .validate import validate_theme

README = """# {title}

acks-office 主题包，继承自 `{extends}`（{parent_title}）。各文件只写和父主题不同的部分，其余沿用父主题。

| 文件 | 内容 |
|---|---|
| theme.json | 名称、继承的主题、说明、授权（默认 UNLICENSED 即私有，公开分享时改成相应的授权） |
| tokens.json | 颜色、字体、字号、间距、版式变体 |
| chrome.json | 品牌名、logo、页眉页脚模板、标签文字 |
| fonts.json | 字体的授权、备选字体与下载键（可选，新增字体时再写） |

常改的几项（写进 tokens.json 或 chrome.json 对应的位置）：

- 强调色：`color.accent`；深色页上的强调色：`color.accent_on_dark`
- 中西文字体：`font.families.sans.cn` / `font.families.sans.en`（衬线体在 `serif`）
- 封面：`layout.doc.cover`（title-block / standard / issue）；幻灯片封面：`layout.slide.cover`（simple / split / issue）
- 表格：`layout.doc.table.head`（fill / label）、`layout.doc.table.total`（rules / bar）
- 品牌名：chrome.json 的 `brand`；页脚：`doc.footer.left` / `doc.footer.right`，可用 {{brand}}、{{title}}、{{page}} 等

改完运行 `acks-office theme validate {name}` 检查对比度、字体与取值，`acks-office theme preview {name}` 生成样张。
"""


def init_theme(name: str, extends: str = "neutral", accent: Optional[str] = None, brand: Optional[str] = None,
               title: Optional[str] = None, directory: Optional[str] = None, force: bool = False) -> Dict[str, Any]:
    """新建主题包，返回目录、文件和校验结果。默认放在用户主题目录，生成后即可按名称使用。"""
    if not NAME.match(name) or name.startswith("_"):
        raise ThemeError("BAD_THEME_NAME", f"主题名只能用小写字母、数字、- 和 _，且不以 _ 开头：{name}",
                         "例如 theme init my-brand")
    if accent and not HEX.match(accent):
        raise ThemeError("BAD_COLOR", f"颜色应写成 #RRGGBB：{accent}", "例如 --accent #0B6E4F")
    parent = load_theme(extends)
    target = Path(directory).expanduser() if directory else user_theme_dir() / name
    # --force 只覆盖这里要写的几个文件，目录里的其他文件（logo 等）保留
    if target.exists() and any(target.iterdir()) and not force:
        raise ThemeError("THEME_EXISTS", f"主题目录已存在：{target}", "换个名字，或加 --force 覆盖主题文件")
    target.mkdir(parents=True, exist_ok=True)

    files = {
        "theme.json": {"schema": "acks-office-theme/1", "name": name, "title": title or name,
                       "description": f"基于 {parent.title} 的主题", "extends": extends, "version": "0.1.0",
                       "license": "UNLICENSED"},  # 私有主题；要公开分享时改成相应的授权
        "tokens.json": {"schema": "acks-office-tokens/1", **({"color": {"accent": accent.upper()}} if accent else {})},
        "chrome.json": {"schema": "acks-office-chrome/1", **({"brand": brand} if brand is not None else {})},
    }
    written = []
    for filename, data in files.items():
        path = target / filename
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        written.append(str(path))
    readme = target / "README.md"
    readme.write_text(README.format(title=title or name, name=name, extends=extends, parent_title=parent.title),
                      encoding="utf-8")
    written.append(str(readme))
    clear_cache()
    report = validate_theme(str(target))
    return {"theme": name, "path": str(target), "extends": extends, "files": written,
            "user_dir": directory is None, "validation": report}
