"""渲染公共部分：文档元数据（Markdown front matter 与参数合并）、标注与说明的识别、数字判断。"""

import re
from typing import Any, Dict, List, Optional, Tuple

from .. import markdown_blocks as mdb

# 封面、页眉页脚可用的元数据；其余 front matter 字段原样保留
META_KEYS = ("title", "subtitle", "kicker", "short_title", "brand", "classification", "publication",
             "issue", "season", "date", "author", "version", "lede", "cover")
_FRONT = re.compile(r"\A﻿?---[ \t]*\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|\Z)", re.S)
_CAPTION = re.compile(r"^\s*(表|Table|图|Figure)\s*[:：]\s*(.+)$", re.S)
_TOTAL = re.compile(r"^\s*(合计|总计|小计|总和|共计|Total|Sum|Subtotal)\b", re.I)
_NUMERIC = re.compile(r"^[\s(（]*[+\-−]?[¥$€£]?[\d,]+(\.\d+)?\s*[%‰]?[)）]*$")


def split_front_matter(text: str) -> Tuple[Dict[str, Any], str]:
    """拆出开头 --- 包围的 front matter（每行 key: value），返回 (元数据, 正文)。

    值两端的引号会去掉；true / false 转成布尔值；值里的 \\n 表示换行（如两行的封面标题）。
    """
    match = _FRONT.match(text or "")
    if not match:
        return {}, text or ""
    meta: Dict[str, Any] = {}
    for line in match.group(1).splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        key, sep, value = line.partition(":")
        if not sep or not key.strip():
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
            value = value[1:-1]
        if value.lower() in ("true", "yes", "on"):
            meta[key.strip().lower()] = True
        elif value.lower() in ("false", "no", "off"):
            meta[key.strip().lower()] = False
        else:
            meta[key.strip().lower()] = value.replace("\\n", "\n")
    return meta, (text or "")[match.end():]


def document_meta(title: Optional[str], content: str, overrides: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    """合并元数据：参数 > front matter。返回 (元数据, 去掉 front matter 的正文)。"""
    meta, body = split_front_matter(content)
    for key, value in overrides.items():
        if value is not None:
            meta[key] = value
    if title:
        meta["title"] = title
    meta.setdefault("title", "")
    return meta, body


def caption_of(block) -> Optional[Tuple[str, str]]:
    """「表：…」「图：…」这样的段落是表格或图片的说明，返回 (kind, 文字)。"""
    if not isinstance(block, mdb.Paragraph):
        return None
    match = _CAPTION.match(mdb.plain(block.spans))
    if not match:
        return None
    kind = "table" if match.group(1) in ("表", "Table") else "figure"
    return kind, match.group(2).strip()


def attach_captions(blocks: List[Any]) -> List[Tuple[Any, Optional[str]]]:
    """把紧挨在表格前后的「表：…」段落并到表格上；返回 [(block, caption)]。"""
    result: List[Tuple[Any, Optional[str]]] = []
    skip = set()
    for i, block in enumerate(blocks):
        if i in skip:
            continue
        if isinstance(block, mdb.Table):
            before = caption_of(blocks[i - 1]) if i and (i - 1) not in skip else None
            after = caption_of(blocks[i + 1]) if i + 1 < len(blocks) else None
            caption = None
            if before and before[0] == "table" and result and result[-1][0] is blocks[i - 1]:
                result.pop()
                caption = before[1]
            elif after and after[0] == "table":
                caption = after[1]
                skip.add(i + 1)
            result.append((block, caption))
        else:
            result.append((block, None))
    return result


def is_total_row(cells: List[str]) -> bool:
    return bool(cells) and bool(_TOTAL.match(cells[0]))


def records_to_rows(records: List[Dict[str, Any]]) -> List[List[Any]]:
    """对象数组（每行一个对象）→ 二维数组：表头为各对象键的并集，按首次出现的顺序。"""
    header: List[str] = []
    for record in records:
        header += [key for key in record if key not in header]
    return [list(header)] + [[record.get(key) for key in header] for record in records]


def numeric_text(text: str) -> bool:
    """像数字的文字：1,234、-5.6%、(1.3)、¥200 等。"""
    return bool(_NUMERIC.match(text))


def numeric_columns(header: List[str], rows: List[List[str]], aligns: List[Optional[str]]) -> List[int]:
    """数字列：Markdown 里标了右对齐，或除表头外的单元格全是数字。"""
    width = max([len(header)] + [len(r) for r in rows]) if (header or rows) else 0
    result = []
    for i in range(width):
        values = [r[i].strip() for r in rows if i < len(r) and r[i].strip()]
        right = i < len(aligns) and aligns[i] == "right"
        if right or (values and all(_NUMERIC.match(v) for v in values)):
            result.append(i)
    return result


def display_width(text: str) -> float:
    """估算文字宽度（以半角字符为 1）：中日韩字符算 2。"""
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in text)
