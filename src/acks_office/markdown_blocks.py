"""把 Markdown 解析成与格式无关的块结构，供 Word、PDF 渲染共用。"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Union

from markdown_it import MarkdownIt


@dataclass
class Span:
    """一段行内文字；text 里的 "\\n" 表示换行。"""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    strike: bool = False
    link: Optional[str] = None


@dataclass
class Heading:
    level: int
    spans: List[Span]


@dataclass
class Paragraph:
    spans: List[Span]


@dataclass
class ListItem:
    spans: List[Span]
    checked: Optional[bool] = None  # 任务列表：True 已完成 / False 未完成 / None 普通条目
    children: List["Block"] = field(default_factory=list)


@dataclass
class ListBlock:
    ordered: bool
    items: List[ListItem]
    start: int = 1


@dataclass
class Table:
    header: List[List[Span]]
    rows: List[List[List[Span]]]
    aligns: List[Optional[str]]


@dataclass
class Quote:
    blocks: List["Block"]
    alert: Optional[str] = None  # GitHub 提示块：note / tip / important / warning / caution


@dataclass
class Code:
    text: str
    lang: str = ""


@dataclass
class Rule:
    pass


@dataclass
class Image:
    src: str
    alt: str = ""


Block = Union[Heading, Paragraph, ListBlock, Table, Quote, Code, Rule, Image]

_ALERT = re.compile(r"^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*", re.IGNORECASE)
_TASK = re.compile(r"^\[([ xX])\]\s+")


def plain(spans: List[Span]) -> str:
    return "".join(s.text for s in spans)


def parse(text: str) -> List[Block]:
    # breaks=True：段落内的单个换行保留为换行，和旧版「一行一段」的写法观感一致
    md = MarkdownIt("commonmark", {"html": False, "breaks": True}).enable(["table", "strikethrough"])
    blocks, _ = _parse(md.parse(text or ""), 0, None)
    return blocks


def _parse(tokens, i: int, stop: Optional[str]) -> Tuple[List[Block], int]:
    blocks: List[Block] = []
    while i < len(tokens):
        t = tokens[i]
        if stop and t.type == stop:
            return blocks, i + 1
        if t.type == "heading_open":
            blocks.append(Heading(int(t.tag[1]), _inline(tokens[i + 1])))
            i += 3
        elif t.type == "paragraph_open":
            image = _only_image(tokens[i + 1])
            blocks.append(image or Paragraph(_inline(tokens[i + 1])))
            i += 3
        elif t.type in ("bullet_list_open", "ordered_list_open"):
            ordered = t.type == "ordered_list_open"
            start = int(t.attrGet("start") or 1) if ordered else 1
            items, i = _parse_list(tokens, i + 1, "ordered_list_close" if ordered else "bullet_list_close")
            blocks.append(ListBlock(ordered, items, start))
        elif t.type == "blockquote_open":
            inner, i = _parse(tokens, i + 1, "blockquote_close")
            blocks.append(_quote(inner))
        elif t.type in ("fence", "code_block"):
            blocks.append(Code(t.content.rstrip("\n"), (t.info or "").strip()))
            i += 1
        elif t.type == "hr":
            blocks.append(Rule())
            i += 1
        elif t.type == "table_open":
            table, i = _parse_table(tokens, i + 1)
            blocks.append(table)
        else:
            i += 1
    return blocks, i


def _parse_list(tokens, i: int, close: str) -> Tuple[List[ListItem], int]:
    items: List[ListItem] = []
    while i < len(tokens) and tokens[i].type != close:
        if tokens[i].type == "list_item_open":
            inner, i = _parse(tokens, i + 1, "list_item_close")
            first = inner[0] if inner and isinstance(inner[0], Paragraph) else None
            spans = first.spans if first else []
            checked = None
            if spans and _TASK.match(spans[0].text):
                checked = _TASK.match(spans[0].text).group(1).lower() == "x"
                spans = [Span(_TASK.sub("", spans[0].text, count=1), spans[0].bold, spans[0].italic,
                              spans[0].code, spans[0].strike, spans[0].link)] + spans[1:]
            items.append(ListItem(spans, checked, inner[1:] if first else inner))
        else:
            i += 1
    return items, i + 1


def _parse_table(tokens, i: int) -> Tuple[Table, int]:
    header: List[List[Span]] = []
    rows: List[List[List[Span]]] = []
    aligns: List[Optional[str]] = []
    row: Optional[List[List[Span]]] = None
    in_head = False
    while i < len(tokens) and tokens[i].type != "table_close":
        t = tokens[i]
        if t.type == "thead_open":
            in_head = True
        elif t.type == "tbody_open":
            in_head = False
        elif t.type == "tr_open":
            row = []
        elif t.type in ("th_open", "td_open"):
            style = t.attrGet("style") or ""
            if in_head:
                aligns.append(style.split(":", 1)[1].strip() if "text-align" in style else None)
            row.append(_inline(tokens[i + 1]))
        elif t.type == "tr_close":
            if in_head:
                header = row
            else:
                rows.append(row)
        i += 1
    return Table(header, rows, aligns), i + 1


def _quote(inner: List[Block]) -> Quote:
    if inner and isinstance(inner[0], Paragraph) and inner[0].spans:
        first = inner[0].spans[0]
        match = _ALERT.match(first.text)
        if match:
            rest = [Span(first.text[match.end():].lstrip("\n"), first.bold, first.italic,
                         first.code, first.strike, first.link)] + inner[0].spans[1:]
            rest = [s for s in rest if s.text]
            body = ([Paragraph(rest)] if rest else []) + inner[1:]
            return Quote(body, match.group(1).lower())
    return Quote(inner)


def _only_image(inline) -> Optional[Image]:
    children = [c for c in (inline.children or []) if not (c.type == "text" and not c.content.strip())]
    if len(children) == 1 and children[0].type == "image":
        img = children[0]
        return Image(img.attrGet("src") or "", img.content or "")
    return None


def _inline(inline) -> List[Span]:
    spans: List[Span] = []
    bold = italic = strike = False
    link: Optional[str] = None
    for c in inline.children or []:
        if c.type == "text":
            spans.append(Span(c.content, bold, italic, False, strike, link))
        elif c.type == "code_inline":
            spans.append(Span(c.content, bold, italic, True, strike, link))
        elif c.type in ("softbreak", "hardbreak"):
            spans.append(Span("\n", bold, italic, False, strike, link))
        elif c.type == "strong_open":
            bold = True
        elif c.type == "strong_close":
            bold = False
        elif c.type == "em_open":
            italic = True
        elif c.type == "em_close":
            italic = False
        elif c.type == "s_open":
            strike = True
        elif c.type == "s_close":
            strike = False
        elif c.type == "link_open":
            link = c.attrGet("href")
        elif c.type == "link_close":
            link = None
        elif c.type == "image":
            spans.append(Span(c.content or "", bold, italic, False, strike, link))
    return _merge(spans)


def _merge(spans: List[Span]) -> List[Span]:
    merged: List[Span] = []
    for s in spans:
        if not s.text:
            continue
        if merged and (merged[-1].bold, merged[-1].italic, merged[-1].code, merged[-1].strike, merged[-1].link) == \
                (s.bold, s.italic, s.code, s.strike, s.link):
            merged[-1].text += s.text
        else:
            merged.append(Span(s.text, s.bold, s.italic, s.code, s.strike, s.link))
    return merged
