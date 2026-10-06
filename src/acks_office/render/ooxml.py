"""Word 底层 XML 辅助：按 OOXML 规定的顺序插入元素，设置边框、底纹、字体、域与分栏。

python-docx 没有封装的格式（段落与单元格边框、底纹、字距、分栏、首字下沉、域）都在这里处理。
"""

from typing import Dict, Iterable, Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_SEQ = {
    "pPr": ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"),
    "rPr": ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
            "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
            "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
            "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"),
    "tcPr": ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"),
    "tblPr": ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
              "tblLook"),
    "trPr": ("cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
             "tblHeader", "tblCellSpacing", "jc", "hidden"),
    "sectPr": ("headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
               "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
               "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"),
}
BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")


def el(tag: str, **attrs) -> OxmlElement:
    element = OxmlElement(tag)
    for key, value in attrs.items():
        element.set(qn(f"w:{key}"), str(value))
    return element


def set_child(parent, child, kind: str):
    """按 schema 顺序放入子元素；同名元素已存在时替换。"""
    tag = child.tag
    old = parent.find(tag)
    if old is not None:
        parent.replace(old, child)
        return child
    order = [qn(f"w:{name}") for name in _SEQ[kind]]
    if tag in order:
        for later in order[order.index(tag) + 1:]:
            found = parent.find(later)
            if found is not None:
                found.addprevious(child)
                return child
    parent.append(child)
    return child


def eighths(points: float) -> int:
    """边框粗细：OOXML 以 1/8 磅为单位。"""
    return max(2, round(points * 8))


def twips(points: float) -> int:
    return round(points * 20)


def borders(container, kind: str, tag: str, sides: Dict[str, Optional[tuple]]) -> None:
    """设置边框。sides: {"top": (粗细磅, "RRGGBB"), "bottom": None 表示无边框}。"""
    box = container.find(qn(f"w:{tag}"))
    if box is None:
        box = set_child(container, el(f"w:{tag}"), kind)
    for side in BORDER_SIDES:
        if side not in sides:
            continue
        old = box.find(qn(f"w:{side}"))
        if old is not None:
            box.remove(old)
        spec = sides[side]
        node = el(f"w:{side}", val="nil") if spec is None else el(
            f"w:{side}", val="single", sz=eighths(spec[0]), space=spec[2] if len(spec) > 2 else 0, color=spec[1])
        # 子元素也有固定顺序：top、left、bottom、right、insideH、insideV
        order = [qn(f"w:{s}") for s in BORDER_SIDES]
        placed = False
        for later in order[order.index(node.tag) + 1:]:
            found = box.find(later)
            if found is not None:
                found.addprevious(node)
                placed = True
                break
        if not placed:
            box.append(node)


def paragraph_borders(paragraph, **sides) -> None:
    borders(paragraph._p.get_or_add_pPr(), "pPr", "pBdr", sides)


def cell_borders(cell, **sides) -> None:
    borders(cell._tc.get_or_add_tcPr(), "tcPr", "tcBorders", sides)


def shade(container, kind: str, fill: str) -> None:
    set_child(container, el("w:shd", val="clear", color="auto", fill=fill), kind)


def shade_paragraph(paragraph, fill: str) -> None:
    shade(paragraph._p.get_or_add_pPr(), "pPr", fill)


def shade_cell(cell, fill: str) -> None:
    shade(cell._tc.get_or_add_tcPr(), "tcPr", fill)


def cell_margins(cell, top: float, right: float, bottom: float, left: float) -> None:
    mar = el("w:tcMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        mar.append(el(f"w:{side}", w=twips(value), type="dxa"))
    set_child(cell._tc.get_or_add_tcPr(), mar, "tcPr")


def table_layout_fixed(table) -> None:
    set_child(table._tbl.tblPr, el("w:tblLayout", type="fixed"), "tblPr")


def table_no_borders(table) -> None:
    borders(table._tbl.tblPr, "tblPr", "tblBorders", {side: None for side in BORDER_SIDES})


def row_flags(row, header: bool = False, cant_split: bool = True) -> None:
    trpr = row._tr.get_or_add_trPr()
    if cant_split:
        set_child(trpr, el("w:cantSplit"), "trPr")
    if header:
        set_child(trpr, el("w:tblHeader"), "trPr")


def run_fonts(run_or_rpr, en: str, cn: str) -> None:
    rpr = run_or_rpr.get_or_add_rPr() if hasattr(run_or_rpr, "get_or_add_rPr") else run_or_rpr
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = set_child(rpr, el("w:rFonts"), "rPr")
    for key, value in (("ascii", en), ("hAnsi", en), ("cs", en), ("eastAsia", cn)):
        fonts.set(qn(f"w:{key}"), value)
    for theme_attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{theme_attr}"), None)


def tracking(rpr, points: float) -> None:
    """字距（每个字符之间额外的间距），单位磅。"""
    if points:
        set_child(rpr, el("w:spacing", val=twips(points)), "rPr")


def keep_lines(paragraph) -> None:
    set_child(paragraph._p.get_or_add_pPr(), el("w:keepLines"), "pPr")


def tab_stop(paragraph, position_pt: float, align: str = "right") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    tabs = ppr.find(qn("w:tabs"))
    if tabs is None:
        tabs = set_child(ppr, el("w:tabs"), "pPr")
    tabs.append(el("w:tab", val=align, pos=twips(position_pt)))


def columns(section, count: int, gap_pt: float) -> None:
    set_child(section._sectPr, el("w:cols", num=count, space=twips(gap_pt)), "sectPr")


def add_field(paragraph, instruction: str, placeholder: str = "", dirty: bool = False, style=None):
    """插入简单域（如 PAGE、PAGEREF）；style(run) 用来给域结果设置字体。"""
    field = el("w:fldSimple", instr=f" {instruction} ")
    if dirty:
        field.set(qn("w:dirty"), "true")
    run = paragraph.add_run(placeholder)
    if style:
        style(run)
    run._r.getparent().remove(run._r)
    field.append(run._r)
    paragraph._p.append(field)
    return field


def bookmark(paragraph, name: str, ident: int) -> None:
    start = el("w:bookmarkStart", id=ident, name=name)
    end = el("w:bookmarkEnd", id=ident)
    ppr = paragraph._p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(start)
    else:
        paragraph._p.insert(0, start)
    paragraph._p.append(end)


def drop_cap(paragraph, lines: int, line_pt: float) -> None:
    """把这个段落变成首字下沉的框（段落里只放下沉的那个字）。"""
    ppr = paragraph._p.get_or_add_pPr()
    set_child(ppr, el("w:framePr", dropCap="drop", lines=lines, wrap="around", vAnchor="text", hAnchor="text"),
              "pPr")
    set_child(ppr, el("w:spacing", line=twips(line_pt * lines), lineRule="exact", before=0, after=0), "pPr")


def remove_paragraph(paragraph) -> None:
    node = paragraph._p
    node.getparent().remove(node)


def iter_ids(start: int = 1) -> Iterable[int]:
    value = start
    while True:
        yield value
        value += 1
