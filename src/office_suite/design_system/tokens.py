# -*- coding: utf-8 -*-
"""
ACKS Studio · Document Design System v2.1
Python-docx token definitions and style/component factories.

Changes from v2.0:
  - New components: add_bullet_list, add_numbered_list, add_checklist,
    add_code_block. v2.0 had no list or code-block primitive at all —
    every generated document needing a bullet list had to hand-roll one
    with raw paragraph formatting, which is exactly the kind of drift
    a design-token system exists to prevent. These reuse GLYPH_BULLET /
    GLYPH_CHECK_EMPTY from tokens.json so the marker glyph is a token,
    not a hardcoded character, and match the pptx-side add_list_slide /
    add_checklist_slide / add_template_slide component-for-component.

Changes from v1.0:
  - Constants imported from tokens_constants.py (auto-generated from tokens.json)
  - FontStyle dataclass replaces scattered keyword args
  - Factory functions accept optional overrides for data reuse
  - __all__ defines the public API
  - COLOR_RULE inconsistency fixed (now D9D9DC everywhere)

REQUIREMENTS
    pip install python-docx

USAGE
    from tokens import build_document, save_demo
    save_demo("ACKS-demo.docx")

NEVER edit token values without updating tokens.json and running:
    python scripts/build_tokens.py
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Mapping, Optional, Sequence, TypedDict

from docx import Document
from docx.document import Document as _Doc
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

# Auto-generated constants from tokens.json
from .tokens_constants import *  # noqa: F401,F403

__all__ = [
    # Setup
    "apply_default_styles", "set_a4_page_margins",
    # Text components
    "add_eyebrow", "add_section_heading", "add_sub_heading",
    "add_label", "add_body_paragraph", "add_caption",
    # Lists
    "add_bullet_list", "add_numbered_list", "add_checklist", "add_code_block",
    # Callouts
    "add_callout_note", "add_callout_key", "add_quote",
    # Tables & KPI
    "add_data_table", "add_kpi_row",
    # Cover & footer
    "add_cover_page", "add_page_footer",
    # Builder
    "build_document", "save_demo",
    # Dataclass & TypedDict
    "FontStyle", "KpiItem", "MetaItem",
]


# =============================================================
#                    TYPE DEFINITIONS
# =============================================================

class KpiItem(TypedDict, total=False):
    """KPI card item for add_kpi_row()."""
    label: str
    value: str
    delta: str
    desc: str


class MetaItem(TypedDict):
    """Metadata key-value pair for cover page."""
    key: str
    value: str


# =============================================================
#                    FONT STYLE DATACLASS
# =============================================================

@dataclass(frozen=True)
class FontStyle:
    """Reusable font specification. Create presets and pass to _apply()."""
    font: str = "DM Sans"
    font_cn: str = "Noto Sans SC"
    size: Pt = Pt(10.5)
    bold: bool = False
    color: RGBColor = RGBColor(0x1A, 0x1A, 0x1A)
    uppercase: bool = False
    mono: bool = False

    def apply(self, run) -> None:
        """Apply this style to a docx Run."""
        _set_run_font(
            run,
            font=self.font,
            font_cn=self.font_cn,
            size=self.size,
            bold=self.bold,
            color=self.color,
            uppercase=self.uppercase,
            mono=self.mono,
        )


# Pre-defined style presets
STYLE_HEADING = FontStyle(font="Space Grotesk", font_cn="Space Grotesk",
                          size=Pt(24), bold=True)
STYLE_BODY = FontStyle()
STYLE_LEDE = FontStyle(size=Pt(12), bold=True)
STYLE_META = FontStyle(mono=True, size=Pt(8),
                       color=RGBColor(0x88, 0x88, 0x88))
STYLE_EYEBROW = FontStyle(mono=True, size=Pt(8),
                          color=RGBColor(0xFF, 0x6B, 0x1A), uppercase=True)


# =============================================================
#                    LOW-LEVEL HELPERS
# =============================================================

def _set_run_font(run, *, font=None, font_cn=None, size=None,
                  bold=None, color=None, uppercase=False, mono=False):
    """Apply font/size/color/bold to a single run, supporting CJK."""
    if mono:
        font = font or FONT_MONO
        font_cn = font_cn or FONT_MONO
    else:
        font = font or FONT_BODY_EN
        font_cn = font_cn or FONT_BODY_CN

    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)

    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        if isinstance(color, str):
            run.font.color.rgb = RGBColor.from_string(color)
        else:
            run.font.color.rgb = color
    if uppercase:
        caps = OxmlElement('w:caps')
        caps.set(qn('w:val'), '1')
        rPr.append(caps)


def _set_paragraph_spacing(par, *, before=None, after=None,
                           line_height=None, alignment=None,
                           keep_with_next=False):
    pf = par.paragraph_format
    if before is not None: pf.space_before = before
    if after is not None: pf.space_after = after
    if line_height is not None:
        pf.line_spacing = line_height
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if alignment is not None:
        par.alignment = alignment
    if keep_with_next:
        pf.keep_with_next = True


def _set_cell_shading(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _set_cell_borders(cell, *, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, spec in (('top', top), ('bottom', bottom),
                       ('left', left), ('right', right)):
        if spec is None:
            continue
        size, color = spec
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), str(size))
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)


# =============================================================
#                    DOCUMENT SETUP
# =============================================================

def apply_default_styles(doc: _Doc) -> None:
    """Re-bind the document's Normal style to ACKS defaults."""
    style = doc.styles['Normal']
    f = style.font
    f.name = FONT_BODY_EN
    f.size = FONT_BODY_SIZE
    f.color.rgb = RGB_INK

    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), FONT_BODY_CN)
    rFonts.set(qn('w:ascii'), FONT_BODY_EN)
    rFonts.set(qn('w:hAnsi'), FONT_BODY_EN)

    pf = style.paragraph_format
    pf.space_after = PARAGRAPH_SPACE_AFTER
    pf.line_spacing = LINE_HEIGHT_CN
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE


def set_a4_page_margins(doc: _Doc) -> None:
    """Set every section to A4 portrait with ACKS margins."""
    for section in doc.sections:
        section.page_height = PAGE_HEIGHT
        section.page_width  = PAGE_WIDTH
        section.top_margin    = PAGE_MARGIN_TOP
        section.bottom_margin = PAGE_MARGIN_BOTTOM
        section.left_margin   = PAGE_MARGIN_LEFT
        section.right_margin  = PAGE_MARGIN_RIGHT
        section.header_distance = PAGE_HEADER_DIST
        section.footer_distance = PAGE_FOOTER_DIST


# =============================================================
#                    TEXT COMPONENTS
# =============================================================

def add_eyebrow(doc: _Doc, text: str):
    """Small UPPERCASE mono label above section titles."""
    par = doc.add_paragraph()
    run = par.add_run(text.upper())
    STYLE_EYEBROW.apply(run)
    _set_paragraph_spacing(par, before=SPACE_16, after=SPACE_4,
                           keep_with_next=True)
    return par


def add_section_heading(doc: _Doc, en: str, zh: Optional[str] = None):
    """H2 — English title + optional Chinese sub-title."""
    par = doc.add_paragraph()
    run = par.add_run(en)
    _set_run_font(run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=FONT_H2_SIZE, bold=True, color=RGB_INK)
    _set_paragraph_spacing(par, before=SPACE_24, after=SPACE_4,
                           line_height=LINE_HEIGHT_TIGHT,
                           keep_with_next=True)
    if zh:
        sub = doc.add_paragraph()
        sub_run = sub.add_run(zh)
        _set_run_font(sub_run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                      size=Pt(14), bold=True, color=RGB_INK_2)
        _set_paragraph_spacing(sub, before=Pt(0), after=SPACE_12,
                               line_height=LINE_HEIGHT_TIGHT)
    return par


def add_sub_heading(doc: _Doc, en: str, zh: Optional[str] = None,
                    tag: Optional[str] = None):
    """H3 — 16pt with a 1pt black underline."""
    par = doc.add_paragraph()
    run = par.add_run(en)
    _set_run_font(run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=FONT_H3_SIZE, bold=True, color=RGB_INK)
    if zh:
        zh_run = par.add_run("  " + zh)
        _set_run_font(zh_run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                      size=Pt(10), color=RGB_INK_2)
    if tag:
        tag_run = par.add_run("    " + tag.upper())
        _set_run_font(tag_run, mono=True, size=FONT_META_SIZE,
                      color=RGB_PRIMARY, uppercase=True)
    _set_paragraph_spacing(par, before=SPACE_16, after=SPACE_4,
                           keep_with_next=True)
    _add_paragraph_border(par, side='bottom', size=8, color=COLOR_TEXT_PRIMARY)
    return par


def add_label(doc: _Doc, text: str):
    """H4 — small UPPERCASE label inside body."""
    par = doc.add_paragraph()
    run = par.add_run(text.upper())
    _set_run_font(run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=FONT_H4_SIZE, bold=True, color=RGB_INK,
                  uppercase=True)
    _set_paragraph_spacing(par, before=SPACE_12, after=SPACE_4,
                           keep_with_next=True)
    return par


def add_body_paragraph(doc: _Doc, text: str,
                       *, justify: bool = True,
                       lede: bool = False,
                       color: Optional[RGBColor] = None,
                       style: Optional[FontStyle] = None):
    """Body paragraph. Pass `style` to override defaults."""
    par = doc.add_paragraph()
    run = par.add_run(text)
    if style:
        style.apply(run)
    else:
        _set_run_font(
            run,
            font=FONT_BODY_EN, font_cn=FONT_BODY_CN,
            size=Pt(12) if lede else FONT_BODY_SIZE,
            bold=lede,
            color=color or RGB_INK,
        )
    _set_paragraph_spacing(
        par,
        before=PARAGRAPH_SPACE_BEFORE,
        after=SPACE_12 if lede else PARAGRAPH_SPACE_AFTER,
        line_height=LINE_HEIGHT_CN,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY if justify else None,
    )
    return par


def add_caption(doc: _Doc, fig_no: str, text: str):
    """Figure / table caption."""
    par = doc.add_paragraph()
    n = par.add_run(fig_no.upper() + "   ")
    _set_run_font(n, mono=True, size=FONT_CAPTION_SIZE,
                  color=RGB_PRIMARY, uppercase=True)
    t = par.add_run(text)
    _set_run_font(t, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                  size=FONT_CAPTION_SIZE, color=RGB_INK_2)
    _set_paragraph_spacing(par, before=SPACE_4, after=SPACE_12,
                           line_height=1.6)
    return par


# =============================================================
#                    LISTS & CODE BLOCKS
#
#  v2.0 exported no bullet/numbered-list or code-block component even
#  though these are some of the most common shapes real content takes.
#  Markers come from GLYPH_BULLET / GLYPH_CHECK_EMPTY (tokens.json ->
#  glyph) rather than being hardcoded here, so docx and pptx render the
#  same marker character by construction.
# =============================================================

def add_bullet_list(doc: _Doc, items: Sequence[str]):
    """GLYPH_BULLET-prefixed list, matching add_list_slide() in slides.py."""
    pars = []
    for item in items:
        par = doc.add_paragraph()
        m = par.add_run(GLYPH_BULLET + " ")
        _set_run_font(m, mono=True, size=Pt(10.5), color=RGB_PRIMARY, bold=True)
        t = par.add_run(item)
        _set_run_font(t, font=FONT_BODY_EN, font_cn=FONT_BODY_CN,
                      size=FONT_BODY_SIZE, color=RGB_INK)
        _set_paragraph_spacing(par, before=Pt(0), after=Pt(4),
                               line_height=LINE_HEIGHT_CN)
        par.paragraph_format.left_indent = Pt(14)
        par.paragraph_format.first_line_indent = Pt(-14)
        pars.append(par)
    return pars


def add_numbered_list(doc: _Doc, items: Sequence[str]):
    """Mono-numbered list, 01 / 02 / 03 ..."""
    pars = []
    for i, item in enumerate(items, start=1):
        par = doc.add_paragraph()
        n = par.add_run(f"{i:02d}  ")
        _set_run_font(n, mono=True, size=Pt(10), color=RGB_PRIMARY_DEEP, bold=True)
        t = par.add_run(item)
        _set_run_font(t, font=FONT_BODY_EN, font_cn=FONT_BODY_CN,
                      size=FONT_BODY_SIZE, color=RGB_INK)
        _set_paragraph_spacing(par, before=Pt(0), after=Pt(4),
                               line_height=LINE_HEIGHT_CN)
        par.paragraph_format.left_indent = Pt(20)
        par.paragraph_format.first_line_indent = Pt(-20)
        pars.append(par)
    return pars


def add_checklist(doc: _Doc, items: Sequence[str]):
    """GLYPH_CHECK_EMPTY checklist, matching add_checklist_slide() in slides.py."""
    pars = []
    for item in items:
        par = doc.add_paragraph()
        box = par.add_run(GLYPH_CHECK_EMPTY + "  ")
        _set_run_font(box, size=Pt(11), color=RGB_PRIMARY)
        t = par.add_run(item)
        _set_run_font(t, font=FONT_BODY_EN, font_cn=FONT_BODY_CN,
                      size=FONT_BODY_SIZE, color=RGB_INK)
        _set_paragraph_spacing(par, before=Pt(0), after=Pt(8),
                               line_height=1.6)
        pars.append(par)
    return pars


def add_code_block(doc: _Doc, text: str, *, caption: Optional[str] = None):
    """Monospace box for prompt templates / config snippets.
    Preserves line breaks; light grey fill, full border, orange top rule —
    matches add_template_slide() in slides.py.
    """
    lines = text.strip("\n").split("\n")
    par = doc.add_paragraph()
    for i, line in enumerate(lines):
        run = par.add_run(line if line else " ")
        _set_run_font(run, mono=True, size=Pt(9.5), color=RGB_INK)
        if i < len(lines) - 1:
            run.add_break()
    _set_paragraph_spacing(par, before=Pt(6), after=Pt(4), line_height=1.55)
    _add_paragraph_shading(par, COLOR_BG_TERTIARY)
    _add_paragraph_border(par, side="top", size=16, color=COLOR_PRIMARY)
    _add_paragraph_border(par, side="left", size=4, color=COLOR_RULE)
    _add_paragraph_border(par, side="right", size=4, color=COLOR_RULE)
    _add_paragraph_border(par, side="bottom", size=4, color=COLOR_RULE)
    par.paragraph_format.left_indent = Pt(10)
    par.paragraph_format.right_indent = Pt(10)
    if caption:
        cp = doc.add_paragraph()
        cr = cp.add_run(caption)
        _set_run_font(cr, mono=True, size=FONT_META_SIZE, color=RGB_INK_3,
                      uppercase=True)
        _set_paragraph_spacing(cp, before=Pt(3), after=Pt(10))
    return par


# =============================================================
#                    CALLOUTS
# =============================================================

def _add_paragraph_border(par, *, side='left', size=24, color="1A1A1A"):
    pPr = par._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    b = OxmlElement(f'w:{side}')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), str(size))
    b.set(qn('w:space'), '4')
    b.set(qn('w:color'), color)
    pBdr.append(b)


def _add_paragraph_shading(par, hex_color: str):
    pPr = par._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)


def _callout(doc: _Doc, label: str, text: str, *,
             fill: str, border_color: str, label_color: RGBColor):
    par = doc.add_paragraph()
    lbl_run = par.add_run(label.upper() + "\n")
    _set_run_font(lbl_run, mono=True, size=FONT_META_SIZE,
                  color=label_color, uppercase=True)
    body_run = par.add_run(text)
    _set_run_font(body_run, font=FONT_BODY_EN, font_cn=FONT_BODY_CN,
                  size=Pt(10), color=RGB_INK)
    _set_paragraph_spacing(par, before=SPACE_8, after=SPACE_8,
                           line_height=1.7)
    _add_paragraph_shading(par, fill)
    _add_paragraph_border(par, side='left', size=24, color=border_color)
    par.paragraph_format.left_indent = Pt(12)
    par.paragraph_format.right_indent = Pt(12)
    return par


def add_callout_note(doc: _Doc, text: str, label: str = "Note · 注释"):
    """Grey-bg callout with black left bar."""
    return _callout(doc, label, text,
                    fill=COLOR_BG_SECONDARY,
                    border_color=COLOR_TEXT_PRIMARY,
                    label_color=RGB_INK_3)


def add_callout_key(doc: _Doc, text: str, label: str = "Key Insight · 关键洞察"):
    """Orange-bg callout with orange left bar."""
    return _callout(doc, label, text,
                    fill=COLOR_PRIMARY_SOFT,
                    border_color=COLOR_PRIMARY,
                    label_color=RGB_PRIMARY_DEEP)


def add_quote(doc: _Doc, text: str, source: Optional[str] = None):
    """Serif quote with top/bottom rules and orange open-quote."""
    par = doc.add_paragraph()
    oq = par.add_run("\u201C")
    _set_run_font(oq, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=Pt(20), color=RGB_PRIMARY, bold=True)
    body = par.add_run(text)
    _set_run_font(body, font=FONT_SERIF_CN, font_cn=FONT_SERIF_CN,
                  size=Pt(12), color=RGB_INK)
    _set_paragraph_spacing(par, before=SPACE_12, after=SPACE_4,
                           line_height=1.65)
    _add_paragraph_border(par, side='top', size=8, color=COLOR_TEXT_PRIMARY)
    if source:
        src_par = doc.add_paragraph()
        src_run = src_par.add_run(f"— {source}")
        _set_run_font(src_run, mono=True, size=FONT_META_SIZE,
                      color=RGB_INK_3)
        _set_paragraph_spacing(src_par, before=Pt(0), after=SPACE_12,
                               line_height=1.5)
        _add_paragraph_border(src_par, side='bottom', size=8,
                              color=COLOR_TEXT_PRIMARY)
    else:
        _add_paragraph_border(par, side='bottom', size=8,
                              color=COLOR_TEXT_PRIMARY)
    return par


# =============================================================
#                    TABLES & KPI ROW
# =============================================================

def add_data_table(doc: _Doc,
                   headers: Sequence[str],
                   rows: Sequence[Sequence[str]],
                   *,
                   numeric_cols: Sequence[int] = (),
                   footer_row: Optional[Sequence[str]] = None,
                   col_widths_mm: Optional[Sequence[float]] = None,
                   ) -> None:
    """Data table with black header / zebra rows / shaded footer."""
    n_cols = len(headers)
    n_rows = 1 + len(rows) + (1 if footer_row else 0)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    if col_widths_mm:
        for i, w in enumerate(col_widths_mm):
            for row in table.rows:
                row.cells[i].width = Mm(w)

    # HEADER
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        _set_run_font(run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                      size=Pt(9), bold=True, color=RGB_WHITE)
        _set_paragraph_spacing(p, before=Pt(2), after=Pt(2))
        if i in numeric_cols:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_shading(cell, COLOR_TEXT_PRIMARY)

    # DATA ROWS
    for ri, row_data in enumerate(rows, start=1):
        is_last = (ri == len(rows))
        is_even = (ri % 2 == 0)
        for ci, val in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            if ci in numeric_cols:
                _set_run_font(run, mono=True, size=Pt(9), color=RGB_INK)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                _set_run_font(run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                              size=Pt(9), color=RGB_INK)
            _set_paragraph_spacing(p, before=Pt(2), after=Pt(2))
            if is_even:
                _set_cell_shading(cell, COLOR_BG_SECONDARY)
            border_color = (COLOR_TEXT_PRIMARY if is_last else COLOR_RULE)
            border_size = 16 if is_last else 4
            _set_cell_borders(cell, bottom=(border_size, border_color))

    # FOOTER
    if footer_row:
        for ci, val in enumerate(footer_row):
            cell = table.rows[-1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            if ci in numeric_cols:
                _set_run_font(run, mono=True, size=Pt(9), bold=True,
                              color=RGB_PRIMARY_DEEP)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                _set_run_font(run, font=FONT_DISPLAY_EN,
                              font_cn=FONT_BODY_CN,
                              size=Pt(9), bold=True, color=RGB_INK)
            _set_paragraph_spacing(p, before=Pt(3), after=Pt(3))
            _set_cell_shading(cell, COLOR_BG_TERTIARY)


def add_kpi_row(doc: _Doc,
                kpis: Sequence[Mapping[str, str]]) -> None:
    """Build a 1-row table of KPI cards.

    Each item: {"label": "Revenue", "value": "¥420M",
                "delta": "+22%", "desc": "vs 2025"}
    """
    n = len(kpis)
    table = doc.add_table(rows=1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, item in enumerate(kpis):
        cell = table.rows[0].cells[i]
        cell.text = ""
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        # label
        p_lbl = cell.paragraphs[0]
        run = p_lbl.add_run(item.get("label", "").upper())
        _set_run_font(run, mono=True, size=Pt(7),
                      color=RGB_INK_3, uppercase=True)
        _set_paragraph_spacing(p_lbl, before=Pt(6), after=Pt(2))
        # value
        p_v = cell.add_paragraph()
        v_run = p_v.add_run(item.get("value", ""))
        _set_run_font(v_run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                      size=Pt(22), bold=True, color=RGB_INK)
        _set_paragraph_spacing(p_v, before=Pt(0), after=Pt(2))
        # delta
        if item.get("delta"):
            p_d = cell.add_paragraph()
            d_run = p_d.add_run(item["delta"])
            _set_run_font(d_run, mono=True, size=Pt(8), color=RGB_PRIMARY)
            _set_paragraph_spacing(p_d, before=Pt(0), after=Pt(2))
        # description
        if item.get("desc"):
            p_de = cell.add_paragraph()
            de_run = p_de.add_run(item["desc"])
            _set_run_font(de_run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                          size=Pt(8), color=RGB_INK_3)
            _set_paragraph_spacing(p_de, before=Pt(0), after=Pt(6))
        # borders
        if i < n - 1:
            _set_cell_borders(cell, right=(4, COLOR_RULE))
        _set_cell_borders(cell, top=(8, COLOR_TEXT_PRIMARY),
                          bottom=(8, COLOR_TEXT_PRIMARY),
                          right=(4, COLOR_RULE) if i < n - 1 else None)


# =============================================================
#                    COVER + HEADER/FOOTER
# =============================================================

def add_cover_page(doc: _Doc, *,
                   title_top: str,
                   title_em: str,
                   zh_title: str,
                   doctype: str,
                   lede: Optional[str] = None,
                   meta: Optional[Mapping[str, str]] = None,
                   brand_name: str = "ACKS STUDIO · 爱驰科驶") -> None:
    """Build a cover page.

    `brand_name` is the top stamp text; pass "" or a custom brand to
    de-brand / re-brand (office_suite D1 integration).
    """
    stamp = doc.add_paragraph()
    run = stamp.add_run(brand_name)
    _set_run_font(run, mono=True, size=Pt(9), color=RGB_INK_2,
                  uppercase=True)
    _set_paragraph_spacing(stamp, before=Pt(0), after=Pt(48))

    if doctype:
        eb = doc.add_paragraph()
        eb_run = eb.add_run(f"— {doctype}")
        _set_run_font(eb_run, mono=True, size=Pt(9), color=RGB_PRIMARY,
                      uppercase=True)
        _set_paragraph_spacing(eb, before=Pt(0), after=Pt(8))

    if title_top or title_em:
        t1 = doc.add_paragraph()
        if title_top:
            t1_a = t1.add_run(title_top + " ")
            _set_run_font(t1_a, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                          size=FONT_H1_SIZE, bold=True, color=RGB_INK)
        if title_em:
            t1_b = t1.add_run(title_em)
            _set_run_font(t1_b, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                          size=FONT_H1_SIZE, bold=True, color=RGB_PRIMARY)
        _set_paragraph_spacing(t1, before=Pt(0), after=Pt(8),
                               line_height=LINE_HEIGHT_DISPLAY)

    if zh_title:
        zh = doc.add_paragraph()
        zh_run = zh.add_run(zh_title)
        _set_run_font(zh_run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                      size=Pt(20), bold=True, color=RGB_INK)
        _set_paragraph_spacing(zh, before=Pt(8), after=Pt(12))

    if lede:
        ld = doc.add_paragraph()
        ld_run = ld.add_run(lede)
        _set_run_font(ld_run, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                      size=Pt(11), color=RGB_INK_2)
        _set_paragraph_spacing(ld, before=Pt(0), after=Pt(24),
                               line_height=1.7)

    if meta:
        n = len(meta)
        table = doc.add_table(rows=2, cols=n)
        for i, (k, v) in enumerate(meta.items()):
            cell_k = table.rows[0].cells[i]
            cell_k.text = ""
            p = cell_k.paragraphs[0]
            r = p.add_run(k.upper())
            _set_run_font(r, mono=True, size=Pt(8),
                          color=RGB_INK_3, uppercase=True)
            _set_paragraph_spacing(p, before=Pt(8), after=Pt(2))
            cell_v = table.rows[1].cells[i]
            cell_v.text = ""
            pv = cell_v.paragraphs[0]
            rv = pv.add_run(v)
            _set_run_font(rv, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                          size=Pt(12), bold=True, color=RGB_INK)
            _set_paragraph_spacing(pv, before=Pt(0), after=Pt(8))
            _set_cell_borders(cell_k, top=(8, COLOR_TEXT_PRIMARY))

    doc.add_page_break()


def add_page_footer(doc: _Doc, *, label: str = "ACKS Studio · v2.0 · 2026"):
    """Wire a simple footer with PAGE / NUMPAGES."""
    section = doc.sections[0]
    footer = section.footer
    par = footer.paragraphs[0]
    par.text = ""
    l = par.add_run(label)
    _set_run_font(l, mono=True, size=FONT_META_SIZE,
                  color=RGB_INK_3, uppercase=True)
    par.add_run("\t\t")
    acc = par.add_run("▪  ")
    _set_run_font(acc, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=Pt(10), color=RGB_PRIMARY)
    _add_page_field(par)
    of = par.add_run(" / ")
    _set_run_font(of, mono=True, size=FONT_META_SIZE, color=RGB_INK_3)
    _add_numpages_field(par)
    pf = par.paragraph_format
    from docx.enum.text import WD_TAB_ALIGNMENT
    pf.tab_stops.add_tab_stop(Mm(166), WD_TAB_ALIGNMENT.RIGHT)


def _add_page_field(par):
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.text = "PAGE"
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'end')
    run = par.add_run()
    _set_run_font(run, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=Pt(11), bold=True, color=RGB_INK)
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


def _add_numpages_field(par):
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.text = "NUMPAGES"
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'end')
    run = par.add_run()
    _set_run_font(run, mono=True, size=FONT_META_SIZE, color=RGB_INK_3)
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


# =============================================================
#                    DEMO BUILD
# =============================================================

def build_document(
    title_top: str = "State of",
    title_em: str = "the Drive",
    zh_title: str = "2026 年第一季度业务回顾",
    doctype: str = "Quarterly Report · 2026 Q1 · 季度报告",
    lede: str = ("汽车改装、AI 智能体、软件服务三条业务线在 2026 年开局的整体表现；"
                 "附 5 章节、12 张图表、3 项关键洞察。"),
    kpis: Optional[Sequence[Mapping[str, str]]] = None,
    table_headers: Optional[Sequence[str]] = None,
    table_rows: Optional[Sequence[Sequence[str]]] = None,
    table_footer: Optional[Sequence[str]] = None,
) -> _Doc:
    """Construct a fully-styled demo document.

    All data parameters have sensible defaults but can be overridden
    for reuse with real data.
    """
    doc = Document()
    apply_default_styles(doc)
    set_a4_page_margins(doc)
    add_page_footer(doc)

    add_cover_page(
        doc,
        title_top=title_top,
        title_em=title_em,
        zh_title=zh_title,
        doctype=doctype,
        lede=lede,
        meta={
            "Issued":   "2026.05.18",
            "Period":   "Q1 · Jan–Mar",
            "Code":     "ACKS-R-026-Q1",
            "Class":    "CONFIDENTIAL",
        },
    )

    add_eyebrow(doc, "01 · Executive Summary")
    add_section_heading(doc, "Executive Summary", "执行摘要 — 一页看完 Q1")
    add_body_paragraph(
        doc,
        "2026 年第一季度整体超出年初制定的 Q1 目标 6.4 个百分点。"
        "性能改装业务持续放量，AI 智能体业务首次跨越月营收 ¥10M 的临界点。",
        lede=True,
    )

    if kpis is None:
        kpis = [
            {"label": "Total Revenue", "value": "¥420M",
             "delta": "+22.0% YoY", "desc": "较 2025 Q1 增加 ¥75.6M"},
            {"label": "Gross Margin", "value": "38.6%",
             "delta": "+2.4 pts", "desc": "高端线占比提升驱动"},
            {"label": "Active Clients", "value": "1,284",
             "delta": "+18% QoQ", "desc": "B2B 占比上升至 38%"},
            {"label": "NPS Score", "value": "62",
             "delta": "+1 pt", "desc": "交付速度仍是首要抱怨"},
        ]
    add_kpi_row(doc, kpis)

    add_sub_heading(doc, "Revenue by Business Line",
                    zh="分业务线季度营收", tag="Section 1.2")

    if table_headers is None:
        table_headers = ["编号", "业务线", "2024", "2025", "同比"]
    if table_rows is None:
        table_rows = [
            ["01", "性能改装",      "462", "588", "+27.3%"],
            ["02", "外观空气动力",  "218", "274", "+25.7%"],
            ["03", "制动系统",      "156", "183", "+17.3%"],
            ["04", "轮毂悬挂",      "312", "355", "+13.8%"],
        ]
    if table_footer is None:
        table_footer = ["合计", "", "1,148", "1,400", "+22.0%"]

    add_data_table(
        doc,
        headers=table_headers,
        numeric_cols=(2, 3, 4),
        rows=table_rows,
        footer_row=table_footer,
    )
    add_caption(doc, "TBL · 02", "单位：亿元；同比基于年化口径。")

    add_callout_key(
        doc,
        "高端品牌 + 内容运营 + 智能体客服是未来 18 个月最重要的差异化组合，"
        "缺一不可。三者协同的客户 LTV 比单点接触高出 2.7 倍。",
        label="KEY · 关键洞察 01",
    )
    add_callout_note(
        doc,
        "本季部分数据由门店系统直采，未经财务团队复核；"
        "Q2 起将引入 BI 看板实现 T+1 自动核对。",
        label="NOTE · 注意",
    )
    add_quote(
        doc,
        "改装不再只是机械的事情，而是一种生活方式的表达。",
        source="ACKS Studio Brand Manifesto, 2026",
    )

    add_sub_heading(doc, "Q2 Priorities", zh="Q2 优先级", tag="Section 1.3")
    add_bullet_list(doc, [
        "高端门店品牌化改造扩展到 3 个试点城市之外",
        "内容团队月更频率提升到 8 条以上",
        "智能体客服转人工率降到 15% 以下",
    ])
    add_label(doc, "落地步骤")
    add_numbered_list(doc, [
        "锁定试点城市并完成物料清单",
        "内容日历排到 Q2 最后一周",
        "客服话术库补齐高频问题的前 20 条",
    ])
    add_code_block(
        doc,
        "角色：门店品牌顾问\n"
        "背景：[城市]门店完成品牌化改造，需要给店长一份开业前检查单\n"
        "任务：1. 列出必须核对的品牌一致性项 2. 给出验收标准\n"
        "格式：表格，列 = 检查项 / 标准 / 状态",
        caption="Template · 门店品牌化交接",
    )
    add_checklist(doc, [
        "试点城市清单已确认？",
        "内容日历是否覆盖到 Q2 最后一周？",
        "客服话术库是否已补齐高频问题？",
    ])

    return doc


def save_demo(path: str = "ACKS-demo.docx") -> str:
    doc = build_document()
    doc.save(path)
    return path


if __name__ == "__main__":
    out = save_demo()
    print(f"Saved → {out}")
