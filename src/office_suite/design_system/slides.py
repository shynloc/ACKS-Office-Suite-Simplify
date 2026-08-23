# -*- coding: utf-8 -*-
"""
ACKS Studio · Presentation Design System v2.1
python-pptx token definitions and slide-layout factory functions.

Changes from v2.0:
  - Fixed: large display titles (SIZE_H1/SECTION/H1_BODY) no longer overlap
    the subtitle/rule/footer when the copy runs longer than the v2.0 demo
    strings — title boxes now reserve safe two-line height and font size
    is chosen via fit_title_size() instead of a fixed constant.
  - Fixed: add_section_slide's chapter badge used the slide's page number
    instead of the chapter number (only matched by coincidence in the demo
    deck, where chapter == page). Chapter number is now derived from `n`
    or passed explicitly via `chapter_no`.
  - Fixed: add_compare_slide's option-name sizing did a raw division of an
    Emu length by an int (not a unit conversion) and could silently wrap
    onto the item list below it. Column geometry is now computed in float
    inches throughout and item rows scale with len(items).
  - Fixed: footer page totals were a hardcoded default (8) baked in at
    render time. finalize_footers() patches every slide's "/NN" after the
    full deck is built, so callers no longer have to pre-declare a total.
  - New components: add_list_slide (bullet/numbered), add_checklist_slide,
    add_template_slide (monospace prompt/code block) — v2.0 had no list or
    code-block primitive on the pptx side even though these are some of
    the most common content shapes in a working deck.
  - New: fit_title_size() — conservative glyph-width based single-line fit
    with a documented per-font ratio table, used by every headline-style
    component. See tokens.json → content_limit for the char budgets this
    is calibrated against.

REQUIREMENTS
    pip install python-pptx

USAGE
    from slides import save_demo
    save_demo("ACKS-deck.pptx")

NEVER edit token values without updating tokens.json and running:
    python scripts/build_tokens.py
"""

from __future__ import annotations
from typing import Mapping, Optional, Sequence, TypedDict

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# Auto-generated constants from tokens.json
from .slides_constants import *  # noqa: F401,F403

__all__ = [
    "init_presentation", "finalize_footers", "fit_title_size", "fit_paragraph_size",
    "add_title_slide", "add_section_slide", "add_content_slide",
    "add_data_slide", "add_compare_slide", "add_quote_slide",
    "add_closing_slide",
    "add_list_slide", "add_checklist_slide", "add_template_slide",
    "build_presentation", "save_demo",
    "SlideKpiItem", "SlideContactItem", "SlideOptionItem", "SlideListItem",
]

# Safe content area shared by every slide: content must not render below
# this line, which leaves clearance above the footer regardless of slide
# type. Components that stack rows (lists, panels, tables) should size
# their rows against CONTENT_BUDGET = FOOTER_SAFE_TOP - CONTENT_TOP.
CONTENT_TOP = Inches(2.35)
FOOTER_SAFE_TOP = Inches(6.75)


class SlideKpiItem(TypedDict, total=False):
    """KPI item for add_data_slide()."""
    label: str
    value: str
    delta: str


class SlideContactItem(TypedDict):
    """Contact item for add_closing_slide()."""
    label: str
    value: str


class SlideOptionItem(TypedDict, total=False):
    """Option item for add_compare_slide()."""
    tag: str
    name: str
    items: Sequence[str]
    lead: bool


class SlideListItem(TypedDict, total=False):
    """Row item for add_list_slide() / add_checklist_slide()."""
    text: str


# =============================================================
#                    LOW-LEVEL HELPERS
# =============================================================

def _set_solid_fill(shape, rgb: RGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb

def _set_no_line(shape):
    shape.line.fill.background()

def _set_line(shape, rgb: RGBColor, width: Pt):
    shape.line.color.rgb = rgb
    shape.line.width = width

def _add_rect(slide, left, top, width, height, *, fill=None, line=None,
              line_width=Pt(0.75)):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill is None:
        rect.fill.background()
    else:
        _set_solid_fill(rect, fill)
    if line is None:
        _set_no_line(rect)
    else:
        _set_line(rect, line, line_width)
    return rect

def _add_text(slide, left, top, width, height, text, *,
              font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
              size=SIZE_BODY, bold=False,
              color=RGB_TEXT_PRIMARY,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              caps=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    p = tf.paragraphs[0]
    p.alignment = align
    if isinstance(text, str):
        runs = [(text, {})]
    else:
        runs = text
    for i, (txt, opts) in enumerate(runs):
        run = p.add_run() if i > 0 or p.runs else (p.runs[0] if p.runs else p.add_run())
        if not p.runs:
            run = p.add_run()
        run.text = txt
        f = run.font
        f.name = opts.get("font", font)
        f.size = opts.get("size", size)
        f.bold = opts.get("bold", bold)
        c = opts.get("color", color)
        f.color.rgb = c
        _set_east_asia_font(run, opts.get("font_cn", font_cn))
        if opts.get("caps", caps):
            _apply_caps(run)
    return tb

def _add_paragraph(tf, text, *, font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                   size=SIZE_BODY, bold=False,
                   color=RGB_TEXT_PRIMARY,
                   align=PP_ALIGN.LEFT, space_before=Pt(0), space_after=Pt(0),
                   caps=False):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = space_before
    p.space_after = space_after
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = font
    f.size = size
    f.bold = bold
    f.color.rgb = color
    _set_east_asia_font(run, font_cn)
    if caps:
        _apply_caps(run)
    return p

def _set_east_asia_font(run, font_name):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font_name)

def _apply_caps(run):
    rPr = run._r.get_or_add_rPr()
    rPr.set('cap', 'all')


# =============================================================
#                    TITLE AUTO-FIT
# =============================================================
#
# python-pptx has a hard dependency on Pillow (for image sizing), so
# Pillow is guaranteed to be importable everywhere slides.py runs —
# fit_title_size() uses real font-metric measurement via PIL rather
# than a hand-tuned "average glyph width" ratio. A per-character ratio
# heuristic looks fine on the string it was tuned against and then
# quietly fails on the next one (e.g. "¥420M" is short but glyph-heavy:
# a ratio calibrated on prose text under-shrinks it). Measuring is the
# only way to actually know.
#
# Because none of DISPLAY/BODY/MONO are guaranteed installed on the
# rendering machine, measurement uses whatever bold sans/mono the local
# fontconfig would substitute — i.e. it measures the same approximation
# LibreOffice/PowerPoint will actually render, instead of guessing at
# Space Grotesk's true metrics and hoping the substitute matches.

import functools
import shutil
import subprocess

try:
    from PIL import ImageFont as _ImageFont
except Exception:  # pragma: no cover — Pillow ships with python-pptx
    _ImageFont = None

# Hardcoded fallback candidates, tried only if fontconfig (fc-match)
# isn't on PATH (e.g. Windows) or resolution fails outright.
_FONT_FILE_CANDIDATES = {
    FONT_DISPLAY_EN: ["SpaceGrotesk-Bold.ttf", "Space Grotesk Bold.ttf"],
    FONT_BODY_EN:    ["DMSans-Bold.ttf", "DM Sans Bold.ttf"],
    FONT_MONO:       ["JetBrainsMono-Bold.ttf", "JetBrains Mono Bold.ttf"],
}
_MEASURE_REF_SIZE = 200  # px; large reference size, widths scale linearly


def _fc_match_font_file(font: str) -> Optional[str]:
    """Ask fontconfig what file it would actually substitute for `font`
    (bold weight). This is what LibreOffice's renderer resolves to on
    Linux, so measuring against *that* file — rather than guessing a
    plausible-sounding fallback — is what keeps fit_title_size() honest
    when the real family isn't installed. Returns None if fc-match
    isn't available (non-Linux) or the query fails.
    """
    if shutil.which("fc-match") is None:
        return None
    try:
        out = subprocess.run(
            ["fc-match", "-f", "%{file}", f"{font}:bold"],
            capture_output=True, text=True, timeout=3,
        )
        path = out.stdout.strip()
        return path or None
    except Exception:
        return None


@functools.lru_cache(maxsize=None)
def _load_measure_font(font: str):
    if _ImageFont is None:
        return None
    resolved = _fc_match_font_file(font)
    if resolved:
        try:
            return _ImageFont.truetype(resolved, _MEASURE_REF_SIZE)
        except Exception:
            pass
    for path_or_name in _FONT_FILE_CANDIDATES.get(font, []):
        try:
            return _ImageFont.truetype(path_or_name, _MEASURE_REF_SIZE)
        except Exception:
            continue
    return None


# Fallback heuristic (used only if PIL/fontconfig measurement fails
# entirely) — deliberately generous per-glyph-width-as-fraction-of-em.
_GLYPH_WIDTH_EM = {
    FONT_DISPLAY_EN: 0.66,
    FONT_BODY_EN:    0.60,
    FONT_MONO:       0.62,
}
_CJK_WIDTH_EM = 1.05


def _measure_width_in(text: str, size_pt: float, font: str) -> Optional[float]:
    fnt = _load_measure_font(font)
    if fnt is None:
        return None
    try:
        variant = fnt.font_variant(size=int(_MEASURE_REF_SIZE))
        bbox = variant.getbbox(text)
        width_px_at_ref = bbox[2] - bbox[0]
    except Exception:
        return None
    width_em = width_px_at_ref / _MEASURE_REF_SIZE
    return width_em * size_pt / 72.0


def fit_title_size(text: str, box_width_in: float, base_pt: float,
                   min_pt: float, *, font: str = FONT_DISPLAY_EN,
                   step: float = 2.0) -> float:
    """Return the largest size <= base_pt (down to min_pt) that keeps
    `text` on a single line within box_width_in.

    Measures real glyph advances via PIL/fontconfig when available
    (see module docstring above) and falls back to a conservative
    per-character ratio only if that measurement isn't possible. Use
    this for every large display headline instead of a hardcoded token
    size — token sizes (SIZE_H1 etc.) are the *ceiling*, not a
    guarantee, once real copy replaces demo copy.
    """
    size = base_pt
    while size > min_pt:
        measured = _measure_width_in(text, size, font)
        if measured is not None:
            if measured <= box_width_in:
                return size
            size -= step
            continue
        # Fallback: heuristic ratio (no font file could be measured).
        ratio = _GLYPH_WIDTH_EM.get(font, 0.60)
        latin = sum(1 for c in text if ord(c) < 0x2E80)
        cjk = len(text) - latin
        width_in = (latin * ratio + cjk * _CJK_WIDTH_EM) * size / 72.0
        if width_in <= box_width_in:
            return size
        size -= step
    return size


def fit_paragraph_size(text: str, box_width_in: float, box_height_in: float,
                       base_pt: float, min_pt: float, *,
                       line_height: float = 1.4, step: float = 1.0) -> float:
    """Return the largest size <= base_pt that lets `text` word-wrap to
    fit box_height_in at box_width_in, for multi-line body copy (KPI
    insight callouts, etc).

    Unlike fit_title_size (single line), this has to know both how many
    characters fit per line AND how many lines fit — both change
    together as the font shrinks, so it re-derives capacity at each
    candidate size instead of computing a single "safe length" at the
    starting size and then guessing how much to shrink for the excess
    (that was v2.1's first attempt here and it under-shrank anything
    much longer than the demo insight string).
    """
    n = len(text)
    size = base_pt
    while size > min_pt:
        chars_per_line = max(1, box_width_in / (_CJK_WIDTH_EM * size / 72.0))
        lines_that_fit = max(1, box_height_in / (line_height * size / 72.0))
        if chars_per_line * lines_that_fit >= n:
            return size
        size -= step
    return min_pt


# =============================================================
#                    CHROME (brand row + footer)
# =============================================================

def _add_brand_row(slide, meta_right: Optional[str] = None,
                   on_dark: bool = False):
    ink = RGB_TEXT_DARK_1 if on_dark else RGB_TEXT_PRIMARY
    sub = RGB_TEXT_DARK_3 if on_dark else RGB_TEXT_TERTIARY
    _add_text(
        slide,
        SLIDE_MARGIN, Inches(0.4), Inches(6), Inches(0.3),
        [
            ("ACKS STUDIO ", {"font": FONT_DISPLAY_EN, "size": Pt(18),
                              "bold": True, "color": ink}),
            ("· ", {"font": FONT_DISPLAY_EN, "size": Pt(18),
                    "bold": True, "color": RGB_PRIMARY}),
            ("爱驰科驶", {"font": FONT_BODY_CN, "font_cn": FONT_BODY_CN,
                          "size": Pt(18), "bold": True, "color": ink}),
        ],
    )
    if meta_right:
        _add_text(
            slide,
            SLIDE_WIDTH - SLIDE_MARGIN - Inches(5),
            Inches(0.4), Inches(5), Inches(0.3),
            meta_right.upper(),
            font=FONT_MONO, font_cn=FONT_MONO,
            size=Pt(14), color=sub,
            align=PP_ALIGN.RIGHT, caps=True,
        )

_FOOTER_TOTAL_MARKER = "acks_footer_total"


def _add_footer(slide, page_no: Optional[int] = None,
                total: Optional[int] = None, on_dark: bool = False):
    """Add the standard footer. `total` is optional: if omitted, the
    "/NN" is written as a placeholder and patched to the real slide
    count by finalize_footers() once the whole deck has been built —
    this is what build_presentation()/save_demo() do automatically.
    Pass `total` explicitly only if you are adding slides one at a time
    outside of a build_*() function and know the final count upfront.
    """
    ink = RGB_TEXT_DARK_1 if on_dark else RGB_TEXT_PRIMARY
    sub = RGB_TEXT_DARK_3 if on_dark else RGB_TEXT_TERTIARY
    y = SLIDE_HEIGHT - Inches(0.7)
    _add_text(
        slide,
        SLIDE_MARGIN, y, Inches(5), Inches(0.3),
        "ACKS Studio · v2.1 · 2026",
        font=FONT_MONO, font_cn=FONT_MONO,
        size=SIZE_FOOTER, color=sub, caps=True,
    )
    if page_no:
        sq = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            SLIDE_WIDTH - SLIDE_MARGIN - Inches(1.4),
            y + Pt(2),
            Pt(10), Pt(10),
        )
        _set_solid_fill(sq, RGB_PRIMARY)
        _set_no_line(sq)
        total_display = f"{total:02d}" if total is not None else "??"
        tb = _add_text(
            slide,
            SLIDE_WIDTH - SLIDE_MARGIN - Inches(1.2),
            y - Pt(2), Inches(1.2), Inches(0.35),
            [
                (f"{page_no:02d}", {"font": FONT_DISPLAY_EN, "size": Pt(20),
                                    "bold": True, "color": ink}),
                (f"  / {total_display}", {"font": FONT_MONO, "size": SIZE_FOOTER,
                                          "color": sub}),
            ],
            align=PP_ALIGN.RIGHT,
        )
        if total is None:
            tb.name = _FOOTER_TOTAL_MARKER


def finalize_footers(prs: Presentation) -> int:
    """Patch every footer's '/NN' to the deck's real slide count.

    Call this once, after all slides have been added (build_presentation()
    and save_demo() already do this for you). Returns the slide count.
    This removes the old requirement that callers predict the final
    slide count before the first slide is drawn — the v2.0 default of
    `total=8` silently stuck on any deck the demo's length.
    """
    total = len(prs.slides)
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "name", "") == _FOOTER_TOTAL_MARKER and shape.has_text_frame:
                p = shape.text_frame.paragraphs[0]
                if len(p.runs) >= 2:
                    p.runs[1].text = f"  / {total:02d}"
    return total


# =============================================================
#                    PRESENTATION SETUP
# =============================================================

def init_presentation() -> Presentation:
    """Create a 1920×1080 16:9 Presentation."""
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs

def _blank_slide(prs: Presentation, fill: RGBColor = RGB_BG_PRIMARY):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = fill
    return slide


# =============================================================
#                    SLIDE FACTORIES
# =============================================================

def add_title_slide(prs, *, doctype: str, title_top: str, title_em: str,
                    zh_sub: str, meta_right: str = "CONFIDENTIAL · 2026"):
    """Dark cover with giant title + orange bar."""
    s = _blank_slide(prs, RGB_BG_DARK_0)
    _add_brand_row(s, meta_right, on_dark=True)

    title_box_w = Inches(11.6)
    # Fit against the *same* width the title textbox actually uses
    # below (11.6in) — v2.1's first pass measured against a wider,
    # hypothetical box (slide width minus margins) while the real
    # textbox was narrower, so "fits" was measured against a box that
    # didn't exist and the title still wrapped.
    size = fit_title_size(f"{title_top} {title_em}", title_box_w / Inches(1),
                          base_pt=92, min_pt=52)
    _add_text(s, SLIDE_MARGIN, Inches(3.5), Inches(11), Inches(0.4),
              f"— {doctype.upper()}",
              font=FONT_MONO, size=Pt(20), color=RGB_PRIMARY, caps=True)
    _add_text(s, SLIDE_MARGIN, Inches(3.95), title_box_w, Inches(1.5),
              [
                  (f"{title_top} ", {"font": FONT_DISPLAY_EN, "size": Pt(size),
                                     "bold": True, "color": RGB_TEXT_DARK_1}),
                  (title_em, {"font": FONT_DISPLAY_EN, "size": Pt(size),
                              "bold": True, "color": RGB_PRIMARY}),
              ])
    _add_text(s, SLIDE_MARGIN, Inches(5.75), Inches(11), Inches(0.9),
              zh_sub, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(24), bold=True, color=RGB_TEXT_DARK_2)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(0),
                             SLIDE_HEIGHT - Inches(0.25),
                             SLIDE_WIDTH, Inches(0.25))
    _set_solid_fill(bar, RGB_PRIMARY)
    _set_no_line(bar)


def add_section_slide(prs, *, n: str, title_top: str, title_em: str,
                      zh_sub: str, page_no: int, on_dark: bool = False,
                      chapter_no: Optional[int] = None):
    """Chapter intro page with huge title.

    `chapter_no` labels the "CHAPTER NN" badge. If omitted, it is parsed
    from the trailing digits of `n` (e.g. n="— SECTION 02" -> 2) rather
    than defaulting to page_no — page_no is this slide's position in the
    deck, which is almost never the same number as its chapter once the
    deck has more than one content slide per chapter (the v2.0 default
    only looked correct in the 8-slide demo, where they coincided).
    """
    s = _blank_slide(prs, RGB_BG_DARK_0 if on_dark else RGB_BG_PRIMARY)
    if chapter_no is None:
        digits = "".join(c for c in n if c.isdigit())
        chapter_no = int(digits) if digits else page_no
    _add_brand_row(s, f"CHAPTER {chapter_no:02d}", on_dark=on_dark)

    _add_text(s, SLIDE_MARGIN, Inches(2.5), Inches(10), Inches(0.4),
              n.upper(), font=FONT_MONO, size=Pt(20),
              color=RGB_PRIMARY, caps=True)
    title_color = RGB_TEXT_DARK_1 if on_dark else RGB_TEXT_PRIMARY
    size = fit_title_size(f"{title_top} {title_em}", 11.6, base_pt=120, min_pt=64)
    _add_text(s, SLIDE_MARGIN, Inches(2.9), Inches(12), Inches(1.9),
              [
                  (f"{title_top} ", {"font": FONT_DISPLAY_EN, "size": Pt(size),
                                     "bold": True, "color": title_color}),
                  (title_em, {"font": FONT_DISPLAY_EN, "size": Pt(size),
                              "bold": True, "color": RGB_PRIMARY}),
              ])
    sub_color = RGB_TEXT_DARK_2 if on_dark else RGB_TEXT_SECONDARY
    _add_text(s, SLIDE_MARGIN, Inches(5.1), Inches(11.4), Inches(0.5),
              zh_sub, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(28), bold=True, color=sub_color)
    rule_color = RGB_TEXT_DARK_1 if on_dark else RGB_TEXT_PRIMARY
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              SLIDE_MARGIN, Inches(5.65),
                              Inches(1.0), Pt(5))
    _set_solid_fill(rule, rule_color)
    _set_no_line(rule)

    _add_footer(s, page_no=page_no, on_dark=on_dark)


def add_content_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                      paragraphs: Sequence[str], page_no: int):
    """Single-column body with eyebrow + h2 + paragraphs.

    Title area reserves room for up to two wrapped lines at a moderate
    fixed size (rather than one line at SIZE_H1_BODY) so a title longer
    than the ~14-character demo string doesn't overlap title_zh/the rule.
    """
    s = _blank_slide(prs)
    _add_brand_row(s, eyebrow)

    _add_text(s, SLIDE_MARGIN, Inches(1.7), Inches(11), Inches(0.35),
              eyebrow.upper(), font=FONT_MONO, size=SIZE_EYEBROW,
              color=RGB_PRIMARY, caps=True)
    _add_text(s, SLIDE_MARGIN, Inches(2.05), Inches(11.4), Inches(1.5),
              title_en, font=FONT_DISPLAY_EN, size=Pt(42),
              bold=True, color=RGB_TEXT_PRIMARY)
    _add_text(s, SLIDE_MARGIN, Inches(3.6), Inches(11), Inches(0.42),
              title_zh, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(21), bold=True, color=RGB_TEXT_SECONDARY)

    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              SLIDE_MARGIN, Inches(4.1),
                              SLIDE_WIDTH - SLIDE_MARGIN * 2, Pt(2))
    _set_solid_fill(rule, RGB_TEXT_PRIMARY)
    _set_no_line(rule)

    tb = s.shapes.add_textbox(SLIDE_MARGIN, Inches(4.4),
                              SLIDE_WIDTH - SLIDE_MARGIN * 2, Inches(2.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    first = tf.paragraphs[0]
    first.alignment = PP_ALIGN.LEFT
    run = first.add_run()
    run.text = paragraphs[0]
    run.font.name = FONT_BODY_CN
    run.font.size = SIZE_BODY
    run.font.color.rgb = RGB_TEXT_PRIMARY
    _set_east_asia_font(run, FONT_BODY_CN)
    for ptxt in paragraphs[1:]:
        _add_paragraph(tf, ptxt, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                       size=SIZE_BODY, color=RGB_TEXT_PRIMARY,
                       space_before=Pt(12))

    _add_footer(s, page_no=page_no)


def add_data_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                   kpis: Sequence[Mapping[str, str]],
                   chart_values: Sequence[float],
                   chart_labels: Sequence[str],
                   insight: str, page_no: int):
    """KPI cells + bar chart + key insight box."""
    s = _blank_slide(prs)
    _add_brand_row(s, eyebrow)

    _add_text(s, SLIDE_MARGIN, Inches(1.7), Inches(11), Inches(0.35),
              eyebrow.upper(), font=FONT_MONO, size=SIZE_EYEBROW,
              color=RGB_PRIMARY, caps=True)
    # This slide is the tightest budget in the deck (title + KPI row +
    # chart/insight all have to fit above the footer), so the title
    # only reserves 1.15in — a modest wrap allowance — and leans on the
    # more conservative fit-size ratio to stay on one line in practice.
    # KPI/chart/insight sizing below is what actually adapts to content
    # length; the title budget is deliberately the tighter one here.
    size = fit_title_size(title_en, 11.2, base_pt=72, min_pt=40)
    _add_text(s, SLIDE_MARGIN, Inches(2.05), Inches(11.4), Inches(1.15),
              title_en, font=FONT_DISPLAY_EN, size=Pt(size),
              bold=True, color=RGB_TEXT_PRIMARY, anchor=MSO_ANCHOR.TOP)
    _add_text(s, SLIDE_MARGIN, Inches(3.25), Inches(11), Inches(0.4),
              title_zh, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(21), bold=True, color=RGB_TEXT_SECONDARY)

    # KPI row
    kpi_y = Inches(3.8)
    kpi_h = Inches(1.4)
    kpi_w = (SLIDE_WIDTH - SLIDE_MARGIN * 2) / max(1, len(kpis))
    top_rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  SLIDE_MARGIN, kpi_y,
                                  SLIDE_WIDTH - SLIDE_MARGIN * 2, Pt(2))
    _set_solid_fill(top_rule, RGB_TEXT_PRIMARY); _set_no_line(top_rule)
    bot_rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  SLIDE_MARGIN, kpi_y + kpi_h,
                                  SLIDE_WIDTH - SLIDE_MARGIN * 2, Pt(2))
    _set_solid_fill(bot_rule, RGB_TEXT_PRIMARY); _set_no_line(bot_rule)

    kpi_w_in = (13.33333 - 1.111 * 2) / max(1, len(kpis))
    for i, item in enumerate(kpis):
        x = SLIDE_MARGIN + kpi_w * i
        _add_text(s, x + Inches(0.2), kpi_y + Inches(0.1),
                  kpi_w - Inches(0.4), Inches(0.3),
                  item.get("label", "").upper(),
                  font=FONT_MONO, size=Pt(14), color=RGB_TEXT_TERTIARY, caps=True)
        # Value font shrinks to fit the column — v2.0 fixed this at 64pt
        # regardless of string length, so anything longer than ~4 chars
        # (e.g. "¥420M") wrapped onto a second line and collided with
        # the delta text below it.
        v_size = fit_title_size(item.get("value", ""), kpi_w_in - 0.55,
                                base_pt=64, min_pt=28)
        _add_text(s, x + Inches(0.2), kpi_y + Inches(0.35),
                  kpi_w - Inches(0.4), Inches(0.7),
                  item.get("value", ""),
                  font=FONT_DISPLAY_EN, size=Pt(v_size),
                  bold=True, color=RGB_TEXT_PRIMARY)
        if item.get("delta"):
            _add_text(s, x + Inches(0.2), kpi_y + Inches(1.1),
                      kpi_w - Inches(0.4), Inches(0.25),
                      item["delta"],
                      font=FONT_MONO, size=Pt(16), color=RGB_PRIMARY)
        if i < len(kpis) - 1:
            div = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                     x + kpi_w - Pt(0.5), kpi_y + Pt(2),
                                     Pt(1), kpi_h - Pt(4))
            _set_solid_fill(div, RGB_BG_QUATERNARY); _set_no_line(div)

    # Chart + insight
    bottom_y = kpi_y + kpi_h + Inches(0.2)
    bottom_h = FOOTER_SAFE_TOP - bottom_y
    chart_w = Inches(8.0)
    chart_box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   SLIDE_MARGIN, bottom_y,
                                   chart_w, bottom_h)
    _set_solid_fill(chart_box, RGB_BG_SECONDARY)
    _set_line(chart_box, RGB_BG_QUATERNARY, Pt(0.75))

    max_v = max(chart_values) if chart_values else 1
    bar_area_left = SLIDE_MARGIN + Inches(0.6)
    bar_area_right = SLIDE_MARGIN + chart_w - Inches(0.3)
    bar_area_w = bar_area_right - bar_area_left
    bar_area_top = bottom_y + Inches(0.2)
    bar_area_bottom = bottom_y + bottom_h - Inches(0.4)
    bar_area_h = bar_area_bottom - bar_area_top
    n = len(chart_values)
    gap = Inches(0.1)
    bar_w = (bar_area_w - gap * (n - 1)) / max(1, n)
    for i, v in enumerate(chart_values):
        h = bar_area_h * (v / max_v)
        x = bar_area_left + (bar_w + gap) * i
        y = bar_area_bottom - h
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, bar_w, h)
        is_hi = (i == n - 1)
        _set_solid_fill(bar, RGB_PRIMARY if is_hi else RGB_TEXT_PRIMARY)
        _set_no_line(bar)
        _add_text(s, x, bar_area_bottom + Pt(4), bar_w, Inches(0.25),
                  chart_labels[i] if i < len(chart_labels) else "",
                  font=FONT_MONO, size=Pt(12), color=RGB_TEXT_SECONDARY,
                  align=PP_ALIGN.CENTER)

    ins_x = SLIDE_MARGIN + chart_w + Inches(0.3)
    ins_w = SLIDE_WIDTH - SLIDE_MARGIN - ins_x
    ins_box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 ins_x, bottom_y, ins_w, bottom_h)
    _set_solid_fill(ins_box, RGB_PRIMARY_SOFT)
    _set_no_line(ins_box)
    lbar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              ins_x, bottom_y, Pt(5), bottom_h)
    _set_solid_fill(lbar, RGB_PRIMARY); _set_no_line(lbar)
    _add_text(s, ins_x + Inches(0.3), bottom_y + Inches(0.2),
              ins_w - Inches(0.4), Inches(0.3),
              "KEY · 关键洞察", font=FONT_MONO, font_cn=FONT_BODY_CN,
              size=Pt(13), color=RGB_PRIMARY_DEEP, caps=True)
    # v2.0 assumed every insight was demo-length (~55 chars) at a fixed
    # Pt(18) and let anything longer run past the box into the footer.
    ins_body_w_in = (ins_w / Inches(1)) - 0.5
    ins_body_h_in = (bottom_h / Inches(1)) - 0.55
    ins_size = fit_paragraph_size(insight, ins_body_w_in, ins_body_h_in,
                                  base_pt=18, min_pt=10)
    _add_text(s, ins_x + Inches(0.3), bottom_y + Inches(0.55),
              ins_w - Inches(0.4), bottom_h - Inches(0.6),
              insight, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(ins_size), color=RGB_TEXT_PRIMARY)

    _add_footer(s, page_no=page_no)


def add_compare_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                      option_a: Mapping, option_b: Mapping, page_no: int):
    """Two columns, right is lead (filled ink).

    Column width and the option-name font size are computed in plain
    float inches throughout (v2.0 divided an Emu length by a bare int
    partway through, which is not a unit conversion and produced a
    meaningless box_width value — the option name would then wrap onto
    the item list below it for anything longer than ~17 characters).
    Item row height also now scales with len(items) instead of a fixed
    assumption of ~4 short bullets per option.
    """
    s = _blank_slide(prs)
    _add_brand_row(s, eyebrow)

    _add_text(s, SLIDE_MARGIN, Inches(1.7), Inches(11), Inches(0.35),
              eyebrow.upper(), font=FONT_MONO, size=SIZE_EYEBROW,
              color=RGB_PRIMARY, caps=True)
    size = fit_title_size(title_en, 11.2, base_pt=76, min_pt=40)
    _add_text(s, SLIDE_MARGIN, Inches(2.05), Inches(11.4), Inches(1.4),
              title_en, font=FONT_DISPLAY_EN, size=Pt(size),
              bold=True, color=RGB_TEXT_PRIMARY, anchor=MSO_ANCHOR.TOP)
    _add_text(s, SLIDE_MARGIN, Inches(3.55), Inches(11), Inches(0.4),
              title_zh, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(22), bold=True, color=RGB_TEXT_SECONDARY)

    col_y = Inches(4.1)
    col_h = FOOTER_SAFE_TOP - col_y
    slide_w_in, margin_in = 13.33333, 1.111
    col_w_in = (slide_w_in - margin_in * 2 - 0.35) / 2
    col_w = Inches(col_w_in)

    for i, opt in enumerate((option_a, option_b)):
        x = SLIDE_MARGIN + (col_w + Inches(0.35)) * i
        is_lead = opt.get("lead", False)
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, col_y, col_w, col_h)
        _set_solid_fill(box, RGB_TEXT_PRIMARY if is_lead else RGB_BG_PRIMARY)
        _set_line(box, RGB_TEXT_PRIMARY, Pt(1.25))
        tag_color = RGB_PRIMARY if is_lead else RGB_TEXT_TERTIARY
        _add_text(s, x + Inches(0.26), col_y + Inches(0.15),
                  col_w - Inches(0.52), Inches(0.28),
                  opt.get("tag", "").upper(), font=FONT_MONO,
                  size=Pt(13), color=tag_color, caps=True)
        name_color = RGB_TEXT_DARK_1 if is_lead else RGB_TEXT_PRIMARY
        name_size = fit_title_size(opt.get("name", ""), col_w_in - 0.52,
                                   base_pt=32, min_pt=17, font=FONT_DISPLAY_EN)
        _add_text(s, x + Inches(0.26), col_y + Inches(0.46),
                  col_w - Inches(0.52), Inches(0.55),
                  opt.get("name", ""), font=FONT_DISPLAY_EN, font_cn=FONT_BODY_CN,
                  size=Pt(name_size), bold=True, color=name_color)
        rule_color = RGB_PRIMARY if is_lead else RGB_TEXT_PRIMARY
        rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  x + Inches(0.26), col_y + Inches(1.05),
                                  col_w - Inches(0.52), Pt(2))
        _set_solid_fill(rule, rule_color); _set_no_line(rule)
        bullet_color = RGB_TEXT_DARK_1 if is_lead else RGB_TEXT_PRIMARY

        items = list(opt.get("items", []))
        item_area_h = col_h - Inches(1.2)
        item_h = item_area_h / max(1, len(items))
        item_font = Pt(20) if len(items) <= 4 else Pt(max(13, 20 - 1.5 * (len(items) - 4)))
        iy = col_y + Inches(1.15)
        for item in items:
            _add_text(s, x + Inches(0.26), iy, Inches(0.3), item_h, "+",
                      font=FONT_MONO, size=item_font, bold=True, color=RGB_PRIMARY)
            _add_text(s, x + Inches(0.55), iy, col_w - Inches(0.8), item_h,
                      item, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                      size=item_font, color=bullet_color)
            iy += item_h

    _add_footer(s, page_no=page_no)


def add_quote_slide(prs, *, quote: str, source_label: str, source: str,
                    page_no: int):
    """Dark, serif Chinese, big orange quote mark."""
    s = _blank_slide(prs, RGB_BG_DARK_0)
    _add_brand_row(s, "VOICE · MANIFESTO", on_dark=True)

    # Quote font shrinks with length so 1-2 short lines still read big
    # (the v2.0 default) while a longer quote doesn't run past the
    # source line below it.
    q_size = 54 if len(quote) <= 26 else max(32, 54 - (len(quote) - 26) // 2)
    _add_text(s, SLIDE_MARGIN, Inches(1.5), Inches(3), Inches(2.2),
              "\u201C", font=FONT_DISPLAY_EN, size=Pt(240),
              bold=True, color=RGB_PRIMARY)
    _add_text(s, SLIDE_MARGIN, Inches(3.1), Inches(11), Inches(2.4),
              quote, font=FONT_SERIF_CN, font_cn=FONT_SERIF_CN,
              size=Pt(q_size), bold=False, color=RGB_TEXT_DARK_1)
    _add_text(s, SLIDE_MARGIN, Inches(5.85), Inches(11), Inches(0.3),
              source_label.upper(), font=FONT_MONO,
              size=Pt(13), color=RGB_TEXT_DARK_3, caps=True)
    _add_text(s, SLIDE_MARGIN, Inches(6.15), Inches(11), Inches(0.5),
              source, font=FONT_DISPLAY_EN, size=Pt(20),
              bold=True, color=RGB_TEXT_DARK_1)

    _add_footer(s, page_no=page_no, on_dark=True)


def add_closing_slide(prs, *, title_top: str, title_em: str,
                      contacts: Sequence[Mapping[str, str]],
                      page_no: int = None):
    """Light, big title bottom-left, contact columns."""
    s = _blank_slide(prs)
    _add_brand_row(s, "THANK YOU · 2026")

    size = fit_title_size(f"{title_top} {title_em}", 11.4, base_pt=96, min_pt=56)
    _add_text(s, SLIDE_MARGIN, Inches(3.4), Inches(11.6), Inches(1.75),
              [
                  (f"{title_top} ", {"font": FONT_DISPLAY_EN, "size": Pt(size),
                                     "bold": True, "color": RGB_TEXT_PRIMARY}),
                  (title_em, {"font": FONT_DISPLAY_EN, "size": Pt(size),
                              "bold": True, "color": RGB_PRIMARY}),
              ])

    rule_y = Inches(5.95)
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              SLIDE_MARGIN, rule_y,
                              SLIDE_WIDTH - SLIDE_MARGIN * 2, Pt(2))
    _set_solid_fill(rule, RGB_TEXT_PRIMARY); _set_no_line(rule)

    col_w = (SLIDE_WIDTH - SLIDE_MARGIN * 2) / max(1, len(contacts))
    for i, c in enumerate(contacts):
        x = SLIDE_MARGIN + col_w * i
        _add_text(s, x, rule_y + Inches(0.2), col_w, Inches(0.3),
                  c.get("label", "").upper(),
                  font=FONT_MONO, size=Pt(14),
                  color=RGB_TEXT_TERTIARY, caps=True)
        _add_text(s, x, rule_y + Inches(0.5), col_w, Inches(0.5),
                  c.get("value", ""),
                  font=FONT_DISPLAY_EN, size=Pt(28),
                  bold=True, color=RGB_TEXT_PRIMARY)


# =============================================================
#                    LIST / CHECKLIST / TEMPLATE SLIDES
#
#  v2.0 had no bullet-list, checklist, or code/prompt-template slide on
#  the pptx side at all, despite these being some of the most common
#  shapes real content takes (the docx module has the same gap — see
#  add_bullet_list / add_checklist / add_code_block in tokens.py).
#  These three share the eyebrow/title/rule header used everywhere else
#  and size their rows against CONTENT_TOP/FOOTER_SAFE_TOP so they never
#  need a fixed "how many bullets fit" assumption.
# =============================================================

def _slide_header(s, eyebrow: str, title_en: str, title_zh: str):
    """Shared eyebrow + title + zh-subtitle + rule, used by the list-
    style factories below. Title is capped to two safe lines instead of
    a single fixed-size line — see add_content_slide for the rationale.
    """
    _add_brand_row(s, eyebrow)
    _add_text(s, SLIDE_MARGIN, Inches(1.7), Inches(11.4), Inches(0.35),
              eyebrow.upper(), font=FONT_MONO, size=SIZE_EYEBROW,
              color=RGB_PRIMARY, caps=True)
    _add_text(s, SLIDE_MARGIN, Inches(2.05), Inches(11.4), Inches(1.5),
              title_en, font=FONT_DISPLAY_EN, size=Pt(42),
              bold=True, color=RGB_TEXT_PRIMARY)
    _add_text(s, SLIDE_MARGIN, Inches(3.6), Inches(11.4), Inches(0.42),
              title_zh, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
              size=Pt(21), bold=True, color=RGB_TEXT_SECONDARY)
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              SLIDE_MARGIN, Inches(4.1),
                              SLIDE_WIDTH - SLIDE_MARGIN * 2, Pt(2))
    _set_solid_fill(rule, RGB_TEXT_PRIMARY)
    _set_no_line(rule)


def add_list_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                   items: Sequence[str], page_no: int,
                   numbered: bool = False, columns: int = 1,
                   note: Optional[str] = None):
    """Bullet or numbered list. Row height is CONTENT_BUDGET / rows, so
    the list degrades gracefully instead of overflowing into the footer
    when it has more items than a demo would.
    """
    s = _blank_slide(prs)
    _slide_header(s, eyebrow, title_en, title_zh)

    content_top = Inches(4.35)
    area_w = SLIDE_WIDTH - SLIDE_MARGIN * 2
    note_h = Inches(0.72) if note else Inches(0)
    note_gap = Inches(0.14) if note else Inches(0)
    budget = FOOTER_SAFE_TOP - content_top - note_h - note_gap

    col_w = area_w if columns == 1 else (area_w - Inches(0.5)) / 2
    n = len(items)
    per_col = n if columns == 1 else -(-n // columns)
    row_h = min(Inches(0.62), budget / max(1, per_col))

    for i, item in enumerate(items):
        col = 0 if columns == 1 else i // per_col
        row = i if columns == 1 else i % per_col
        x = SLIDE_MARGIN + col * (col_w + Inches(0.5))
        y = content_top + row_h * row
        marker = f"{i + 1:02d}" if numbered else GLYPH_BULLET
        mcolor = RGB_PRIMARY_DEEP if numbered else RGB_PRIMARY
        _add_text(s, x, y, Inches(0.55), Inches(0.42), marker,
                  font=FONT_MONO, size=Pt(16), bold=True, color=mcolor)
        _add_text(s, x + Inches(0.5), y, col_w - Inches(0.5), row_h,
                  item, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                  size=Pt(15.5), color=RGB_TEXT_PRIMARY)

    if note:
        note_y = content_top + row_h * per_col + note_gap
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SLIDE_MARGIN, note_y,
                                 area_w, note_h)
        _set_solid_fill(box, RGB_PRIMARY_SOFT); _set_no_line(box)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SLIDE_MARGIN, note_y,
                                 Pt(5), note_h)
        _set_solid_fill(bar, RGB_PRIMARY); _set_no_line(bar)
        _add_text(s, SLIDE_MARGIN + Inches(0.25), note_y + Inches(0.08),
                  area_w - Inches(0.5), note_h - Inches(0.16), note,
                  font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                  size=Pt(13.5), color=RGB_TEXT_PRIMARY)

    _add_footer(s, page_no=page_no)


def add_checklist_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                        items: Sequence[str], page_no: int):
    """Open checkbox + item, row height distributed across CONTENT_BUDGET."""
    s = _blank_slide(prs)
    _slide_header(s, eyebrow, title_en, title_zh)

    content_top = Inches(4.35)
    area_w = SLIDE_WIDTH - SLIDE_MARGIN * 2
    row_h = (FOOTER_SAFE_TOP - content_top) / max(1, len(items))
    y = content_top
    for item in items:
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SLIDE_MARGIN,
                                 y + Pt(3), Pt(14), Pt(14))
        box.fill.background()
        _set_line(box, RGB_PRIMARY, Pt(1.5))
        _add_text(s, SLIDE_MARGIN + Inches(0.3), y, area_w - Inches(0.3),
                  row_h, item, font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                  size=Pt(15), color=RGB_TEXT_PRIMARY)
        y += row_h

    _add_footer(s, page_no=page_no)


def add_template_slide(prs, *, eyebrow: str, title_en: str, title_zh: str,
                       code: str, page_no: int,
                       caption: Optional[str] = None,
                       side_note: Optional[str] = None):
    """Monospace box for prompt/code templates, with an optional
    "pitfalls" side panel — the shape most of ACKS's own scenario
    templates in real decks take, and something v2.0 had no primitive
    for at all.
    """
    s = _blank_slide(prs)
    _slide_header(s, eyebrow, title_en, title_zh)

    content_top = Inches(4.35)
    budget = FOOTER_SAFE_TOP - content_top
    area_w = SLIDE_WIDTH - SLIDE_MARGIN * 2
    box_h = budget - (Inches(0.28) if caption else Inches(0))
    code_w = area_w if not side_note else area_w * 0.60

    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SLIDE_MARGIN, content_top,
                             code_w, box_h)
    _set_solid_fill(box, RGB_BG_TERTIARY); _set_no_line(box)
    top_rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SLIDE_MARGIN,
                                  content_top, code_w, Pt(3))
    _set_solid_fill(top_rule, RGB_PRIMARY); _set_no_line(top_rule)

    lines = code.strip("\n").split("\n")
    font_size = Pt(11.5) if len(lines) <= 11 else Pt(10)
    tb = s.shapes.add_textbox(SLIDE_MARGIN + Inches(0.22),
                              content_top + Inches(0.14),
                              code_w - Inches(0.44), box_h - Inches(0.24))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2)
        r = p.add_run()
        r.text = line if line else " "
        r.font.name = FONT_MONO
        r.font.size = font_size
        r.font.color.rgb = RGB_TEXT_PRIMARY
        _set_east_asia_font(r, FONT_MONO)

    if caption:
        _add_text(s, SLIDE_MARGIN, content_top + box_h + Inches(0.04),
                  code_w, Inches(0.26), caption, font=FONT_MONO, size=Pt(11.5),
                  color=RGB_TEXT_TERTIARY, caps=True)

    if side_note:
        nx = SLIDE_MARGIN + code_w + Inches(0.28)
        nw = area_w - code_w - Inches(0.28)
        nbox = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, nx, content_top, nw, budget)
        _set_solid_fill(nbox, RGB_PRIMARY_SOFT); _set_no_line(nbox)
        nbar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, nx, content_top, Pt(4), budget)
        _set_solid_fill(nbar, RGB_PRIMARY); _set_no_line(nbar)
        _add_text(s, nx + Inches(0.22), content_top + Inches(0.14),
                  nw - Inches(0.44), Inches(0.28), "注意 · PITFALLS",
                  font=FONT_MONO, size=Pt(11.5), color=RGB_PRIMARY_DEEP, caps=True)
        _add_text(s, nx + Inches(0.22), content_top + Inches(0.46),
                  nw - Inches(0.44), budget - Inches(0.55), side_note,
                  font=FONT_BODY_CN, font_cn=FONT_BODY_CN,
                  size=Pt(12.5), color=RGB_TEXT_PRIMARY)

    _add_footer(s, page_no=page_no)


# =============================================================
#                    DEMO BUILD
# =============================================================

def build_presentation(
    kpis: Optional[Sequence[Mapping[str, str]]] = None,
    chart_values: Optional[Sequence[float]] = None,
    chart_labels: Optional[Sequence[str]] = None,
    insight: Optional[str] = None,
) -> Presentation:
    """Construct the full 11-slide ACKS demo deck.

    v2.1 adds add_list_slide / add_template_slide / add_checklist_slide
    to the demo (previously 8 slides) so the three new components ship
    with working, screenshot-able examples instead of only unit tests.

    All data parameters have sensible defaults but can be overridden.
    """
    prs = init_presentation()

    add_title_slide(
        prs,
        doctype="Q1 REPORT · 2026",
        title_top="State of",
        title_em="the Drive",
        zh_sub="2026 年第一季度业务回顾 · 汽车改装、AI、软件服务三条业务线",
    )

    add_section_slide(
        prs,
        n="— SECTION 02",
        title_top="Market",
        title_em="Overview",
        zh_sub="市场概览 · 中国汽车改装行业 Q1 走势",
        page_no=2,
    )

    add_content_slide(
        prs,
        eyebrow="2.1 · Industry Context",
        title_en="Aftermarket Landscape",
        title_zh="汽车改装行业格局",
        paragraphs=[
            "2024 年中国汽车改装市场规模突破 ¥1,420 亿，"
            "年复合增长率 17.4%。性能改装与外观空气动力两条主线下，"
            "市场正在从「线下门店」加速向「品牌 + 服务 + 内容」的复合体迁移。",
            "高净值年轻用户的崛起、二手豪华车保有量的快速增长，"
            "以及社交平台对赛车文化的放大效应，共同构成了未来三年的核心增长曲线。",
        ],
        page_no=3,
    )

    add_list_slide(
        prs,
        eyebrow="2.2 · Growth Drivers",
        title_en="Three Growth Drivers",
        title_zh="三条核心增长曲线",
        items=[
            "高净值年轻用户占比持续提升，改装预算高于行业均值 2.3 倍",
            "二手豪华车保有量三年复合增长 21%，售后改装是天然延伸需求",
            "社交平台对赛车文化的放大效应，让内容成为获客的第一入口",
        ],
        page_no=4,
        note="三条曲线共同的前提：门店必须先完成「品牌化」，否则流量接不住。",
    )

    if kpis is None:
        kpis = [
            {"label": "Revenue",  "value": "¥420M", "delta": "▲ +22.0% YoY"},
            {"label": "Margin",   "value": "38.6%", "delta": "▲ +2.4 pts"},
            {"label": "Clients",  "value": "1,284", "delta": "▲ +18% QoQ"},
            {"label": "NPS",      "value": "62",    "delta": "▲ +1 pt"},
        ]
    if chart_values is None:
        chart_values = [28, 38, 48, 60, 72, 86]
    if chart_labels is None:
        chart_labels = ["Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
    if insight is None:
        insight = ("高端品牌 + 内容运营 + 智能体客服是未来 18 个月最重要的差异化组合 — "
                   "缺一不可。三者协同的客户 LTV 比单点接触高出 2.7 倍。")

    add_data_slide(
        prs,
        eyebrow="Q1 Snapshot · 2026",
        title_en="Q1 in Numbers",
        title_zh="关键指标 — 全部高于 Q4",
        kpis=kpis,
        chart_values=chart_values,
        chart_labels=chart_labels,
        insight=insight,
        page_no=5,
    )

    add_compare_slide(
        prs,
        eyebrow="Option Comparison",
        title_en="Build vs Partner",
        title_zh="自建 vs 与 ACKS 合作 — 三年视角",
        option_a={
            "tag": "Option A",
            "name": "Self-build",
            "items": [
                "3 年累计成本 ¥120M",
                "第 18 月才进入市场",
                "团队规模 40+ 人",
                "独立掌控全部决策",
            ],
        },
        option_b={
            "lead": True,
            "tag": "Option B · 推荐",
            "name": "Partner with ACKS",
            "items": [
                "3 年累计成本 ¥56M",
                "第 4 月即开始售卖",
                "团队规模 8 人对接即可",
                "共享中国渠道资源",
            ],
        },
        page_no=6,
    )

    add_template_slide(
        prs,
        eyebrow="Playbook · Scenario 01",
        title_en="Store Reopening Brief",
        title_zh="门店品牌化 · 交接模板",
        code=(
            "角色：门店品牌顾问\n"
            "背景：[城市]门店完成品牌化改造，需要给店长一份开业前检查单\n"
            "任务：\n"
            "1. 列出开业前必须核对的 5 个品牌一致性项\n"
            "2. 每项给出验收标准，不满足的标记待整改\n"
            "格式：表格，列 = 检查项 / 标准 / 状态"
        ),
        caption="Template · 门店品牌化交接",
        side_note="专属坑：验收标准必须可肉眼判断（颜色/物料/陈列），"
                 "不要写“风格统一”这类无法验收的描述。",
        page_no=7,
    )

    add_quote_slide(
        prs,
        quote="改装不再只是机械的事情，而是一种生活方式的表达。",
        source_label="Quoted from",
        source="ACKS Studio Brand Manifesto, 2026",
        page_no=8,
    )

    add_checklist_slide(
        prs,
        eyebrow="Pre-Q2 Checklist",
        title_en="Before We Move to Q2",
        title_zh="进入 Q2 前的最后检查",
        items=[
            "高端门店品牌化改造是否已在 3 个试点城市完成？",
            "内容团队的月更频率是否达到 8 条以上？",
            "智能体客服的转人工率是否已降到 15% 以下？",
            "Q1 的客户 LTV 数据是否已经和财务口径对齐？",
        ],
        page_no=9,
    )

    add_section_slide(
        prs,
        n="— SECTION 05",
        title_top="What's",
        title_em="Next",
        zh_sub="下季展望 · Q2 优先级与风险",
        page_no=10,
        on_dark=True,
    )

    add_closing_slide(
        prs,
        title_top="Let's",
        title_em="Drive Together",
        contacts=[
            {"label": "Contact", "value": "S. Lu · CEO"},
            {"label": "Email",   "value": "studio@acks.cn"},
            {"label": "Next",    "value": "05.30 · Workshop"},
        ],
    )

    finalize_footers(prs)
    return prs


def save_demo(path: str = "ACKS-deck.pptx") -> str:
    prs = build_presentation()
    prs.save(path)
    return path


if __name__ == "__main__":
    out = save_demo()
    print(f"Saved → {out}")
