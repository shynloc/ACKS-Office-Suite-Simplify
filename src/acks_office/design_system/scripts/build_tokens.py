#!/usr/bin/env python3
"""
Build script: reads tokens.json and generates Python constant modules.

Usage:
    python scripts/build_tokens.py [--check]

Without --check: overwrites tokens.py, slides.py, xlsx.py with constants
    derived from tokens.json.

With --check: exits non-zero if any generated file differs from disk,
    useful for CI.
"""

from __future__ import annotations

import json
import sys
import textwrap
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOKENS_PATH = ROOT / "tokens.json"


def load_tokens() -> dict:
    with open(TOKENS_PATH, encoding="utf-8") as f:
        return json.load(f)


def hex_to_rgb_expr(hex_str: str) -> str:
    """Convert '#FF6B1A' to 'RGBColor(0xFF, 0x6B, 0x1A)'."""
    h = hex_str.lstrip("#")
    r, g, b = h[0:2], h[2:4], h[4:6]
    return f"RGBColor(0x{r.upper()}, 0x{g.upper()}, 0x{b.upper()})"


def hex_clean(hex_str: str) -> str:
    """Convert '#FF6B1A' to 'FF6B1A'."""
    return hex_str.lstrip("#").upper()


# ── tokens.py constants ─────────────────────────────────────────────

def build_tokens_constants(tok: dict) -> str:
    lines = [
        '# -*- coding: utf-8 -*-',
        '"""',
        'Auto-generated from tokens.json — DO NOT EDIT BY HAND.',
        'Run: python scripts/build_tokens.py',
        '"""',
        '',
        'from docx.shared import Mm, Pt, RGBColor',
        '',
        '',
        '# ── COLOR TOKENS ──────────────────────────────────────────────',
        '',
    ]

    # Brand
    lines.append('# Brand / Accent')
    for name, val in tok["color"]["brand"].items():
        lines.append(f'COLOR_{name} = "{hex_clean(val)}"')
    lines.append('')

    # Light
    lines.append('# Light scheme')
    for name, val in tok["color"]["light"].items():
        lines.append(f'COLOR_{name} = "{hex_clean(val)}"')
    lines.append('')

    # Dark
    lines.append('# Dark scheme')
    for name, val in tok["color"]["dark"].items():
        lines.append(f'COLOR_{name} = "{hex_clean(val)}"')
    lines.append('')

    # RGB shortcuts
    lines.append('# Quick-access RGBColor objects')
    brand = tok["color"]["brand"]
    light = tok["color"]["light"]
    lines.append(f'RGB_PRIMARY      = {hex_to_rgb_expr(brand["PRIMARY"])}')
    lines.append(f'RGB_PRIMARY_DEEP = {hex_to_rgb_expr(brand["PRIMARY_DEEP"])}')
    lines.append(f'RGB_PRIMARY_SOFT = {hex_to_rgb_expr(brand["PRIMARY_SOFT"])}')
    lines.append(f'RGB_INK          = {hex_to_rgb_expr(light["TEXT_PRIMARY"])}')
    lines.append(f'RGB_INK_2        = {hex_to_rgb_expr(light["TEXT_SECONDARY"])}')
    lines.append(f'RGB_INK_3        = {hex_to_rgb_expr(light["TEXT_TERTIARY"])}')
    lines.append(f'RGB_RULE         = {hex_to_rgb_expr(light["RULE"])}')
    lines.append(f'RGB_BG_2         = {hex_to_rgb_expr(light["BG_SECONDARY"])}')
    lines.append(f'RGB_BG_3         = {hex_to_rgb_expr(light["BG_TERTIARY"])}')
    lines.append(f'RGB_WHITE        = RGBColor(0xFF, 0xFF, 0xFF)')
    lines.append('')

    # Fonts
    lines.append('# ── FONT TOKENS ───────────────────────────────────────────────')
    lines.append('')
    for name, val in tok["font"]["family"].items():
        lines.append(f'FONT_{name} = "{val}"')
    lines.append('')

    # Sizes
    lines.append('# ── SIZE TOKENS ───────────────────────────────────────────────')
    lines.append('')
    for name, val in tok["font"]["size_pt"].items():
        lines.append(f'FONT_{name}_SIZE = Pt({val})')
    lines.append('')

    # Line heights
    lines.append('# ── LINE HEIGHT TOKENS ────────────────────────────────────────')
    lines.append('')
    for name, val in tok["font"]["line_height"].items():
        lines.append(f'LINE_HEIGHT_{name} = {val}')
    lines.append('')

    # Spacing
    lines.append('# ── SPACING TOKENS ────────────────────────────────────────────')
    lines.append('')
    for name, val in tok["spacing_pt"].items():
        if name == "BASE":
            lines.append(f'SPACE_BASE = Pt({val})')
        elif name.startswith("PARAGRAPH"):
            lines.append(f'PARAGRAPH_SPACE_{name.split("_")[1]} = Pt({val})')
        else:
            # S4 -> SPACE_4, S8 -> SPACE_8, etc.
            num = name.lstrip("S")
            lines.append(f'SPACE_{num} = Pt({val})')
    lines.append('')

    # Page geometry
    lines.append('# ── PAGE GEOMETRY ─────────────────────────────────────────────')
    lines.append('')
    a4 = tok["page"]["a4"]
    # Map token names to match the original tokens.py conventions
    name_map = {
        "WIDTH_MM": "WIDTH", "HEIGHT_MM": "HEIGHT",
        "MARGIN_TOP_MM": "MARGIN_TOP", "MARGIN_BOTTOM_MM": "MARGIN_BOTTOM",
        "MARGIN_LEFT_MM": "MARGIN_LEFT", "MARGIN_RIGHT_MM": "MARGIN_RIGHT",
        "HEADER_DIST_MM": "HEADER_DIST", "FOOTER_DIST_MM": "FOOTER_DIST",
    }
    for name, val in a4.items():
        py_name = name_map.get(name, name)
        lines.append(f'PAGE_{py_name} = Mm({val})')
    lines.append('')

    # Glyphs
    lines.append('# ── GLYPH TOKENS ──────────────────────────────────────────────')
    lines.append('')
    for name, val in tok.get("glyph", {}).items():
        if name.startswith("_"):
            continue
        lines.append(f'GLYPH_{name} = {val!r}')
    lines.append('')

    # Content-length budgets (docx only has DOC_H2_CHARS today)
    lines.append('# ── CONTENT LIMIT TOKENS ──────────────────────────────────────')
    lines.append('')
    for name, val in tok.get("content_limit", {}).items():
        if name.startswith("_") or not name.startswith("DOC_"):
            continue
        lines.append(f'LIMIT_{name} = {val}')
    lines.append('')

    return '\n'.join(lines) + '\n'


# ── slides.py constants ─────────────────────────────────────────────

def build_slides_constants(tok: dict) -> str:
    lines = [
        '# -*- coding: utf-8 -*-',
        '"""',
        'Auto-generated from tokens.json — DO NOT EDIT BY HAND.',
        'Run: python scripts/build_tokens.py',
        '"""',
        '',
        'from pptx.dml.color import RGBColor',
        'from pptx.util import Inches, Pt',
        '',
        '',
        '# ── SLIDE TOKENS ──────────────────────────────────────────────',
        '',
    ]

    slide = tok["slide"]
    lines.append(f'SLIDE_WIDTH  = Inches({slide["WIDTH_IN"]})')
    lines.append(f'SLIDE_HEIGHT = Inches({slide["HEIGHT_IN"]})')
    lines.append(f'SLIDE_MARGIN = Inches({slide["MARGIN_IN"]})')
    lines.append('')

    # Colors (same as tokens.py but using pptx RGBColor)
    lines.append('# ── COLOR TOKENS ──────────────────────────────────────────────')
    lines.append('')
    for group_name, group in tok["color"].items():
        lines.append(f'# {group_name}')
        for name, val in group.items():
            lines.append(f'RGB_{name} = {hex_to_rgb_expr(val)}')
        lines.append('')

    # Fonts
    lines.append('# ── FONT TOKENS ───────────────────────────────────────────────')
    lines.append('')
    for name, val in tok["font"]["family"].items():
        lines.append(f'FONT_{name} = "{val}"')
    lines.append('')

    # Slide sizes
    lines.append('# ── SIZE TOKENS ───────────────────────────────────────────────')
    lines.append('')
    for name, val in slide["size_pt"].items():
        lines.append(f'SIZE_{name} = Pt({val})')
    lines.append('')

    # Glyphs
    lines.append('# ── GLYPH TOKENS ──────────────────────────────────────────────')
    lines.append('')
    for name, val in tok.get("glyph", {}).items():
        if name.startswith("_"):
            continue
        lines.append(f'GLYPH_{name} = {val!r}')
    lines.append('')

    # Content-length budgets
    lines.append('# ── CONTENT LIMIT TOKENS ──────────────────────────────────────')
    lines.append('')
    for name, val in tok.get("content_limit", {}).items():
        if name.startswith("_") or not name.startswith("SLIDE_"):
            continue
        lines.append(f'LIMIT_{name} = {val}')
    lines.append('')

    return '\n'.join(lines) + '\n'


# ── xlsx.py constants ───────────────────────────────────────────────

def build_xlsx_constants(tok: dict) -> str:
    lines = [
        '# -*- coding: utf-8 -*-',
        '"""',
        'Auto-generated from tokens.json — DO NOT EDIT BY HAND.',
        'Run: python scripts/build_tokens.py',
        '"""',
        '',
        '',
        '# ── COLOR TOKENS ──────────────────────────────────────────────',
        '',
    ]

    # XLSX uses hex without '#'
    brand = tok["color"]["brand"]
    light = tok["color"]["light"]
    semantic = tok["color"]["semantic"]

    lines.append('# Brand')
    for name, val in brand.items():
        lines.append(f'XLSX_{name} = "{hex_clean(val)}"')
    lines.append('')

    lines.append('# Layout')
    lines.append(f'XLSX_HEADER_BG    = "{hex_clean(light["TEXT_PRIMARY"])}"')
    lines.append(f'XLSX_SUBHEADER_BG = "{hex_clean(light["TEXT_SECONDARY"])}"')
    lines.append(f'XLSX_TOTAL_BG     = "{hex_clean(light["BG_TERTIARY"])}"')
    lines.append(f'XLSX_BAND         = "{hex_clean(light["BG_SECONDARY"])}"')
    lines.append(f'XLSX_GRID         = "{hex_clean(light["BG_QUATERNARY"])}"')
    lines.append('')

    lines.append('# Text')
    for name, val in light.items():
        if name.startswith("TEXT") or name.startswith("BG_PRIMARY"):
            lines.append(f'XLSX_{name} = "{hex_clean(val)}"')
    lines.append('')

    lines.append('# Semantic')
    for name, val in semantic.items():
        lines.append(f'XLSX_{name} = "{hex_clean(val)}"')
    lines.append('')

    # Fonts
    lines.append('# ── FONT TOKENS ───────────────────────────────────────────────')
    lines.append('')
    lines.append(f'FONT_HEADER  = "{tok["font"]["family"]["XLSX"]}"')
    lines.append(f'FONT_BODY    = "{tok["font"]["family"]["XLSX"]}"')
    lines.append(f'FONT_NUMERIC = "{tok["font"]["family"]["MONO"]}"')
    lines.append(f'FONT_KPI     = "{tok["font"]["family"]["DISPLAY_EN"]}"')
    lines.append(f'FONT_NOTE    = "{tok["font"]["family"]["XLSX"]}"')
    lines.append('')

    # Number formats
    lines.append('# ── NUMBER FORMAT TOKENS ──────────────────────────────────────')
    lines.append('')
    for name, val in tok["xlsx"]["number_format"].items():
        escaped = val.replace('"', '\\"')
        lines.append(f'FMT_{name} = "{escaped}"')
    lines.append('')

    return '\n'.join(lines) + '\n'


# ── Main ────────────────────────────────────────────────────────────

def main():
    check = "--check" in sys.argv
    tok = load_tokens()

    outputs = {
        ROOT / "tokens_constants.py": build_tokens_constants(tok),
        ROOT / "slides_constants.py": build_slides_constants(tok),
        ROOT / "xlsx_constants.py": build_xlsx_constants(tok),
    }

    if check:
        ok = True
        for path, content in outputs.items():
            if not path.exists():
                print(f"MISSING: {path}")
                ok = False
            elif path.read_text(encoding="utf-8") != content:
                print(f"STALE:   {path}")
                ok = False
            else:
                print(f"OK:      {path}")
        sys.exit(0 if ok else 1)
    else:
        for path, content in outputs.items():
            path.write_text(content, encoding="utf-8")
            print(f"Wrote {path}")


if __name__ == "__main__":
    main()
