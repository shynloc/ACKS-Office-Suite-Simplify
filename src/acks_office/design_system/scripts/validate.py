#!/usr/bin/env python3
"""
Validation script for ACKS Studio Design System v2.0.

Usage:
    python scripts/validate.py
"""

from __future__ import annotations

import json
import sys
import os
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Try to find a Python that has the required packages
PYTHON_CANDIDATES = [
    sys.executable,
    "/usr/local/bin/python3",
    "/usr/bin/python3",
]
PYTHON = sys.executable
for p in PYTHON_CANDIDATES:
    try:
        result = subprocess.run(
            [p, "-c", "import docx, pptx, openpyxl"],
            capture_output=True, timeout=10
        )
        if result.returncode == 0:
            PYTHON = p
            break
    except Exception:
        continue

errors = []


def check(name: str, ok: bool, msg: str = ""):
    status = "PASS" if ok else "FAIL"
    print(f"  [{status}] {name}" + (f" — {msg}" if msg else ""))
    if not ok:
        errors.append(f"{name}: {msg}")


def main():
    print("ACKS Studio Design System v2.1 — Validation")
    print("=" * 50)

    # 1. tokens.json
    print("\n1. tokens.json")
    tokens_path = ROOT / "tokens.json"
    try:
        with open(tokens_path, encoding="utf-8") as f:
            tok = json.load(f)
        check("Parse tokens.json", True)
        for key in ("color", "font", "spacing_pt", "page", "slide", "xlsx"):
            check(f"  Has key '{key}'", key in tok)
    except Exception as e:
        check("Parse tokens.json", False, str(e))

    # 2. Build constants
    print("\n2. Generated constants")
    build_script = ROOT / "scripts" / "build_tokens.py"
    try:
        result = subprocess.run(
            [PYTHON, str(build_script)],
            capture_output=True, text=True, cwd=str(ROOT)
        )
        check("build_tokens.py runs", result.returncode == 0,
              result.stderr.strip() if result.returncode else "")
    except Exception as e:
        check("build_tokens.py runs", False, str(e))

    try:
        result = subprocess.run(
            [PYTHON, str(build_script), "--check"],
            capture_output=True, text=True, cwd=str(ROOT)
        )
        check("Constants in sync (--check)", result.returncode == 0,
              result.stdout.strip() if result.returncode else "")
    except Exception as e:
        check("Constants in sync", False, str(e))

    # 3. Module imports
    print("\n3. Module imports")
    for mod_name in ("tokens", "slides", "xlsx"):
        try:
            result = subprocess.run(
                [PYTHON, "-c", f"import {mod_name}; print('OK')"],
                capture_output=True, text=True, cwd=str(ROOT)
            )
            check(f"import {mod_name}", result.returncode == 0,
                  result.stderr.strip() if result.returncode else "")
        except Exception as e:
            check(f"import {mod_name}", False, str(e))

    # 4. Demo generation
    print("\n4. Demo generation")
    demos = [
        ("tokens.save_demo", "ACKS-demo.docx"),
        ("slides.save_demo", "ACKS-deck.pptx"),
        ("xlsx.save_demo",   "ACKS-workbook.xlsx"),
    ]
    for func_path, filename in demos:
        try:
            demo_file = ROOT / filename
            if demo_file.exists():
                demo_file.unlink()
            code = f"from {func_path.split('.')[0]} import {func_path.split('.')[1]}; {func_path.split('.')[1]}('{filename}')"
            result = subprocess.run(
                [PYTHON, "-c", code],
                capture_output=True, text=True, cwd=str(ROOT)
            )
            if result.returncode == 0 and demo_file.exists():
                size = demo_file.stat().st_size
                check(f"{func_path}() → {filename}", size > 0,
                      f"{size:,} bytes")
            else:
                check(f"{func_path}()", False,
                      result.stderr.strip() if result.returncode else "File not created")
        except Exception as e:
            check(f"{func_path}()", False, str(e))

    # 5. __all__ completeness
    print("\n5. Public API (__all__)")
    for mod_name in ("tokens", "slides", "xlsx"):
        try:
            code = f"import {mod_name}; print(len({mod_name}.__all__))"
            result = subprocess.run(
                [PYTHON, "-c", code],
                capture_output=True, text=True, cwd=str(ROOT)
            )
            if result.returncode == 0:
                count = result.stdout.strip()
                check(f"{mod_name}.__all__ defined", True,
                      f"{count} exports")
            else:
                check(f"{mod_name}.__all__", False,
                      result.stderr.strip())
        except Exception as e:
            check(f"{mod_name}.__all__", False, str(e))

    # 6. Slide geometry — static bounds/overflow check
    print("\n6. Slide geometry (pptx)")
    try:
        code = r'''
import sys
from pptx.util import Emu
from slides import save_demo
save_demo("_geom_check.pptx")
from pptx import Presentation
prs = Presentation("_geom_check.pptx")
slide_w = prs.slide_width
slide_h = prs.slide_height
footer_safe_top = Emu(int(6.75 * 914400))
problems = []
for i, slide in enumerate(prs.slides, start=1):
    for shape in slide.shapes:
        try:
            l, t, w, h = shape.left, shape.top, shape.width, shape.height
        except Exception:
            continue
        if None in (l, t, w, h):
            continue
        right = l + w
        bottom = t + h
        if l < 0 or t < 0:
            problems.append(f"slide {i}: shape '{shape.name}' has negative origin ({l},{t})")
        if right > slide_w + Emu(9144):
            problems.append(f"slide {i}: shape '{shape.name}' right edge {right/914400:.2f}in exceeds slide width {slide_w/914400:.2f}in")
        if bottom > slide_h + Emu(9144):
            problems.append(f"slide {i}: shape '{shape.name}' bottom edge {bottom/914400:.2f}in exceeds slide height {slide_h/914400:.2f}in")
        # footer-zone intrusion: any non-footer shape starting below the
        # safe content line and extending into where the footer rule/
        # page-number sit is a strong signal of overflow text.
        is_footer_shape = (shape.name or "").startswith("acks_footer") or (shape.top is not None and shape.top >= footer_safe_top - Emu(int(0.05*914400)) and (shape.name or "").lower().find("footer") >= 0)
        if not is_footer_shape and t < footer_safe_top and bottom > footer_safe_top + Emu(int(0.35*914400)):
            problems.append(f"slide {i}: shape '{shape.name}' spans into footer-safe zone (bottom {bottom/914400:.2f}in)")
if problems:
    for p in problems:
        print("PROBLEM: " + p)
    sys.exit(1)
else:
    print("geometry OK — " + str(len(prs.slides._sldIdLst)) + " slides checked")
'''
        result = subprocess.run(
            [PYTHON, "-c", code],
            capture_output=True, text=True, cwd=str(ROOT)
        )
        out = result.stdout.strip()
        check("No shapes out of slide bounds / footer-zone overflow",
              result.returncode == 0,
              out if result.returncode else "")
    except Exception as e:
        check("Slide geometry check", False, str(e))

    # Summary
    print("\n" + "=" * 50)
    if errors:
        print(f"FAILED — {len(errors)} error(s):")
        for e in errors:
            print(f"  ✗ {e}")
        sys.exit(1)
    else:
        print("ALL CHECKS PASSED")
        sys.exit(0)


if __name__ == "__main__":
    main()
