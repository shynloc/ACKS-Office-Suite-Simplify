#!/usr/bin/env python3
"""
ACKS Studio Design System — CLI entry point.

Usage:
    python -m acks build docx    # Generate .docx demo
    python -m acks build pptx    # Generate .pptx demo
    python -m acks build xlsx    # Generate .xlsx workbook
    python -m acks build all     # Generate all demos
    python -m acks validate      # Run validation checks
    python -m acks tokens check  # Check token sync
    python -m acks tokens build  # Rebuild *_constants.py from tokens.json
    python -m acks fonts download # Download Google Fonts for offline use
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def cmd_build(args):
    """Build demo documents."""
    targets = args.targets
    if "all" in targets:
        targets = ["docx", "pptx", "xlsx"]

    for t in targets:
        if t == "docx":
            from tokens import save_demo
            out = save_demo("ACKS-demo.docx")
            print(f"  {out}")
        elif t == "pptx":
            from slides import save_demo
            out = save_demo("ACKS-deck.pptx")
            print(f"  {out}")
        elif t == "xlsx":
            from xlsx import save_demo
            out = save_demo("ACKS-workbook.xlsx")
            print(f"  {out}")


def cmd_validate(args):
    """Run validation."""
    script = ROOT / "scripts" / "validate.py"
    result = subprocess.run([sys.executable, str(script)])
    sys.exit(result.returncode)


def cmd_tokens(args):
    """Token management."""
    script = ROOT / "scripts" / "build_tokens.py"
    if args.action == "build":
        subprocess.run([sys.executable, str(script)])
    elif args.action == "check":
        result = subprocess.run([sys.executable, str(script), "--check"])
        sys.exit(result.returncode)


def cmd_fonts(args):
    """Font management."""
    script = ROOT / "scripts" / "download_fonts.py"
    if args.action == "download":
        subprocess.run([sys.executable, str(script)])


def main():
    parser = argparse.ArgumentParser(
        prog="acks",
        description="ACKS Studio Design System v2.0 — CLI",
    )
    sub = parser.add_subparsers(dest="command")

    # build
    p_build = sub.add_parser("build", help="Build demo documents")
    p_build.add_argument(
        "targets", nargs="+", choices=["docx", "pptx", "xlsx", "all"],
        help="Which format(s) to build"
    )
    p_build.set_defaults(func=cmd_build)

    # validate
    p_val = sub.add_parser("validate", help="Run validation checks")
    p_val.set_defaults(func=cmd_validate)

    # tokens
    p_tok = sub.add_parser("tokens", help="Token management")
    p_tok.add_argument("action", choices=["build", "check"])
    p_tok.set_defaults(func=cmd_tokens)

    # fonts
    p_fonts = sub.add_parser("fonts", help="Font management")
    p_fonts.add_argument("action", choices=["download"])
    p_fonts.set_defaults(func=cmd_fonts)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    args.func(args)


if __name__ == "__main__":
    main()
