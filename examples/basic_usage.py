#!/usr/bin/env python3
"""Generate four formats using illustrative data; check every API result."""

from pathlib import Path
from acks_office import OfficeSuite


def require_success(result):
    if not result.get("success"):
        raise RuntimeError(result.get("error", "Operation failed"))
    return result


def main():
    output = Path("output")
    output.mkdir(exist_ok=True)
    suite = OfficeSuite()
    content = "# Overview\n\nIllustrative sales data, not current business results.\n\n- Revenue: 1200"
    for kind, suffix, options in (
        ("word", "docx", {"content": content}),
        ("pdf", "pdf", {"content": content}),
        ("excel", "xlsx", {"theme": "default", "create_chart": True,
                          "data": [["Department", "Sales"], ["East", 450], ["South", 380]]}),
        ("pptx", "pptx", {"slides": [
            {"title": "Illustrative report", "layout": "title"},
            {"title": "Revenue", "content": "1200", "layout": "content"},
        ]}),
    ):
        path = output / ("example." + suffix)
        require_success(suite.create(kind, title="Illustrative report", output_path=str(path), **options))
        require_success(suite.extract_data(str(path)))
        print("[PASS] Created and read " + str(path))
    result = suite.convert(str(output / "example.docx"), to="pdf", output_path=str(output / "converted.pdf"))
    if result.get("success"):
        print("[PASS] LibreOffice conversion: " + result["output_path"])
    else:
        print("[SKIP] LibreOffice conversion: " + result["error"])


if __name__ == "__main__":
    main()
