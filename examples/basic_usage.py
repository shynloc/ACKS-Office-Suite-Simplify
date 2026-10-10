#!/usr/bin/env python3
"""Generate the four formats with illustrative data and read each file back.

The functions raise an exception when something goes wrong, so there is no result dict to check.
"""

from pathlib import Path

import acks_office


def main():
    output = Path("output")
    output.mkdir(exist_ok=True)
    content = "# Overview\n\nIllustrative sales data, not current business results.\n\n- Revenue: 1200"
    for kind, suffix, options in (
        ("word", "docx", {"content": content}),
        ("pdf", "pdf", {"content": content}),
        ("excel", "xlsx", {"create_chart": True,
                          "data": [["Department", "Sales"], ["East", 450], ["South", 380]]}),
        ("pptx", "pptx", {"slides": [
            {"title": "Illustrative report", "layout": "title"},
            {"title": "Revenue", "bullets": ["East 450", "South 380"]},
        ]}),
    ):
        path = output / ("example." + suffix)
        acks_office.create(kind, str(path), title="Illustrative report", theme="slate", **options)
        acks_office.extract(str(path))
        print("[PASS] Created and read " + str(path))
    try:
        result = acks_office.convert(str(output / "example.docx"), "pdf", str(output / "converted.pdf"))
        print("[PASS] LibreOffice conversion: " + result["output_path"])
    except Exception as exc:  # LibreOffice is optional
        print("[SKIP] LibreOffice conversion: " + str(exc))


if __name__ == "__main__":
    main()
