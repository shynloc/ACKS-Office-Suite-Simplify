#!/usr/bin/env python3
"""
批量处理示例：把目录里的 Word 文档转成 PDF（需要 LibreOffice），再给 PDF 加水印。
每个文件单独处理，一个出错不影响其他文件。
"""

from pathlib import Path

import acks_office


def main():
    source_dir, pdf_dir, stamped_dir = Path("word_docs"), Path("pdf_docs"), Path("pdf_with_watermark")
    for directory in (source_dir, pdf_dir, stamped_dir):
        directory.mkdir(exist_ok=True)

    for docx in sorted(source_dir.glob("*.docx")):
        target = pdf_dir / (docx.stem + ".pdf")
        try:
            acks_office.convert(str(docx), "pdf", str(target))
            print(f"✅ {docx.name} -> {target}")
        except Exception as exc:
            print(f"❌ {docx.name} 转换失败：{exc}")

    for pdf in sorted(pdf_dir.glob("*.pdf")):
        target = stamped_dir / pdf.name
        try:
            acks_office.add_watermark(str(pdf), "内部资料 请勿外传", str(target), position="center", opacity=0.3)
            print(f"✅ {pdf.name} -> {target}")
        except Exception as exc:
            print(f"❌ {pdf.name} 添加水印失败：{exc}")


if __name__ == "__main__":
    main()
