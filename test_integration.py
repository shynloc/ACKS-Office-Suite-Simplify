#!/usr/bin/env python3
"""
Office Suite 集成测试：验证所有核心功能。
运行：python test_integration.py
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from office_suite import OfficeSuite  # noqa: E402


def main():
    suite = OfficeSuite(theme="acks")
    tmp = tempfile.mkdtemp(prefix="office_suite_test_")
    results = []

    def check(name, cond):
        results.append((name, cond))
        print(f"{'✅' if cond else '❌'} {name}")

    # 1. Word（ACKS 主题）
    word_path = os.path.join(tmp, "report.docx")
    r = suite.create("word", title="测试报告", content="# 概述\n\n正文内容\n\n- 项目一\n- 项目二",
                     output_path=word_path)
    check("create word (acks)", r.get("success") and os.path.exists(word_path))

    # 2. Word（default 主题）
    word_def = os.path.join(tmp, "report_default.docx")
    r = suite.create("word", title="测试", content="# 标题\n\n内容", output_path=word_def, theme="default")
    check("create word (default)", r.get("success") and os.path.exists(word_def))

    # 3. Excel（ACKS 主题）
    excel_path = os.path.join(tmp, "data.xlsx")
    r = suite.create("excel", title="销售数据",
                     data=[["部门", "1月", "2月"], ["一部", 150, 140], ["二部", 120, 130]],
                     output_path=excel_path)
    check("create excel (acks)", r.get("success") and os.path.exists(excel_path))

    # 4. PDF
    pdf_path = os.path.join(tmp, "report.pdf")
    r = suite.create("pdf", title="测试报告", content="# 概述\n\nPDF 内容", output_path=pdf_path)
    check("create pdf", r.get("success") and os.path.exists(pdf_path))

    # 5. PPT（ACKS 主题）
    ppt_path = os.path.join(tmp, "deck.pptx")
    r = suite.create("pptx", title="季度报告",
                     slides=[{"title": "季度报告", "content": "汇报人：市场部", "layout": "title"},
                             {"title": "概述", "content": "• 总销售额 1200 万元\n• 同比增长 25%", "layout": "content"}],
                     output_path=ppt_path)
    check("create pptx (acks)", r.get("success") and os.path.exists(ppt_path))

    # 6. 提取文本
    r = suite.extract_data(word_path)
    check("extract text (docx)", r.get("success") and isinstance(r.get("data"), str))

    # 7. 提取数据（Excel，默认读取第一个工作表）
    r = suite.extract_data(excel_path)
    check("extract data (xlsx)", r.get("success") and r.get("data") == [
        {"部门": "一部", "1月": 150, "2月": 140},
        {"部门": "二部", "1月": 120, "2月": 130},
    ])

    # 8. PDF 水印
    wm_path = os.path.join(tmp, "report_wm.pdf")
    r = suite.add_watermark(pdf_path, "机密文件", output_path=wm_path)
    check("pdf watermark", r.get("success") and os.path.exists(wm_path))

    # 9. 合并 PDF
    merged_path = os.path.join(tmp, "merged.pdf")
    from office_suite.pdf import merge_pdfs
    r = merge_pdfs([pdf_path, wm_path], merged_path)
    check("merge pdfs", os.path.exists(merged_path))

    # 汇总
    passed = sum(1 for _, ok in results if ok)
    total = len(results)
    print(f"\n{'='*40}")
    print(f"通过 {passed}/{total} 项测试")
    if passed == total:
        print("✅ 所有功能测试通过")
    else:
        print("❌ 存在失败项")
        sys.exit(1)


if __name__ == "__main__":
    main()
