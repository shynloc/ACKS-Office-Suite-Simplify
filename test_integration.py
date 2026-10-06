#!/usr/bin/env python3
"""
集成测试：按三套内置主题生成四种格式，读取、加水印、合并。
运行：python test_integration.py
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

import acks_office  # noqa: E402

CONTENT = "---\ntitle: 测试报告\nauthor: 测试\n---\n\n# 概述\n\n正文内容\n\n| 区域 | 营收 |\n|---|--:|\n| 华东 | 5,888 |\n\n- 项目一\n- 项目二"
SLIDES = [{"layout": "title", "title": "季度报告", "subtitle": "汇报人：市场部"},
          {"layout": "section", "number": "01", "title": "概述"},
          {"title": "概述", "bullets": ["总销售额 1200 万元", "同比增长 25%"]}]
DATA = [["部门", "1月", "2月"], ["一部", 150, 140], ["二部", 120, 130]]


def main():
    tmp = tempfile.mkdtemp(prefix="acks_office_test_")
    results = []

    def check(name, action):
        # Windows CI 的重定向标准输出可能是 CP1252，测试日志使用 ASCII。
        try:
            ok = bool(action())
        except Exception as exc:
            print(f"[FAIL] {name}: {type(exc).__name__}: {exc}".encode("ascii", "replace").decode("ascii"))
            ok = None
        else:
            print(f"{'[PASS]' if ok else '[FAIL]'} {name}")
        results.append((name, ok))

    def path(name):
        return os.path.join(tmp, name)

    for theme in ("neutral", "slate", "folio"):
        check(f"create word ({theme})", lambda t=theme: acks_office.create(
            "word", path(f"{t}.docx"), content=CONTENT, theme=t)["theme"] == t)
        check(f"create pdf ({theme})", lambda t=theme: acks_office.create(
            "pdf", path(f"{t}.pdf"), content=CONTENT, theme=t)["pages"] >= 1)
        check(f"create pptx ({theme})", lambda t=theme: acks_office.create(
            "pptx", path(f"{t}.pptx"), title="季度报告", slides=SLIDES, theme=t)["slides_count"] == 3)
        check(f"create excel ({theme})", lambda t=theme: acks_office.create(
            "excel", path(f"{t}.xlsx"), data=DATA, theme=t)["rows"] == 2)

    check("default is an alias of neutral", lambda: acks_office.create(
        "word", path("default.docx"), content=CONTENT, theme="default")["theme"] == "neutral")
    check("acks style still available", lambda: acks_office.create(
        "word", path("acks.docx"), title="测试", content="# 标题\n\n内容", theme="acks")["theme"] == "acks")
    check("extract text (docx)", lambda: "正文内容" in acks_office.extract(path("slate.docx")))
    check("extract data (xlsx)", lambda: acks_office.extract(path("neutral.xlsx")) == [
        {"部门": "一部", "1月": 150, "2月": 140}, {"部门": "二部", "1月": 120, "2月": 130}])
    check("pdf watermark keeps the source", lambda: acks_office.add_watermark(path("slate.pdf"), "机密文件")
          ["output_path"].endswith("slate_watermarked.pdf"))
    check("merge pdfs", lambda: acks_office.merge([path("slate.pdf"), path("slate_watermarked.pdf")],
                                                   path("merged.pdf"))["merged_pages"] >= 2)

    passed = sum(1 for _, ok in results if ok)
    print(f"\n{'=' * 40}")
    print(f"Passed {passed}/{len(results)} integration checks")
    if passed != len(results):
        print("Integration checks failed")
        sys.exit(1)
    print("All integration checks passed")


if __name__ == "__main__":
    main()
