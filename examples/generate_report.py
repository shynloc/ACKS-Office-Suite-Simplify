#!/usr/bin/env python3
"""
生成完整报告示例：用 ACKS 设计规范主题生成 Word 报告并转换为 PDF。
"""

import os
from office_suite import OfficeSuite


def main():
    print("=== Office Suite 完整报告生成示例 ===")

    suite = OfficeSuite(theme="acks")  # 默认 ACKS 设计规范主题
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)

    # 1. 生成规范化的 Word 报告
    report_path = os.path.join(output_dir, "季度报告.docx")
    result = suite.create(
        "word",
        title="2025年第一季度销售报告",
        content="""
# 销售情况概述

第一季度总销售额达到 1200 万元，同比增长 25%，超额完成季度目标。

## 各部门业绩

- 销售一部：450 万元，完成率 112.5%
- 销售二部：380 万元，完成率 95%
- 销售三部：370 万元，完成率 92.5%

## 后续计划

1. 第二季度目标设定为 1500 万元
2. 重点拓展华东和华南区域市场
3. 推出 3 款新产品
        """,
        output_path=report_path,
    )
    print(f"✅ Word 报告已生成: {report_path}（主题: {result.get('theme')}）")

    # 2. 转换为 PDF（走 LibreOffice；未安装则跳过）
    pdf_path = os.path.join(output_dir, "季度报告.pdf")
    try:
        converted = suite.convert(report_path, to="pdf", output_path=pdf_path)
        print(f"✅ PDF 已生成: {converted.get('output_path')}")
    except Exception as e:
        print(f"⚠️ PDF 转换跳过（可能未安装 LibreOffice）: {e}")

    print("\n🎉 报告生成完成！")


if __name__ == "__main__":
    main()
