#!/usr/bin/env python3
"""
生成完整报告示例：用 slate 主题生成 Word 与 PDF 报告（PDF 由引擎直接排版，不需要 LibreOffice）。
"""

import os

import acks_office

REPORT = """---
title: 2025 年第一季度\\n销售报告
kicker: 销售分析
author: 销售管理部
date: 2025 年 4 月 8 日
version: v1.0
---

# 销售情况概述

第一季度总销售额达到 1200 万元，同比增长 25%，超额完成季度目标。

## 各部门业绩

表：各部门销售额（单位：万元）

| 部门 | 销售额 | 完成率 |
|---|--:|--:|
| 销售一部 | 450 | 112.5% |
| 销售二部 | 380 | 95.0% |
| 销售三部 | 370 | 92.5% |
| 合计 | 1,200 | 100.0% |

> [!TIP]
> 二部、三部的差距主要来自新客户拓展，第二季度重点补齐。

## 后续计划

1. 第二季度目标设定为 1500 万元
2. 重点拓展华东和华南区域市场
3. 推出 3 款新产品
"""


def main():
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)
    for kind, name in (("word", "季度报告.docx"), ("pdf", "季度报告.pdf")):
        path = os.path.join(output_dir, name)
        result = acks_office.create(kind, path, content=REPORT, theme="slate")
        print(f"已生成 {path}（主题：{result['theme']}）")
        for warning in result.get("warnings", []):
            print(f"  提醒：{warning['message']}")


if __name__ == "__main__":
    main()
