
#!/usr/bin/env python3
"""
基础使用示例：创建不同类型的Office文档
"""

from office_suite import OfficeSuite
import os

def main():
    print("=== Office Suite 基础使用示例 ===")
    
    # 初始化套件
    suite = OfficeSuite()
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 创建Word文档
    print("\n1. 正在创建Word文档...")
    word_path = os.path.join(output_dir, "测试报告.docx")
    suite.create("word",
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
3. 推出3款新产品
        """,
        output_path=word_path
    )
    print(f"✅ Word文档已保存到: {word_path}")
    
    # 2. 创建Excel表格
    print("\n2. 正在创建Excel表格...")
    excel_path = os.path.join(output_dir, "销售数据.xlsx")
    sales_data = [
        ["部门", "1月销售额", "2月销售额", "3月销售额", "季度总额", "完成率"],
        ["销售一部", 150, 140, 160, 450, "112.5%"],
        ["销售二部", 120, 130, 130, 380, "95%"],
        ["销售三部", 110, 120, 140, 370, "92.5%"]
    ]
    suite.create("excel",
        title="2025年第一季度销售数据",
        data=sales_data,
        output_path=excel_path,
        create_chart=True
    )
    print(f"✅ Excel表格已保存到: {excel_path}")
    
    # 3. 创建PDF文档
    print("\n3. 正在创建PDF文档...")
    pdf_path = os.path.join(output_dir, "销售报告.pdf")
    suite.create("pdf",
        title="2025年第一季度销售报告",
        content="""
销售情况概述
第一季度总销售额达到1200万元，同比增长25%，超额完成季度目标。

各部门业绩
- 销售一部：450万元，完成率112.5%
- 销售二部：380万元，完成率95%
- 销售三部：370万元，完成率92.5%

后续计划
1. 第二季度目标设定为1500万元
2. 重点拓展华东和华南区域市场
3. 推出3款新产品
        """,
        output_path=pdf_path
    )
    print(f"✅ PDF文档已保存到: {pdf_path}")
    
    # 4. 创建PPT演示文稿
    print("\n4. 正在创建PPT演示文稿...")
    ppt_path = os.path.join(output_dir, "销售报告演示.pptx")
    slides = [
        {
            "title": "2025年第一季度销售报告",
            "content": "汇报人：市场部\n汇报时间：2025年4月",
            "layout": "title"
        },
        {
            "title": "销售情况概述",
            "content": "• 总销售额：1200万元\n• 同比增长：25%\n• 完成率：108%",
            "layout": "content"
        },
        {
            "title": "各部门业绩",
            "content": "• 销售一部：450万元，完成率112.5%\n• 销售二部：380万元，完成率95%\n• 销售三部：370万元，完成率92.5%",
            "layout": "content"
        },
        {
            "title": "后续计划",
            "content": "• 第二季度目标：1500万元\n• 重点拓展区域：华东、华南\n• 新产品：3款",
            "layout": "content"
        }
    ]
    suite.create("pptx",
        title="第一季度销售报告",
        slides=slides,
        output_path=ppt_path
    )
    print(f"✅ PPT演示文稿已保存到: {ppt_path}")
    
    # 5. 格式转换示例：Word转PDF
    print("\n5. 正在转换Word到PDF...")
    converted_pdf = os.path.join(output_dir, "Word转换后的报告.pdf")
    suite.convert(word_path, to="pdf", output_path=converted_pdf)
    print(f"✅ 格式转换完成: {converted_pdf}")
    
    print("\n🎉 所有示例运行完成！")
    print(f"📁 生成的文件都保存在: {os.path.abspath(output_dir)}")

if __name__ == "__main__":
    main()
