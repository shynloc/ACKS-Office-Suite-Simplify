
#!/usr/bin/env python3
"""
批量处理示例：批量转换整个目录的Word文档为PDF
"""

from office_suite import OfficeSuite
import os

def main():
    print("=== Office Suite 批量处理示例 ===")
    
    # 初始化套件
    suite = OfficeSuite()
    
    # 源目录和目标目录
    source_dir = "./word_docs"  # 存放Word文档的目录
    target_dir = "./pdf_docs"   # 转换后的PDF保存目录
    os.makedirs(source_dir, exist_ok=True)
    os.makedirs(target_dir, exist_ok=True)
    
    # 批量转换
    print(f"\n📂 正在批量转换目录 {source_dir} 中的Word文档到PDF...")
    results = suite.batch_convert(
        source_dir=source_dir,
        target_dir=target_dir,
        source_format="docx",
        target_format="pdf"
    )
    
    # 打印结果
    print(f"\n✅ 批量处理完成，共处理 {len(results)} 个文件：")
    for filename, status in results.items():
        if status["success"]:
            print(f"✅ {filename} -> {status['output_path']}")
        else:
            print(f"❌ {filename} 转换失败: {status['error']}")
    
    # 批量添加水印示例
    print(f"\n🖋️ 正在批量给PDF添加水印...")
    watermark_results = suite.batch_add_watermark(
        source_dir=target_dir,
        target_dir="./pdf_with_watermark",
        watermark_text="内部资料 请勿外传",
        position="center",
        opacity=0.3
    )
    
    print(f"\n✅ 水印添加完成，共处理 {len(watermark_results)} 个文件：")
    for filename, status in watermark_results.items():
        if status["success"]:
            print(f"✅ {filename} -> {status['output_path']}")
        else:
            print(f"❌ {filename} 添加水印失败: {status['error']}")
    
    print("\n🎉 批量处理示例运行完成！")

if __name__ == "__main__":
    main()
