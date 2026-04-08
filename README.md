
# 🔥 Hermes Easy Office Suite
> 全能Python办公自动化工具包，单个API搞定Word/Excel/PDF/PPT四种格式，中文友好，开箱即用，支持与Hermes AI智能体无缝集成。

---

## ✨ 核心功能
| 功能 | 支持 |
|------|------|
| 📄 Word文档 | 创建、编辑、合并、水印、模板、格式转换 |
| 📊 Excel表格 | 创建、数据分析、图表生成、数据导入导出、批量处理 |
| 📑 PDF文档 | 生成、合并、拆分、内容提取、水印、格式转换 |
| 🎨 PowerPoint | 演示文稿创建、设计、动画、模板、批量生成 |
| 🔄 格式互转 | 支持4种Office格式一键互转 |
| 📦 批量处理 | 整个目录文件批量处理、格式转换 |
| ⚙️ 工作流 | YAML配置自动化工作流，日常报告自动生成 |
| 📧 邮件集成 | 生成的文档自动发送到指定邮箱 |
| 🛡️ 安全可靠 | 所有处理本地完成，无数据上传泄露风险 |
| 🇨🇳 中文友好 | 彻底解决中文乱码问题，全中文文档 |
| 🤖 AI集成 | 原生支持作为Hermes AI技能调用，适配AI工作流 |

---

## 🚀 快速开始
### 1. 安装
```bash
# 安装核心依赖
pip install python-docx openpyxl reportlab PyPDF2 python-pptx pandas Pillow

# 可选：安装完整功能依赖
pip install Jinja2 click rich
```

### 2. 3行代码生成文档
```python
from office_suite import OfficeSuite

# 初始化套件
suite = OfficeSuite()

# 创建Word报告
suite.create("word", 
         title="2025年第一季度销售报告", 
         content="报告内容...",
         output="季度报告.docx")

# 一键转换为PDF
suite.convert("季度报告.docx", to="pdf", output="季度报告.pdf")
```

### 3. 发送邮件
```python
# 配置邮箱（密码从环境变量读取，安全无泄露）
suite.config_email(smtp_server="smtp.qq.com", 
                smtp_port=465,
                username="your@email.com")

# 发送邮件
suite.send_email(to="recipient@email.com", 
                subject="季度报告",
                attachments=["季度报告.pdf"])
```

---

## 📚 使用示例
| 示例文件 | 功能 |
|----------|------|
| [examples/basic_usage.py](examples/basic_usage.py) | 基础使用方法 |
| [examples/batch_process.py](examples/batch_process.py) | 批量处理文件 |
| [examples/generate_report.py](examples/generate_report.py) | 生成完整报告 |
| [examples/workflow_example.yaml](examples/workflow_example.yaml) | 自动化工作流配置 |

---

## 🛡️ 安全说明
⚠️ 重要提醒：
1. **所有文档处理都在本地完成**，不会上传任何数据到第三方服务器，确保敏感文档安全
2. **密码/授权码切勿硬编码**，请使用环境变量存储敏感信息
3. 配置文件请设置权限为`600`，仅所有者可读写

---

## 📋 系统要求
- Python 3.8+
- Windows/macOS/Linux均可
- 可选：安装LibreOffice（用于高质量格式转换）

---

## 🤝 贡献
欢迎提交Issue和PR！有功能建议或者Bug反馈欢迎提出~

---

## 🙏 致谢
本项目的设计思路受到 **[MiniMax](https://www.minimaxi.com/)** 官方 Office Skill 的启发，在此基础上进行了全面的功能扩展和优化：
1. 适配为 Hermes AI 智能体的可调用技能套件，支持与 AI 工作流无缝集成
2. 增加了批量处理、工作流自动化、邮件集成等高级功能
3. 优化了中文支持，彻底解决中文乱码问题
4. 提供统一API接口，降低使用门槛

感谢 MiniMax 团队提供的优秀产品思路参考！

---

## 📄 许可证
MIT License | 详见 [LICENSE](LICENSE)
