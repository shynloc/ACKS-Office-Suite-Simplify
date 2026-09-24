# 🎨 ACKS Office Suite Simplify

> 全能 Python 办公自动化工具包 —— 单个 API 搞定 Word / Excel / PDF / PPT，内置 ACKS 设计规范，开箱即出规范化文档。

<p align="center">
  <a href="#"><img alt="Python" src="https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white"></a>
  <a href="#"><img alt="Language" src="https://img.shields.io/badge/中文-友好-red"></a>
  <a href="#"><img alt="Platform" src="https://img.shields.io/badge/平台-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey"></a>
  <a href="#"><img alt="License" src="https://img.shields.io/badge/License-MIT-green"></a>
  <a href="#"><img alt="Agent" src="https://img.shields.io/badge/Agent-Hermes%20%7C%20DSH%20%7C%20MCP-orange"></a>
  <a href="#"><img alt="Design" src="https://img.shields.io/badge/设计规范-ACKS%20v2.1-FF6B1A"></a>
</p>

---

## ✨ 特性

| 能力 | 说明 |
|------|------|
| 📄 Word | 创建、编辑、合并、水印、模板、格式转换 |
| 📊 Excel | 创建、数据分析、图表生成、数据导入导出、批量处理 |
| 📑 PDF | 生成、合并、拆分、内容提取、**真水印**、格式转换 |
| 🎨 PowerPoint | 演示文稿创建、设计、动画、模板、批量生成 |
| 🔄 格式互转 | 4 种 Office 格式一键互转（走 LibreOffice） |
| 📦 批量处理 | 整目录文件批量转换、批量水印 |
| ⚙️ 工作流 | YAML 配置自动化工作流，日常报告自动生成 |
| 📧 邮件集成 | 生成的文档自动发送到指定邮箱 |
| 🎨 设计规范 | 内置 ACKS Studio 设计规范 v2.1，产出规范化文档 |
| 🎭 双主题 | `theme="acks"`（规范）与 `theme="default"`（简洁）一键切换 |
| 🇨🇳 中文友好 | 彻底解决中文乱码，全中文文档 |
| 🤖 AI 集成 | 原生适配 Hermes Agent / DSH / MCP，可作 AI 工作流工具 |

---

## 🎯 设计规范（ACKS Studio v2.1）

内置 [ACKS Studio 文档设计规范 v2](src/office_suite/design_system/README.md)，使产出的 Word / PPT / Excel 文档统一规范：

- **橙色 `#FF6B1A` 是唯一强调色**（覆盖 ≤ 5%），`#D4530E` 用于正文链接（WCAG AA）
- **黑色实线主分隔**：章节边界 2pt，段落组 1pt，不用阴影渐变
- **中英双语并置**：中文行高 1.75 / 英文 1.6，Space Grotesk / DM Sans / Noto Sans SC 字体体系
- **组件开箱即用**：封面、章节标题、正文、列表、表格、KPI、引用、callout、代码块
- **Token 单一信源**：`tokens.json` → 自动生成三套常量，`scripts/validate.py` 校验一致性

> 设计规范源码位于 `src/office_suite/design_system/`，来源与同步说明见 `SOURCE.md`。

---

## 📦 安装

### 环境要求

- **Python** ≥ 3.9
- **系统**：Windows 10+ / macOS 10.15+ / Linux（Ubuntu 20.04+ / CentOS 7+ / OpenCloudOS 9+）
- **可选**：LibreOffice（高质量格式转换）、中文字体（Noto Sans SC 等）

### pip 安装

```bash
# 核心依赖（必装）
pip install python-docx openpyxl reportlab PyPDF2 python-pptx pandas xlrd Pillow

# 完整功能（可选）
pip install Jinja2 click rich python-magic chardet tqdm python-dotenv
```

或直接从源码安装：

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
pip install -r requirements.txt
pip install -e .
```

### 系统工具（高质量转 PDF）

```bash
# Ubuntu / Debian
sudo apt install -y libreoffice fonts-noto-cjk

# macOS
brew install libreoffice

# Windows：下载安装 LibreOffice 与中文字体
```

---

## 🚀 快速开始

```python
from office_suite import OfficeSuite

# 初始化（theme="acks" 默认，产出符合 ACKS 设计规范的文档）
suite = OfficeSuite(theme="acks")

# 3 行代码生成规范化 Word 报告
suite.create("word",
         title="2025年第一季度销售报告",
         content="# 销售情况概述\n\n总销售额 1200 万元，同比增长 25%。\n\n- 销售一部：450 万元\n- 销售二部：380 万元",
         output_path="季度报告.docx")

# 一键转 PDF
suite.convert("季度报告.docx", to="pdf", output_path="季度报告.pdf")
```

---

## 📚 使用指南

### 1. 创建文档

```python
suite = OfficeSuite(theme="acks")   # 默认设计规范主题

# Word（支持 Markdown 标题/列表）
suite.create("word", title="标题", content="...", output_path="a.docx")

# Excel（第一行为表头，可选柱状图）
suite.create("excel", title="数据表", data=[["部门","1月"],["一部",150]], create_chart=True, output_path="b.xlsx")

# PDF
suite.create("pdf", title="标题", content="...", output_path="c.pdf")

# PPT（slides 列表，支持 title/content 布局）
suite.create("pptx", title="演示", slides=[{"title":"封面","layout":"title"},{"title":"内容","content":"• 要点一","layout":"content"}], output_path="d.pptx")
```

### 2. 主题切换与去品牌化

```python
# 单次覆盖主题
suite.create("word", ..., theme="default")   # 简洁样式

# 去品牌化（brand_name="" 去掉 ACKS 品牌 stamp）
suite.create("word", ..., brand_name="")
suite.create("word", ..., brand_name="我的品牌")   # 换品牌
```

### 3. 格式转换 / 批量 / 水印

```python
suite.convert("a.docx", to="pdf")                                    # 单文件转换
suite.batch_convert("word_docs", "pdf_docs", "docx", "pdf")          # 批量转换
suite.add_watermark("c.pdf", "机密文件", output_path="c_wm.pdf")     # PDF 真水印
suite.batch_add_watermark("pdf_docs", "wm_docs", "内部资料")         # 批量水印
```

### 4. 自动化工作流

```python
import yaml
with open("examples/workflow_example.yaml") as f:
    cfg = yaml.safe_load(f)
suite.execute_workflow(cfg)   # 数据提取 → 生成 → 转换 → 水印 → 邮件
```

### 5. 邮件集成

```python
suite.config_email(smtp_server="smtp.qq.com", smtp_port=465, username="you@email.com")
suite.send_email(to=["boss@example.com"], subject="季度报告", attachments=["季度报告.pdf"])
# 密码从环境变量 OFFICE_EMAIL_PASSWORD 读取（见 .env.template）
```

### 6. Excel 数据提取

```python
suite.extract_data("b.xlsx")                       # 默认读取第一个工作表 → {"success": True, "data": [{"部门": "一部", "1月": 150}]}
suite.extract_data("b.xlsx", sheet_name="数据表")   # 指定工作表名，或从 0 开始的索引（如 sheet_name=1）
suite.extract_data("b.xlsx", usecols=["部门"])      # 其余参数透传给 pandas.read_excel
```

> 按单元格原值读取（工号 `00123`、身份证号等文本型数字保持原样），结果可直接 `json.dumps`：空单元格为 `None`，日期/时间为 ISO 8601 字符串。支持 `.xlsx` 与 `.xls`。

---

## 🐳 部署

### 作为 AI Agent 工具（MCP 方式，推荐）

将本库包装为 MCP server，供 Hermes / DSH / Claude Code / Codex 等任意支持 MCP 的 Agent 调用：

```yaml
# DSH cordis.patch.yml 接入示例
- id: mcp-office
  name: '@deepseek-ai/dsh-mcp-client'
  config:
    serverName: office
    transport: streamable-http
    url: https://your-server/mcp
    headers:
      Authorization: 'Bearer ${OFFICE_MCP_TOKEN}'
```

> 工具将呈现为 `mcp__office__create_document`、`mcp__office__convert_document` 等原生工具。

### 本地 / 敏感场景（stdio 方式）

```bash
# 文档不出本机
uvx acks-office-suite-mcp   # 或 python -m office_suite_mcp
```

### Docker 部署（含 LibreOffice + 中文字体）

```dockerfile
FROM python:3.11-slim
RUN apt-get update && apt-get install -y libreoffice fonts-noto-cjk && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . /app
WORKDIR /app
CMD ["python", "-m", "office_suite_mcp"]
```

---

## 🏷️ 技术栈

| 类别 | 技术 |
|------|------|
| 语言 | Python 3.9+ |
| Word | python-docx |
| Excel | openpyxl · pandas · xlrd（.xls） |
| PDF | reportlab · PyPDF2 |
| PPT | python-pptx |
| 图像/字体度量 | Pillow |
| 转档 | LibreOffice（headless） |
| 工作流 | YAML 配置 |

---

## 🤝 适配平台与 Agent

| 形态 | 适配 |
|------|------|
| **Python 库** | 直接 `pip install`，任意 Python 项目可调用 |
| **AI Agent** | Hermes Agent、DeepSeek Harness (DSH)、Claude Code、Codex、任意 MCP 客户端 |
| **操作系统** | Windows / macOS / Linux |
| **协议** | MCP（stdio / streamable-http）、CLI、Python API |

---

## 📁 目录结构

```
src/office_suite/
├── core.py            # OfficeSuite 统一门面（theme/工作流/邮件）
├── docx.py / xlsx.py / pdf.py / pptx.py   # 四格式处理
├── email.py / utils.py                    # 邮件 / 工具
└── design_system/     # ★ ACKS 设计规范 v2.1（tokens.json 单一信源）
```

---

## 🛡️ 安全说明

1. **所有文档处理本地完成**，不上传任何数据到第三方服务器
2. **密码/授权码切勿硬编码**，使用环境变量（见 `.env.template`）
3. 配置文件权限建议 `600`

---

## 🤝 贡献

欢迎提交 Issue 和 PR，有功能建议或 Bug 反馈欢迎提出～

---

## 📄 许可证

MIT License | 详见 [LICENSE](LICENSE)

---

*设计思路受 [MiniMax](https://www.minimaxi.com/) Office Skill 启发，感谢其优秀的产品思路参考。*
