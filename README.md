# ACKS Office Suite Simplify

[中文](#中文) · [English](#english)

> 面向 AI Agent 和 Python 程序的办公文档引擎：按主题生成 Word / Excel / PPT / PDF，读取、转换、加水印、合并，
> 提供命令行、环境检查和技能包。
> A document engine for AI agents and Python: themed Word / Excel / PowerPoint / PDF generation,
> plus reading, conversion, watermarking and merging, with a CLI, an environment check and an agent skill.

文档对应版本：**3.0.0**。GitHub 与 PyPI 使用本文件作为项目说明；历史版本的说明随发行包保留。

<p align="center">
  <a href="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml"><img alt="tests" src="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml/badge.svg?branch=main"></a>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white">
  <img alt="Platform" src="https://img.shields.io/badge/平台-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey">
  <a href="https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/LICENSE"><img alt="License" src="https://img.shields.io/badge/License-MIT-green"></a>
</p>

## 中文

由 **ACKS Studio** 出品。安装包名 `acks-office`，导入名 `acks_office`。
生成与读取文件由 Python 库完成，不需要 Microsoft Office 或 WPS；格式转换需要 LibreOffice。
从 2.x 升级请先看[迁移指南](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/docs/migration-3.0.md)。

### 当前能力

| 格式 / 入口 | 已实现 | 范围与条件 |
|---|---|---|
| 主题 | 内置 `neutral`（默认）、`slate`（商务）、`folio`（杂志）；自建主题、校验、HTML 样张 | 主题决定配色、字体、字号、版式与页眉页脚，四种格式共用 |
| Word `.docx` | 封面（标题块 / 独立封面 / 期刊封面）、编号标题或章节页、分栏与首字下沉、表格、提示块、引文、代码、图片；提取正文与表格；合并、水印 | 写入字体名称，缺字体时用本机备选字体并提醒 `FONT_SUBSTITUTED` |
| PDF | 引擎直接排版，版式与 Word 一致，字体嵌入文件；提取文字、合并、水印；Python API 按页拆分 | 只能嵌入 TrueType 字体；都找不到时中文由阅读器替换显示（`FONT_NOT_EMBEDDED`）；扫描件需外部 OCR |
| PPT `.pptx` | 封面、章节页、内容页、数据页（关键数字 + 条形图）、表格页、大数字页、引文页、图片页；文字提取；Python API 添加切换效果 | 标题按实际字宽自动缩放；版式随主题变化 |
| Excel | 标题行、表头与合计行样式、数字格式与负数色、多张工作表与图表；读取 `.xlsx` / `.xls` | 公式由打开文件的软件计算（自动生成的合计行除外） |
| `convert` | 委托 LibreOffice 转换文件，如 DOCX → PDF | 转换方向取决于引擎的导入 / 导出过滤器 |
| `doctor` | 只读检查 Python、依赖、字体（按主题）、LibreOffice、宿主线索，给出补齐计划 | 不执行安装；其他 Office 程序只检测、不调用 |
| 扩展 | SMTP 邮件、批量操作、顺序执行工作流配置 | 在已弃用的 `OfficeSuite` 上，4.0 前保留 |

直接读取支持 `.docx`、`.pdf`、`.pptx`、`.xlsx`、`.xls`。旧版 `.doc` / `.ppt` 可先交给 LibreOffice 转换。

### 安装与升级

推荐在专用虚拟环境中从 [PyPI](https://pypi.org/project/acks-office/) 安装（Python 3.10 或更高）：

```bash
python -m pip install acks-office==3.0.0
acks-office doctor --json
```

升级到最新发行版：`python -m pip install --upgrade acks-office`。下载主题用到的开源字体需要字体可选依赖：

```bash
python -m pip install "acks-office[fonts]==3.0.0"
acks-office fonts install --theme slate --system
```

各系统虚拟环境、依赖、字体和 LibreOffice 的准备步骤见
[安装指南](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/INSTALL_GUIDE.md)。
也可从 [GitHub Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v3.0.0)
下载 wheel 或技能包。

### Agent 技能

技能包由 `SKILL.md`、Python 入口脚本和程序代码组成，供支持技能规范及本地命令执行的 Agent 使用。

1. 从 [GitHub Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v3.0.0)
   下载 `acks-office-skill-3.0.0.zip`，解压到宿主技能目录，得到 `acks-office/SKILL.md`。
   Claude Code、Codex、OpenClaw、WorkBuddy、Hermes 的目录示例见安装指南。
2. Agent 运行 `scripts/acks.py doctor --json`，说明补齐计划并取得用户同意后，安装依赖到专用虚拟环境。

技能包含程序代码；依赖与字体的安装由 Agent 或用户在征得同意后执行，脚本本身不会自动安装。

### 命令行

子命令：`create`、`extract`、`convert`、`watermark`、`merge`、`theme`、`fonts`、`doctor`、`version`。
输入格式（front matter、幻灯片、Excel 数据）见
[命令参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/commands.md)。

```bash
acks-office create word -o 报告.docx --content-file 正文.md --theme slate
acks-office create pdf -o 报告.pdf --content-file 正文.md --theme slate
acks-office create pptx -o 汇报.pptx --slides-file slides.json --theme folio
acks-office create excel -o 数据.xlsx --data-file data.json
acks-office theme preview folio
acks-office theme init my-brand --extends slate --accent "#0B6E4F" --brand "栖木咖啡"
acks-office extract 报告.docx
acks-office convert 报告.docx --to pdf -o 转换版.pdf
acks-office watermark 报告.pdf --text "内部资料"
acks-office merge a.pdf b.pdf -o 合并.pdf
```

每个子命令可加 `--json`，返回 `{ok, data, artifacts, warnings, error}`；退出码 0 成功、1 执行失败、2 参数错误。
默认不覆盖已有输出，需显式加 `--overwrite`。命令不在 PATH 上时，可用 `python -m acks_office`。

### Python 库

```python
import acks_office

acks_office.create(
    "word", "报告.docx", theme="slate",
    content="---\ntitle: 经营回顾\nauthor: 经营分析部\n---\n\n# 概览\n\n营收 **增长 24%**。\n\n"
            "| 区域 | 营收 |\n|---|--:|\n| 华东 | 5,888 |",
)
rows = acks_office.extract("数据.xlsx")
acks_office.add_watermark("报告.pdf", "内部资料")  # 写到 报告_watermarked.pdf，不改原文件
```

`create`、`extract`、`convert`、`add_watermark`、`merge` 出错时抛出异常（`FileNotFoundError`、`ValueError`、
`acks_office.themes.ThemeError` 等）。各格式的完整参数见 `acks_office.docx.create_word`、`xlsx.create_excel`、
`pdf.create_pdf`、`pptx.create_pptx` 的说明。PDF 拆分与 PPT 切换通过格式模块调用：

```python
from acks_office.pdf import split_pdf
from acks_office.pptx import add_transition_effects

split_pdf("报告.pdf", "拆分页")
add_transition_effects("汇报.pptx", "汇报_淡入.pptx", effect="fade", duration=0.5)
```

### 主题与字体

| 主题 | 定位 | 字体 |
|---|---|---|
| `neutral`（默认） | 零依赖，任何环境都能出品 | 系统字体（苹方 / 微软雅黑、Helvetica Neue / Segoe UI） |
| `slate` | 报告、方案、汇报：清晰板正 | 思源黑体 + Source Sans 3 |
| `folio` | 品牌手册、年报、作品集：杂志感 | 思源宋体 + Source Serif 4、Fraunces、霞鹜文楷 |

- 主题是一个数据包（`theme.json`、`tokens.json`、`chrome.json`、`fonts.json`），沿 `extends` 继承；
  `theme init` 新建，`theme validate` 检查对比度与字体，`theme preview` 生成 HTML 样张和四种格式的示例文件。
  格式与字段见[设计参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/design-system.md)。
- 主题字体没装时，生成文件改用本机备选字体；`fonts install --theme <主题>` 在用户同意后下载主题缺的开源字体
  （地址固定、校验哈希），`--system` 同时装到当前用户的字体目录。
- PDF 的中文字体依次取 `--font` / `ACKS_OFFICE_PDF_FONT`、主题字体及其备选、同风格的系统 TrueType 字体，
  都没有时用阅读器自带的 STSong-Light（不嵌入）。能否嵌入以 doctor 检测为准。
- `theme="acks"` 暂时保留 2.x 的 ACKS 样式，只支持基础版式；`theme="default"` 是 `neutral` 的别名。

正文写法见
[Markdown 参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/markdown.md)，
环境等级与字段见
[doctor 参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/doctor.md)。

### 兼容与工作流

`OfficeSuite` 与旧导入名 `office_suite` 保留到 4.0，使用时提示弃用；新代码使用上面的函数。
邮件、批量与 `OfficeSuite.execute_workflow(config)` 目前仍在 `OfficeSuite` 上：工作流按顺序执行步骤，
调度由调用方负责；加载 YAML 或使用 API 数据源时安装 `acks-office[workflow]`。示例见
[workflow_example.yaml](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/examples/workflow_example.yaml)。
当前不执行 `schedule`、`on_success` 等根配置，`notification` 只打印通知。
程序提供技能与 CLI，不包含内置 MCP 服务。

### 验证与安全

最低 Python 3.10；CI 覆盖 Ubuntu / macOS / Windows × Python 3.10 / 3.13。
无 LibreOffice 的 runner 跳过实际转换测试，生成与读取测试仍运行。

文档在本机处理。doctor 联网检查可用 `--no-network` 关闭；字体下载固定版本并校验哈希；
SMTP 默认验证证书并拒绝明文登录。版本变化见
[CHANGELOG](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/CHANGELOG.md)，
安全与环境变量见
[SECURITY.md](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/SECURITY.md)。

### 开发与文档维护

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
python -m pip install -e ".[dev]"
python tools/check_release_docs.py
python -m pytest
python test_integration.py
python tools/build_skill_bundle.py
```

README 是 GitHub 与 PyPI 的共同说明来源。功能新增、移除或弃用时同步更新文档，维护流程见
[CONTRIBUTING.md](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/CONTRIBUTING.md)。

## English

Made by **ACKS Studio**. Install `acks-office`, import `acks_office`. Files are written and read by Python
libraries, so Microsoft Office or WPS is not needed; format conversion uses LibreOffice. Upgrading from 2.x?
Read the [migration guide](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/docs/migration-3.0.md).

### What it does

- **Themes**: built-in `neutral` (default), `slate` (business) and `folio` (magazine); create, validate and
  preview your own. A theme sets colors, fonts, sizes, layouts and headers / footers for all four formats.
- **Word**: covers, numbered headings or chapter openers, columns and drop caps, tables, callouts, quotes,
  code and images from Markdown; text and table extraction; merge and watermark.
- **PDF**: laid out directly by the engine with the same design as Word, fonts embedded; text extraction,
  merge, watermark, page split.
- **PowerPoint**: cover, section, content, data (KPIs + bar chart), table, big number, quote and image slides.
- **Excel**: title row, themed header and total rows, number formats, negative colors, several sheets and charts;
  reads `.xlsx` / `.xls`.
- **CLI and agent skill** with JSON output, `doctor` environment checks and consent-based setup plans.

### Install

Python 3.10 or newer:

```bash
python -m pip install acks-office==3.0.0
acks-office doctor --json
```

Theme fonts are open-source and downloaded only when you ask:
`python -m pip install "acks-office[fonts]==3.0.0"` then `acks-office fonts install --theme slate --system`.
The agent skill is `acks-office-skill-3.0.0.zip` on the
[GitHub Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v3.0.0).

### Use

```bash
acks-office create word -o report.docx --content-file report.md --theme slate
acks-office create pptx -o deck.pptx --slides-file slides.json --theme folio
acks-office theme preview folio
acks-office theme init my-brand --extends slate --accent "#0B6E4F" --brand "My Brand"
```

```python
import acks_office

acks_office.create("pdf", "report.pdf", theme="slate", title="Quarterly Review",
                   content="# Overview\n\nRevenue grew **24%**.")
rows = acks_office.extract("data.xlsx")
```

Functions raise exceptions on failure. `OfficeSuite` and the old `office_suite` import name remain until 4.0 with a
deprecation warning. Commands, input formats and error codes are in the
[command reference](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/commands.md);
theme packages are described in the
[design reference](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/skills/acks-office/references/design-system.md)
(both in Chinese).

[MIT License](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/LICENSE).
设计思路受 [MiniMax Office Skill](https://www.minimaxi.com/) 启发。
