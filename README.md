# ACKS Office Suite Simplify

> 让 AI Agent 和 Python 程序在任何电脑上产出排版规范的 Word / Excel / PPT / PDF——**不需要安装 Office**。

<p align="center">
  <a href="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml"><img alt="tests" src="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml/badge.svg"></a>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white">
  <img alt="Platform" src="https://img.shields.io/badge/平台-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey">
  <a href="LICENSE"><img alt="License" src="https://img.shields.io/badge/License-MIT-green"></a>
</p>

文件全部由 Python 直接写出，用户电脑上有没有 Microsoft Office、WPS 都不影响出品；LibreOffice 只在需要
格式转换时才用到。内置的 ACKS 主题是一套完整的文档设计规范参照，也可以一键去掉品牌、换成你自己的品牌。

由 **ACKS Studio** 出品。包名 `acks-office`，导入名 `acks_office`。

## 能做什么

| 格式 | 生成 | 读取 | 其他 |
|---|---|---|---|
| Word | Markdown 正文 → 规范排版（标题、列表、任务清单、表格、提示块、代码、图片、链接） | 正文与表格文字 | 合并、水印 |
| PDF | 同上，自动嵌入中文字体 | 文字 | 合并、拆分、水印 |
| PPT | 封面页、内容页，品牌与页脚可配置 | 每页文字 | 切换效果 |
| Excel | 表头、斑马纹、数字格式（`default` 主题可附柱状图） | 每行一个对象（保留文本型编号，日期为 ISO 字符串） | — |
| 转换 | docx / xlsx / pptx / pdf 等互转（需要 LibreOffice） | | |

## 三种用法

### 1. 作为 Agent 技能（Skill）

适用于 Claude Code、Codex、OpenClaw、WorkBuddy、Hermes 等支持 [Agent Skills](https://agentskills.io) 的 Agent。

1. 从 [Releases](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases) 下载
   `acks-office-skill-<版本>.zip`（已包含程序代码）；
2. 解压到 Agent 的技能目录，得到 `<技能目录>/acks-office/SKILL.md`。常见位置：

   | Agent | 技能目录（以各 Agent 文档为准） |
   |---|---|
   | Claude Code | `~/.claude/skills/` |
   | Codex | `~/.codex/skills/` |
   | OpenClaw | `~/.openclaw/skills/` 或工作区的 `skills/` |
   | WorkBuddy | `~/.workbuddy/skills/` |
   | Hermes | `~/.hermes/skills/` |

3. 在 Agent 里直接说「帮我写一份季度报告，导出 Word 和 PDF」即可。

第一次使用时，Agent 会运行 `doctor` 检查环境：识别自己所在的 Agent、检查 Python 与依赖、PDF 中文字体、
LibreOffice 及本机其他 Office 程序，给出能力等级和补齐计划。**任何安装都会先征得你的同意**；依赖装在
专用虚拟环境里，不会改动系统 Python。

### 2. 命令行

```bash
pip install "acks-office @ https://github.com/shynloc/ACKS-Office-Suite-Simplify/archive/refs/tags/v2.1.0.zip"

acks-office doctor                                   # 检查环境
acks-office create word -o 报告.docx --title "2026 Q3 经营回顾" --content-file 正文.md
acks-office create pdf  -o 报告.pdf  --title "2026 Q3 经营回顾" --content-file 正文.md
acks-office create pptx -o 汇报.pptx --slides-file slides.json --brand-name "栖木咖啡"
acks-office create excel -o 数据.xlsx --data-file data.csv
acks-office extract 报告.docx
acks-office convert 报告.docx --to pdf                # 需要 LibreOffice
acks-office watermark 报告.pdf --text "内部资料"
acks-office merge a.pdf b.pdf -o 合并.pdf
acks-office fonts install noto-sans-sc               # 下载开源中文字体（PDF 嵌入用）
```

每个命令都可以加 `--json`，输出统一结构 `{"ok", "data", "artifacts", "warnings", "error"}`，
退出码 0 成功、1 失败、2 参数错误；已有文件默认不覆盖（加 `--overwrite`）。完整说明见
[命令参考](skills/acks-office/references/commands.md)。

### 3. Python 库

```python
from acks_office import OfficeSuite

suite = OfficeSuite()  # 默认 theme="acks"

suite.create("word", title="2026 Q3 经营回顾", subtitle="营收、门店与会员",
             content="# 经营概览\n\n营收 **1.28 亿元**，同比增长 24%。\n\n"
                     "| 区域 | 营收（万元） |\n|---|--:|\n| 华东 | 5,888 |",
             output_path="报告.docx", brand_name="")          # brand_name="" 去掉品牌

suite.create("pdf", title="2026 Q3 经营回顾", content="…", output_path="报告.pdf")
suite.create("pptx", title="经营回顾", output_path="汇报.pptx", brand_name="栖木咖啡",
             slides=[{"title": "经营回顾", "subtitle": "2026 Q3", "layout": "title"},
                     {"title": "核心结论", "content": "营收增长 24%\n会员 12 万"}])
suite.create("excel", title="销售", data=[["部门", "1月"], ["一部", 150]], output_path="数据.xlsx")

suite.extract_data("数据.xlsx")       # {"success": True, "data": [{"部门": "一部", "1月": 150}]}
suite.convert("报告.docx", to="pdf")  # 需要 LibreOffice
```

所有方法返回字典：成功时 `success` 为 True，并带输出路径等信息；失败时为 False 并附 `error`。

## 主题与品牌

| 需求 | 做法 |
|---|---|
| 设计规范排版（默认） | `theme="acks"`：配色、字体、版式全部来自 `design_system/tokens.json` |
| 朴素排版、便于他人二次编辑 | `theme="default"`：Word 内置样式 |
| 去掉品牌 | `brand_name=""`：封面、页眉、页脚不出现任何品牌字样 |
| 换成你的品牌 | `brand_name="品牌名"`，页脚文字可用 `footer_label` 指定 |

3.0 将提供主题引擎：用一份主题文件描述你自己的设计规范，附带偏商务（Slate）与偏杂志（Folio）两套示例主题，
并能由 Agent 引导你逐项创建。

## 字体

- **PDF**：自动选择可嵌入的中文字体——参数 `font` → 环境变量 `ACKS_OFFICE_PDF_FONT` →
  `fonts install` 下载的 Noto Sans SC → 系统自带（macOS 华文黑体、Windows 微软雅黑、Linux 文泉驿等）。
  都没有时回退到不嵌入的 STSong-Light，并在结果里给出 `FONT_FALLBACK` 提醒。
- **Word / PPT**：使用主题声明的字体名，打开文档的电脑缺字体时会被替换显示。`doctor` 会列出缺少的主题字体
  及下载地址（均为可免费商用的开源字体）。

## LibreOffice（可选）

只有 `convert`（以及基于它的批量转换）需要。查找顺序：环境变量 `ACKS_OFFICE_SOFFICE` → PATH →
各系统标准安装位置（如 macOS 的 `/Applications/LibreOffice.app`）。安装：

```bash
brew install --cask libreoffice                                # macOS
winget install -e --id TheDocumentFoundation.LibreOffice       # Windows
sudo apt-get install -y libreoffice                            # Ubuntu / Debian
```

## 从 2.0 升级

- 导入名改为 `acks_office`；`import office_suite` 仍可用（会提示弃用），3.0 移除。
- 段落内的单个换行保留为换行；封面标题不再重复；页脚、合并、PDF 字体等行为有调整，
  详见 [CHANGELOG](CHANGELOG.md)。

## 目录结构

```
src/acks_office/
├── core.py              OfficeSuite：统一入口、批量处理、工作流、邮件
├── cli.py · doctor.py   命令行与环境检查
├── docx.py · pdf.py · pptx.py · xlsx.py
├── markdown_blocks.py   Markdown 解析（Word 与 PDF 共用）
├── fonts.py · utils.py  字体选择与下载、LibreOffice 查找与转换
└── design_system/       ACKS 设计规范（tokens.json 为唯一信源，来源见 SOURCE.md）
skills/acks-office/      Agent 技能：SKILL.md、入口脚本、参考文档
tools/build_skill_bundle.py   生成技能压缩包
```

## 开发

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
python -m pip install -e ".[dev]"
python -m pytest                    # 单元测试
python test_integration.py          # 集成测试
python tools/build_skill_bundle.py  # 生成 dist/acks-office-skill-<版本>.zip
```

## 安全

文档全部在本机处理，不收集任何使用数据。联网只发生在：`doctor` 的连通性检查（可用 `--no-network` 关闭）、
`fonts install` 下载字体（固定版本并校验 SHA-256）、`send_email` 连接你的邮件服务器（校验证书、
拒绝明文登录），以及工作流里你自己配置的 API 数据源。
详见 [SECURITY.md](SECURITY.md)。

## 许可证

[MIT](LICENSE)

---

*设计思路受 [MiniMax](https://www.minimaxi.com/) Office Skill 启发，感谢其优秀的产品思路参考。*
