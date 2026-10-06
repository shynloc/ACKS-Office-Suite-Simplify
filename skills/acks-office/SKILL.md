---
name: acks-office
description: >-
  Create, read and convert Word (.docx), Excel (.xlsx), PowerPoint (.pptx) and PDF files with a
  consistent document design system, without needing Microsoft Office or WPS installed (works on
  macOS, Windows and Linux). Use when the user asks to write or export a report, proposal, memo,
  slide deck, spreadsheet or PDF; to pull text or tables out of office files; to convert between
  formats; or to watermark or merge documents. 生成、读取、转换 Word / Excel / PPT / PDF，按文档设计规范排版，
  不依赖本机 Office；用于写报告、方案、汇报 PPT、表格、PDF，提取文档内容，格式转换，加水印，合并文档。
license: MIT
compatibility: >-
  Needs Python 3.9+. Python packages are installed on first use into a private virtual environment,
  only after the user agrees. LibreOffice is optional (format conversion only).
metadata:
  version: "2.1.0"
  author: "ACKS Studio"
  homepage: "https://github.com/shynloc/ACKS-Office-Suite-Simplify"
---

# acks-office

用一套命令生成符合设计规范的 Word / Excel / PPT / PDF，也能读取、转换、加水印、合并。
所有文件都由 Python 直接写出，**不需要用户装 Office**。

入口脚本：`scripts/acks.py`（下文的 `ACKS` 指 `python3 <本技能目录>/scripts/acks.py`；
Windows 上把 `python3` 换成 `python` 或 `py -3`）。命令一律加 `--json`，按 JSON 结果行事。

## 第一次使用：先检查环境

在本会话第一次生成文档之前运行一次：

```bash
python3 <本技能目录>/scripts/acks.py doctor --json
```

- 找不到 Python：告诉用户需要 Python 3.9 以上，给出 https://www.python.org/downloads/ ，
  问用户是否要安装；**不要自行安装**。
- 结果里 `error.code` 为 `PACKAGE_MISSING`：说明要安装程序本体，见下面「补齐」。
- 正常时看 `data.level`：
  - `none`：缺依赖，还不能生成文档，先补齐；
  - `L0`：能生成全部格式，但缺字体（PDF 中文可能不嵌入，或 Word/PPT 主题字体会被替换）；
  - `L1`：字体齐全，排版与设计一致；
  - `L2`：另有 LibreOffice，可做格式转换。
- 只为当前任务需要的项补齐。例如用户只要 Word，就不必提 LibreOffice。

### 补齐（必须先征得同意）

`data.plan`（或 `PACKAGE_MISSING` 结果里的 `data.plan`）每一项都有 `why`、`steps`，并标了
`needs_consent: true`。做法：

1. 用一两句话向用户说明要装什么、装到哪里（`target`）、为什么（`why`），**等用户明确同意**；
2. 同意后逐条执行 `steps`：每一步是参数列表，按原样执行，不要改写成别的安装方式；
   某步失败时，若有 `alternative_steps` 可改用它，否则把错误告诉用户；
3. 完成后重新运行 `doctor --json` 确认。

依赖默认装进专用虚拟环境（`target` 所示目录），不改动系统 Python；入口脚本之后会自动使用它。
`needs_admin: true` 的步骤需要管理员权限，交给用户自己执行。`links` 是需要用户手动下载的地址。
网络检查 `data.network.github` 为 false 时，字体下载可能失败，可请用户提供镜像地址并设置环境变量
`ACKS_OFFICE_DOWNLOAD_MIRROR`。

## 常用命令

| 任务 | 命令 |
|---|---|
| Word | `ACKS create word -o 报告.docx --title "标题" --content-file 正文.md --json` |
| PDF | `ACKS create pdf -o 报告.pdf --title "标题" --content-file 正文.md --json` |
| PPT | `ACKS create pptx -o 汇报.pptx --slides-file slides.json --json` |
| Excel | `ACKS create excel -o 数据.xlsx --data-file data.csv --json` |
| 读取 | `ACKS extract 文件 --json`（Excel 加 `--sheet 名称或序号`） |
| 转换 | `ACKS convert 文件 --to pdf --json`（需要 LibreOffice） |
| 水印 | `ACKS watermark 文件.pdf --text "内部资料" --json` |
| 合并 | `ACKS merge a.pdf b.pdf -o 合并.pdf --json`（全部 PDF 或全部 docx） |
| 字体 | `ACKS fonts list --json` / `ACKS fonts install noto-sans-sc --json` |

- 正文写成 Markdown 文件再用 `--content-file`：支持标题、粗体、斜体、链接、列表、任务清单、
  表格、引用、`> [!NOTE]` 提示块、代码块、图片（相对正文文件所在目录）。详见
  [references/markdown.md](references/markdown.md)。
- PPT 的 `slides.json`：`[{"title": "封面", "subtitle": "副标题", "layout": "title"},
  {"title": "要点", "content": "第一行\n第二行"}]`。
- 品牌：默认是 ACKS 参考主题。用户没有品牌时加 `--brand-name ""`（去掉品牌字样），有品牌时
  `--brand-name "品牌名"`；页脚文字用 `--footer-label`。
- 主题：`--theme acks`（默认，设计规范排版）或 `--theme default`（朴素样式）。

完整参数与错误码见 [references/commands.md](references/commands.md)。

## 处理结果

输出统一为 `{"ok", "data", "artifacts", "warnings", "error"}`，退出码 0 成功、1 失败、2 参数错误。

- 成功：把 `artifacts[].path` 告诉用户；`warnings` 里的内容要转告（例如 `FONT_FALLBACK`
  表示 PDF 中文字体没有嵌入，可建议安装 noto-sans-sc）。
- 失败：看 `error.code` 与 `error.hint`。`OUTPUT_EXISTS` 时先问用户是否覆盖，同意后再加
  `--overwrite`；`ENGINE_UNAVAILABLE` 表示需要 LibreOffice，按「补齐」处理或改为直接生成目标格式。
- 不要覆盖用户已有文件，不要把文件写到用户没有指定的位置以外（未指定时放在当前工作目录）。

## 设计规范

内置的 ACKS 主题是一套参照规范：配色、字体、版式都有明确的令牌（tokens）。用户想要自己的风格时，
先用 `--brand-name`、`--footer-label` 做轻量定制；完整的「自定义主题」将在 3.0 提供，
思路见 [references/design-system.md](references/design-system.md)。

## 参考

- [references/commands.md](references/commands.md)：命令、参数、JSON 结构、错误码
- [references/doctor.md](references/doctor.md)：环境报告各字段、能力等级、补齐计划
- [references/markdown.md](references/markdown.md)：Markdown 在 Word / PDF 中的呈现
- [references/design-system.md](references/design-system.md)：主题、品牌与自定义设计规范
