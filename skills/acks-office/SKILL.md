---
name: acks-office
description: >-
  Create, read and convert Word (.docx), Excel (.xlsx), PowerPoint (.pptx) and PDF files with
  built-in or custom design themes, without needing Microsoft Office or WPS installed (works on
  macOS, Windows and Linux). Use when the user asks to write or export a report, proposal, memo,
  slide deck, spreadsheet or PDF; to pull text or tables out of office files; to convert between
  formats; to watermark or merge documents; or to set up a brand theme. 生成、读取、转换 Word / Excel /
  PPT / PDF，按内置或自定义主题排版，不依赖本机 Office；用于写报告、方案、汇报 PPT、表格、PDF，
  提取文档内容，格式转换，加水印，合并文档，定制品牌主题。
license: MIT
compatibility: >-
  Needs Python 3.10+. Python packages are installed on first use into a private virtual environment,
  only after the user agrees. LibreOffice is optional (format conversion only).
metadata:
  version: "3.0.0"
  author: "ACKS Studio"
  homepage: "https://github.com/shynloc/ACKS-Office-Suite-Simplify"
---

# acks-office

用一套命令按主题生成 Word / Excel / PPT / PDF，也能读取、转换、加水印、合并，还能为用户定制主题。
所有文件都由 Python 直接写出，**不需要用户装 Office**。

入口脚本：`scripts/acks.py`（下文的 `ACKS` 指 `python3 <本技能目录>/scripts/acks.py`；
Windows 上把 `python3` 换成 `python` 或 `py -3`）。命令一律加 `--json`，按 JSON 结果行事。

## 第一次使用：先检查环境

在本会话第一次生成文档之前运行一次：

```bash
python3 <本技能目录>/scripts/acks.py doctor --json
```

- 找不到 Python：告诉用户需要 Python 3.10 以上，给出 https://www.python.org/downloads/ ，
  问用户是否要安装；**不要自行安装**。
- 结果里 `error.code` 为 `PACKAGE_MISSING`：说明要安装程序本体，见下面「补齐」。
- 正常时看 `data.level`：
  - `none`：缺依赖，还不能生成文档，先补齐；
  - `L0`：能生成全部格式，但缺字体（PDF 中文可能不嵌入，或默认主题的字体会被替换）；
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
| Word | `ACKS create word -o 报告.docx --content-file 正文.md --theme slate --json` |
| PDF | `ACKS create pdf -o 报告.pdf --content-file 正文.md --theme slate --json` |
| PPT | `ACKS create pptx -o 汇报.pptx --slides-file slides.json --theme slate --json` |
| Excel | `ACKS create excel -o 数据.xlsx --data-file data.json --theme slate --json` |
| 读取 | `ACKS extract 文件 --json`（Excel 加 `--sheet 名称或序号`） |
| 转换 | `ACKS convert 文件 --to pdf --json`（需要 LibreOffice） |
| 水印 | `ACKS watermark 文件.pdf --text "内部资料" --json` |
| 合并 | `ACKS merge a.pdf b.pdf -o 合并.pdf --json`（全部 PDF 或全部 docx） |
| 主题 | `ACKS theme list --json` / `ACKS theme preview slate --json` / `ACKS theme init 新主题 --extends slate --json` |
| 字体 | `ACKS fonts list --json` / `ACKS fonts install --theme slate --system --json`（需用户同意） |

- 正文写成 Markdown 文件再用 `--content-file`：支持标题、粗体、斜体、链接、列表、任务清单、
  表格、引用、`> [!NOTE]` 提示块、代码块、图片（相对正文文件所在目录）。详见
  [references/markdown.md](references/markdown.md)。
- 正文开头可写 front matter（`title`、`subtitle`、`kicker`、`author`、`date`、`version` 等），
  封面和页眉页脚会用到；也可用 `--meta 键=值`。
- PPT 的 `slides.json`：`[{"layout": "title", "title": "封面", "subtitle": "副标题"},
  {"title": "要点", "bullets": ["第一条", "第二条"]}]`；另有章节页、数据页（关键数字 + 图表）、表格页、
  大数字页、引文页、图片页。Excel 可以是二维数组、对象数组或多张工作表。格式见 commands.md。
- 主题：默认 `neutral`（只用系统字体）；`slate` 适合商务报告与汇报，`folio` 适合品牌手册、年报这类
  杂志感的文档。用户没指定时按用途选，拿不准就用默认。品牌用 `--brand-name`（`""` 去掉品牌）。

完整参数与错误码见 [references/commands.md](references/commands.md)。

## 处理结果

输出统一为 `{"ok", "data", "artifacts", "warnings", "error"}`，退出码 0 成功、1 失败、2 参数错误。

- 成功：把 `artifacts[].path` 告诉用户；`warnings` 里的内容要转告：`FONT_SUBSTITUTED` 表示主题字体
  没装、用了本机的备选字体（`install` 是可下载的字体键，征得同意后可 `fonts install`）；
  `FONT_NOT_EMBEDDED` 表示 PDF 里的中文没能嵌入，换设备显示可能不同。
- 失败：看 `error.code` 与 `error.hint`。`OUTPUT_EXISTS` 时先问用户是否覆盖，同意后再加
  `--overwrite`；`ENGINE_UNAVAILABLE` 表示需要 LibreOffice，按「补齐」处理或改为直接生成目标格式。
- 不要覆盖用户已有文件，不要把文件写到用户没有指定的位置以外（未指定时放在当前工作目录）。

## 主题与定制

主题决定配色、字体、字号、版式与页眉页脚，Word、PDF、PPT、Excel 共用。用户想要自己的风格时：

1. 问清品牌名、主色、字体偏好和使用场景；
2. `theme init 名称 --extends 最接近的内置主题 --accent #色值 --brand 品牌` 新建主题，按
   [references/design-system.md](references/design-system.md) 修改；
3. `theme validate` 检查，`theme preview` 生成 HTML 样张给用户确认（HTML 可直接展示）；
4. 确认后用 `--theme 名称` 出品。

## 参考

- [references/commands.md](references/commands.md)：命令、参数、JSON 结构、错误码
- [references/doctor.md](references/doctor.md)：环境报告各字段、能力等级、补齐计划
- [references/markdown.md](references/markdown.md)：Markdown 在 Word / PDF 中的呈现
- [references/design-system.md](references/design-system.md)：主题、主题包格式与定制方法
