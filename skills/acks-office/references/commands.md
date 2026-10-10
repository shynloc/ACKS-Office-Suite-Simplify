# 命令参考

入口：`python3 <技能目录>/scripts/acks.py`（已用 pip 安装时也可用 `acks-office` 或
`python -m acks_office`，参数相同）。`--json` 放在子命令前后都可以。

## 输出格式

```json
{
  "ok": true,
  "data": {},
  "artifacts": [{"path": "/abs/path/报告.docx", "size": 38211}],
  "warnings": [{"code": "FONT_SUBSTITUTED", "message": "…"}],
  "error": null
}
```

失败时 `ok` 为 false，`error` 为 `{"code", "message", "hint"}`。退出码：0 成功，1 执行失败，2 参数错误。

## create：生成文档

```
create {word|docx|pdf|excel|xlsx|pptx|ppt} -o 输出路径 [选项]
```

输出路径的扩展名需与类型一致：Word 为 `.docx`、PDF 为 `.pdf`、Excel 为 `.xlsx`、PPT 为 `.pptx`。

| 选项 | 适用 | 说明 |
|---|---|---|
| `--theme` | 全部 | 主题名称或主题目录，默认 `neutral`；另有 `slate`（商务）、`folio`（杂志）与已安装的主题，见 [design-system.md](design-system.md) |
| `--title` | 全部 | 标题；不给时依次用正文 front matter、数据文件 `meta` 里的 title，再用输出文件名（Excel 没给标题时不加标题行） |
| `--content` / `--content-file` | Word、PDF | 正文 Markdown；`--content-file -` 从标准输入读取。写法见 [markdown.md](markdown.md) |
| `--data-file` | Excel | JSON 或 CSV，格式见下文 |
| `--slides-file` | PPT | JSON，格式见下文 |
| `--meta 键=值` | 全部 | 文档元数据，可重复：`subtitle`、`kicker`、`author`、`date`、`version`、`brand`、`classification`、`publication`、`issue`、`season`、`lede`、`short_title`；值里的 `\n` 表示换行 |
| `--brand-name` | Word、PDF、PPT | 品牌名（用在封面和页眉页脚）；`""` 去掉品牌 |
| `--footer-label` | Word、PDF、PPT | 页脚左侧文字，默认按主题模板生成 |
| `--subtitle` | Word、PDF、PPT | 副标题（等同 `--meta subtitle=…`） |
| `--font-policy` | Word、PPT、Excel | `local`（默认）：缺主题字体时写入本机的备选字体；`theme`：总写主题字体名，由打开文件的软件替换 |
| `--font` | PDF | 中文 TrueType 字体文件（.ttf/.ttc），代替主题的中文字体 |
| `--chart` | Excel | 附一张图表（分类取第一列，数值取第一个数字列） |
| `--overwrite` | 全部 | 允许覆盖已有文件 |

`data` 含 `output_path`、`file_size`、`theme`，以及：PDF 的 `pages`（实际页数）与 `fonts`（嵌入的字体），
PPT 的 `slides_count`，Excel 的 `sheets`（每张工作表的名称、行数、列数）。主题字体没装时有
`FONT_SUBSTITUTED` 提醒（附 `install` 字体键，可用 `fonts install` 下载）；PDF 里中文没能嵌入时有
`FONT_NOT_EMBEDDED`。

### Word / PDF 的 front matter

正文开头可写元数据，和 `--meta` 一样（命令行参数优先）：

```markdown
---
title: 2026 年第三季度\n经营回顾
kicker: 经营分析报告
subtitle: 营收、门店与会员
author: 经营分析部
date: 2026 年 10 月 8 日
version: v1.0
---
```

封面随主题变化：neutral 把标题放在第一页顶部；slate 单独一页封面，日期、编制、版本固定在页底；
folio 是期刊式封面（`publication`、`issue`、`season`、`lede`），并在页底列出本期目录。
`cover: false` 不单独做封面。

### Excel 数据

`--data-file` 接受四种写法：

- JSON 二维数组：第一行为表头，`[["区域", "营收"], ["华东", 5888]]`；
- JSON 对象数组：每行一个对象，键为表头（和 `extract` 的输出相同）；
- CSV：第一行为表头，数字自动识别；
- 多张工作表：`{"meta": {"title": "…"}, "sheets": [工作表, …]}`，每张工作表：

| 字段 | 说明 |
|---|---|
| `name` | 工作表名（去掉 `[]:*?/\`，最长 31 字；重名自动加序号） |
| `title` | 第 1 行的标题；不写则表头在第 1 行 |
| `header` / `rows` | 表头与数据行；以 `=` 开头的字符串写成公式 |
| `formats` | 数字格式，键为列序号（从 0 起）或表头文字，如 `{"同比": "+0.0%;-0.0%"}`；不写时按数据自动选（千位分隔，负数加括号） |
| `total` | 最后一行以「合计 / 总计 / 小计 / Total」开头时自动按合计行处理；`true` 时追加 SUM 合计行；`false` 时不处理 |
| `highlight` | 重点单元格加框：`[[数据行, 列], …]`，都从 0 数起 |
| `note` | 表格下方的说明文字 |
| `chart` | 图表：`{"type": "column" / "bar" / "line", "values": [列序号…], "title": "…"}`，`{}` 用默认值 |

年份、编号、序号、代码等列不加千位分隔。负数用主题的负数色；表头与合计行的样式随主题变化。

### PPT 幻灯片

`--slides-file` 是数组，或 `{"meta": {"title": …, "brand": …}, "slides": [...]}`。每页按 `layout` 取字段：

| layout | 字段 |
|---|---|
| `title`（封面） | `kicker`、`title`、`subtitle`、`footer`、`kpis`（`[{label, value, unit, delta, highlight}]`） |
| `section`（章节页） | `number`、`kicker`、`title`、`title_en`、`subtitle` |
| `content`（默认） | `title`、`subtitle`、`kicker`、`label`，正文用 `content`（`\n` 分行）或 `bullets`（数组） |
| `data` | `title`、`subtitle`、`kpis`、`chart`（`{title, unit, data: [[名称, 数值]…], highlight: 序号, note, format}`） |
| `table` | `title`、`subtitle`、`table`（`{header, rows, widths}`） |
| `number`（大数字） | `label`、`value`、`unit`、`delta_label`、`delta`、`text`，可附 `chart` |
| `quote` | `quote`、`author`、`role`、`source` |
| `image` | `image`（相对幻灯片文件所在目录）、`caption`、`title` |

任何一页都可以加 `notes`（演讲者备注）。不认识的 `layout` 按 `content` 处理并给出 `UNKNOWN_LAYOUT` 提醒。
封面、章节页、内容页的具体版式随主题变化（例如 folio 的章节页为深色底、内容页为左窄右宽两栏）。

## extract：读取内容

```
extract 文件 [--sheet 名称或从0开始的序号]
```

- `.docx`、`.pdf`：`data.text`（Word 表格按行输出，单元格用 ` | ` 分隔）
- `.xlsx`、`.xls`：`data.rows`，每行一个对象（表头为键；空单元格为 null；日期为 ISO 字符串）。
  第 1 行只有一个单元格有内容、第 2 行至少两个时，视第 1 行为标题，表头从第 2 行读
- `.pptx`：`data.slides`，`{"1": "第一页文字", …}`

## convert：格式转换（需要 LibreOffice）

```
convert 文件 --to pdf [-o 输出路径] [--overwrite]
```

`--to` 可为 pdf、docx、xlsx、pptx、odt 等。默认输出与原文件同名、换扩展名。
LibreOffice 的查找顺序：环境变量 `ACKS_OFFICE_SOFFICE` → PATH 中的 soffice → 各系统标准安装位置。
实际转换方向取决于 LibreOffice 的导入 / 导出过滤器，不保证任意两种格式都能互转。

## watermark：加水印

```
watermark 文件 --text "文字" [-o 输出路径] [--overwrite]
```

支持 PDF 与 Word。默认写到 `原文件名_watermarked.扩展名`，不改原文件。

## merge：合并

```
merge 文件1 文件2 … -o 输出路径 [--overwrite]
```

全部为 PDF 或全部为 .docx。Word 以第一份文档为底（沿用其样式与页面设置），后续文档另起一页。
任一文件不存在时报 `FILE_NOT_FOUND`，不会只合并一部分。

## theme：主题

```
theme list
theme show 主题
theme validate 主题 [--no-fonts]
theme init 新主题名 [--extends 父主题] [--accent #RRGGBB] [--brand 品牌] [--title 显示名称] [--dir 目录] [--force]
theme preview 主题 [-o 目录] [--formats html,docx,pdf,pptx,xlsx]
```

- `list`：可用主题（内置、用户主题目录、环境变量 `ACKS_OFFICE_THEMES` 列出的目录）与默认主题。
- `show`：主题的版式变体、字体、颜色等摘要。
- `validate`：检查字段与取值、文字对比度（WCAG AA）、字体是否安装与授权、常用汉字覆盖、页眉页脚模板；
  有错误时 `ok` 为 false，错误码 `THEME_INVALID`，`data` 里是完整报告。
- `init`：在用户主题目录新建继承父主题的主题包（默认继承 neutral），只写和父主题不同的部分，生成后即可按名称使用；
  已存在时报 `THEME_EXISTS`，`--force` 只覆盖主题文件。
- `preview`：生成一页 HTML 样张（色板、对比度、字体、字号、各格式版式示意，可直接展示给用户），
  以及用同一份示例内容生成的 Word、PDF、PPT、Excel。默认目录为 `<主题>-preview`。

## fonts：字体

```
fonts list
fonts install 字体键 [--system]
fonts install --theme 主题 [--system]
```

`list` 返回 PDF 将使用的中文字体、可下载的开源字体（`download_mb` 为下载大小）、各内置主题的字体情况
（`present` 已安装、`substituted` 用备选字体替代、`missing` 都没装、`install` 可下载的字体键）。

`install` 从固定版本地址下载、校验哈希后生成固定字重的 TrueType 文件，放在用户数据目录（生成 PDF 时嵌入）；
`--theme` 安装该主题缺少的全部字体；`--system` 同时装到当前用户的字体目录，Word、PowerPoint 等软件也能用
（不需要管理员权限）。可下载：`noto-sans-sc`、`noto-serif-sc`、`source-sans-3`、`source-serif-4`、
`fraunces`、`lxgw-wenkai`、`jetbrains-mono`（均为 SIL OFL 1.1，可免费商用）。可变字体需要 fonttools。
**下载与安装前要征得用户同意。**

## doctor：环境检查

```
doctor [--no-network] [--no-probe]
```

只读，不安装、不修改任何东西。字段说明见 [doctor.md](doctor.md)。

## 错误码

| code | 含义 | 处理 |
|---|---|---|
| `USAGE_ERROR` | 参数不对 | 按 `message` 修正参数 |
| `FILE_NOT_FOUND` | 输入文件不存在（合并时任一文件不存在也报这个） | 核对路径 |
| `OUTPUT_EXISTS` | 输出文件已存在 | 问用户是否覆盖，同意后加 `--overwrite`，或换路径 |
| `INVALID_INPUT` | 输入内容格式不对（JSON、表格、幻灯片、样张格式），或输出扩展名与类型不符 | 按 `message` / `hint` 修正 |
| `THEME_NOT_FOUND` | 找不到主题 | `theme list` 查看可用主题 |
| `THEME_INVALID` | 主题校验有错误 | 按 `data.errors` 修改主题文件 |
| `THEME_EXISTS` / `BAD_THEME_NAME` / `BAD_COLOR` | `theme init` 的名称、颜色或目录问题 | 按 `hint` 修正 |
| `DEPENDENCY_MISSING` | 缺 Python 依赖 | 运行 doctor，按计划补齐 |
| `ENGINE_UNAVAILABLE` | 需要 LibreOffice | 运行 doctor；或直接生成目标格式 |
| `UNSUPPORTED_FORMAT` | 不支持的格式：读取只支持 docx / pdf / xlsx / xls / pptx；合并需全部同一格式 | `.doc` / `.ppt` 可先 `convert --to docx` / `pptx` |
| `CREATE_FAILED` / `EXTRACT_FAILED` / `CONVERSION_FAILED` / `WATERMARK_FAILED` / `MERGE_FAILED` | 执行出错 | 把 `message` 告诉用户 |
| `FONT_INSTALL_FAILED` | 字体下载或校验失败 | 检查网络，或设 `ACKS_OFFICE_DOWNLOAD_MIRROR` |
| `PACKAGE_MISSING` | 找不到程序本体（入口脚本给出） | 按 `data.plan` 安装，需用户同意 |
| `PYTHON_TOO_OLD` | Python 版本过低（入口脚本给出） | 请用户升级到 Python 3.10 或更高 |
| `INTERNAL_ERROR` | 未预料的错误 | 把 `message` 告诉用户，可附上 doctor 结果反馈问题 |

## 环境变量

| 变量 | 作用 |
|---|---|
| `ACKS_OFFICE_HOME` | 用户数据目录（字体、用户主题、专用虚拟环境），默认见 doctor 的 `paths.data_dir` |
| `ACKS_OFFICE_THEMES` | 额外的主题目录（多个用系统路径分隔符隔开），排在用户主题目录之前 |
| `ACKS_OFFICE_SOFFICE` | LibreOffice 可执行文件路径 |
| `ACKS_OFFICE_PDF_FONT` / `ACKS_OFFICE_PDF_FONT_BOLD` | PDF 中文字体文件，代替主题的中文字体 |
| `ACKS_OFFICE_DOWNLOAD_MIRROR` | 字体下载镜像，替换 `https://raw.githubusercontent.com` |
| `ACKS_OFFICE_NO_VENV` | 设为 1 时入口脚本不切换到专用虚拟环境 |
