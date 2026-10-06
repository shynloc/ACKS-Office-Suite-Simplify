# 命令参考

入口：`python3 <技能目录>/scripts/acks.py`（已用 pip 安装时也可用 `acks-office` 或
`python -m acks_office`，参数相同）。`--json` 放在子命令前后都可以。

## 输出格式

```json
{
  "ok": true,
  "data": {},
  "artifacts": [{"path": "/abs/path/报告.docx", "size": 38211}],
  "warnings": [{"code": "FONT_FALLBACK", "message": "…"}],
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
| `--title` | 全部 | 标题，默认用输出文件名 |
| `--content` / `--content-file` | Word、PDF | 正文 Markdown；`--content-file -` 从标准输入读取 |
| `--data-file` | Excel | JSON 二维数组（第一行为表头）或 CSV（第一行为表头，数字自动识别） |
| `--slides-file` | PPT | JSON 数组：`title`、`content`（用 `\n` 分行）、`layout`（`title` 或 `content`）、`subtitle` |
| `--theme` | Word、PPT、Excel | `acks`（默认）或 `default` |
| `--brand-name` | Word、PPT | 品牌名；`""` 去掉品牌 |
| `--footer-label` | Word、PPT | 页脚文字 |
| `--subtitle` | Word | 封面副标题 |
| `--font` | PDF | 中文 TrueType 字体文件（.ttf/.ttc） |
| `--chart` | Excel | 附柱状图（`default` 主题） |
| `--overwrite` | 全部 | 允许覆盖已有文件 |

`data` 含 `output_path`、`file_size`，以及：Word/PDF 的 `pages`（Word 为估算），PDF 的 `font`
（实际使用的中文字体与是否嵌入），PPT 的 `slides_count`。

## extract：读取内容

```
extract 文件 [--sheet 名称或从0开始的序号]
```

- `.docx`、`.pdf`：`data.text`（Word 表格按行输出，单元格用 ` | ` 分隔）
- `.xlsx`、`.xls`：`data.rows`，每行一个对象（表头为键；空单元格为 null；日期为 ISO 字符串）
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
不存在的文件会跳过并给出 `INPUT_SKIPPED` 提醒。

## fonts：字体

```
fonts list
fonts install noto-sans-sc
```

`list` 返回 PDF 将使用的中文字体、可下载的开源字体、主题字体的安装情况。
`install` 从固定版本地址下载、校验 SHA-256 后生成常规与粗体两个字重，放在用户数据目录（需要
fonttools；用户同意后再执行）。

## doctor：环境检查

```
doctor [--no-network] [--no-probe]
```

只读，不安装、不修改任何东西。字段说明见 [doctor.md](doctor.md)。

## 错误码

| code | 含义 | 处理 |
|---|---|---|
| `USAGE_ERROR` | 参数不对 | 按 `message` 修正参数 |
| `FILE_NOT_FOUND` | 输入文件不存在 | 核对路径 |
| `OUTPUT_EXISTS` | 输出文件已存在 | 问用户是否覆盖，同意后加 `--overwrite`，或换路径 |
| `INVALID_INPUT` | 输入内容格式不对（JSON、表格、幻灯片），或输出扩展名与类型不符 | 按 `message` / `hint` 修正 |
| `DEPENDENCY_MISSING` | 缺 Python 依赖 | 运行 doctor，按计划补齐 |
| `ENGINE_UNAVAILABLE` | 需要 LibreOffice | 运行 doctor；或直接生成目标格式 |
| `UNSUPPORTED_FORMAT` | 不支持的格式：读取只支持 docx / pdf / xlsx / xls / pptx；合并需全部同一格式 | `.doc` / `.ppt` 可先 `convert --to docx` / `pptx` |
| `CREATE_FAILED` / `EXTRACT_FAILED` / `CONVERSION_FAILED` / `WATERMARK_FAILED` | 执行出错 | 把 `message` 告诉用户 |
| `FONT_INSTALL_FAILED` | 字体下载或校验失败 | 检查网络，或设 `ACKS_OFFICE_DOWNLOAD_MIRROR` |
| `PACKAGE_MISSING` | 找不到程序本体（入口脚本给出） | 按 `data.plan` 安装，需用户同意 |
| `PYTHON_TOO_OLD` | Python 版本过低（入口脚本给出） | 请用户升级 Python |
| `INTERNAL_ERROR` | 未预料的错误 | 把 `message` 告诉用户，可附上 doctor 结果反馈问题 |

## 环境变量

| 变量 | 作用 |
|---|---|
| `ACKS_OFFICE_HOME` | 用户数据目录（字体、专用虚拟环境），默认见 doctor 的 `paths.data_dir` |
| `ACKS_OFFICE_SOFFICE` | LibreOffice 可执行文件路径 |
| `ACKS_OFFICE_PDF_FONT` / `ACKS_OFFICE_PDF_FONT_BOLD` | PDF 中文字体文件 |
| `ACKS_OFFICE_DOWNLOAD_MIRROR` | 字体下载镜像，替换 `https://raw.githubusercontent.com` |
| `ACKS_OFFICE_NO_VENV` | 设为 1 时入口脚本不切换到专用虚拟环境 |
