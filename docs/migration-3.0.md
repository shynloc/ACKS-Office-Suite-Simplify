# 从 2.x 升级到 3.0 / Migrating from 2.x to 3.0

[中文](#中文) · [English](#english)

## 中文

3.0 把设计做成了主题：同一份内容可以按 `neutral`、`slate`、`folio` 或自己的主题出品。
大多数调用不改代码也能运行，但默认样式、部分返回值和几处文件写入行为变了。

### 1. 默认主题从 ACKS 改为 neutral

| 2.x | 3.0 |
|---|---|
| 不传 `theme` 时为 `acks`（ACKS 品牌样式） | 不传时为 `neutral`：中性配色、系统字体、不带品牌 |
| `theme="default"`：各格式的简单样式 | `default` 是 `neutral` 的别名 |
| `theme="acks"` | 暂时保留 2.x 的 ACKS 样式，只支持基础版式 |

需要旧样式时显式传 `theme="acks"`（CLI：`--theme acks`）。新文档建议改用 `slate` 或自己的主题：
`acks-office theme init 品牌名 --extends slate --accent "#色值" --brand "品牌"`。

### 2. 函数式接口取代 OfficeSuite

```python
# 2.x
from acks_office import OfficeSuite
result = OfficeSuite().create("word", title="报告", content=md, output_path="报告.docx")
if not result["success"]:
    raise RuntimeError(result["error"])

# 3.0
import acks_office
result = acks_office.create("word", "报告.docx", title="报告", content=md, theme="slate")
```

| OfficeSuite 方法 | 3.0 函数 |
|---|---|
| `create(kind, **kwargs)` | `acks_office.create(kind, output_path, **kwargs)` |
| `extract_data(path)` | `acks_office.extract(path)`：直接返回数据 |
| `convert(path, to, output_path)` | `acks_office.convert(path, to, output_path)` |
| `add_watermark(path, text, output_path)` | `acks_office.add_watermark(path, text, output_path)` |
| `pdf.merge_pdfs` / `docx.merge_documents` | `acks_office.merge(paths, output_path)` |

新函数出错时抛出异常：`FileNotFoundError`（文件不存在）、`ValueError`（参数或格式不支持）、
`acks_office.themes.ThemeError`（主题不存在或有误，`code` 属性为错误码）等。
`OfficeSuite` 与 `import office_suite` 仍可用到 4.0，使用时发出 `DeprecationWarning`；
邮件、批量与工作流功能目前仍在 `OfficeSuite` 上。

### 3. 不再默认覆盖原文件

- `add_watermark` 不传输出路径时写到 `<原名>_watermarked.<扩展名>`（2.x 会覆盖原文件）。
- `pptx.add_transition_effects` 不传输出路径时写到 `<原名>_transitions.pptx`。
- 要修改原文件时，把原路径显式传为输出路径。

### 4. 合并缺文件时报错

`merge_pdfs`、`merge_documents` 和 CLI `merge` 遇到不存在的输入文件时直接报错（CLI：`FILE_NOT_FOUND`），
不再跳过后只合并一部分；结果里也不再有 `skipped`。

### 5. 返回值变化

| 结果 | 2.x | 3.0 |
|---|---|---|
| Word | `pages`（按段落数估算） | 不再返回页数 |
| PDF | `font`（中文字体信息） | `fonts`（嵌入的字体家族列表）；中文没能嵌入时有 `FONT_NOT_EMBEDDED` 提醒 |
| 全部 | — | `theme`（实际使用的主题）；缺主题字体时 `FONT_SUBSTITUTED` 提醒 |
| Excel | `rows` 含表头行 | `rows` 为数据行数；另有 `sheets`；给了标题时第 1 行是标题行 |

读取带标题行的 Excel 时，`extract` 会自动从第 2 行读表头；用 pandas 直接读时传 `header=1`。

### 6. Python 3.10

最低版本从 3.9 升到 3.10。macOS 自带的 Python 3.9 需要另装新版，`doctor` 会给出安装计划。

## English

3.0 turns the design into themes: the same content can be rendered with `neutral`, `slate`, `folio`
or your own theme. Most calls keep working, but the default look, some return values and a few
file-writing defaults changed.

1. **Default theme is `neutral`** (neutral colors, system fonts, no brand). `theme="default"` is now an
   alias of `neutral`; the 2.x "simple" style is gone. `theme="acks"` keeps the 2.x ACKS style for now
   (basic layouts only). Create your own with `acks-office theme init my-brand --extends slate`.
2. **Function API**: `acks_office.create(kind, output_path, **kwargs)`, `extract`, `convert`,
   `add_watermark` and `merge` raise exceptions (`FileNotFoundError`, `ValueError`, `ThemeError`)
   instead of returning `{"success": False}`. `OfficeSuite` and `import office_suite` still work until
   4.0 and emit a `DeprecationWarning`.
3. **No silent overwrites**: `add_watermark` writes `<name>_watermarked.<ext>` and
   `add_transition_effects` writes `<name>_transitions.pptx` unless you pass an output path.
4. **Strict merge**: a missing input file is an error; nothing is merged partially.
5. **Results**: Word no longer returns an estimated `pages`; PDF returns `fonts` instead of `font`;
   every result includes `theme`; Excel `rows` counts data rows and a title row is added when a title is given
   (`extract` detects it; with pandas use `header=1`).
6. **Python 3.10** or newer is required.
