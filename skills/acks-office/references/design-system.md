# 主题与设计规范

同一份内容，换一个主题就换一套设计：配色、字体、字号、版式、页眉页脚都由主题决定，
Word、PDF、PPT、Excel 共用同一套主题。主题是一个数据包（几份 JSON），不是代码。

## 内置主题

| 主题 | 定位 | 字体 | 版式特征 |
|---|---|---|---|
| `neutral`（默认） | 零依赖，任何环境都能出品 | 系统字体：苹方 / 微软雅黑 + Helvetica Neue / Segoe UI | 标题放在第一页顶部；黑灰加一个蓝色强调；表头浅底 |
| `slate` | 商务：报告、方案、汇报 | 思源黑体（Noto Sans SC）+ Source Sans 3；代码 JetBrains Mono | 独立封面，日期 / 编制 / 版本固定在页底；章节自动编号；钢蓝强调；PPT 封面左右分栏 |
| `folio` | 杂志：品牌手册、年报、作品集 | 思源宋体（Noto Serif SC）+ Source Serif 4；标题与大号数字 Fraunces；引文霞鹜文楷 | 期刊式封面与本期目录；每章一页章节页；正文两栏、首字下沉；朱砂点睛；PPT 章节页靛青底 |

`theme="default"` 是 `neutral` 的别名。`acks` 是 2.x 保留的 ACKS 品牌样式，只支持基础版式。

主题字体没装时，生成文件写入本机的备选字体（结果里有 `FONT_SUBSTITUTED` 提醒和可下载的字体键）；
`fonts install --theme slate --system` 在用户同意后补齐。PDF 只能嵌入 TrueType 轮廓的字体，
找不到时用同风格的系统字体（如黑体、宋体），都没有时中文由阅读器替换显示（`FONT_NOT_EMBEDDED`）。

## 主题包

```
my-brand/
├── theme.json    名称、显示名、说明、继承的主题（extends）、版本、授权
├── tokens.json   颜色、字体、字号、行距、间距、版式变体
├── chrome.json   品牌名、logo、页眉页脚模板、标签文字（可选）
├── fonts.json    字体的授权、备选字体、下载键（可选）
└── assets/       logo 等资源（可选）
```

各文件只写和父主题不同的部分，其余沿 `extends` 继承；所有主题最终都继承内部的 `_base`，
它包含全部字段的默认值。查找顺序：主题目录路径 → 环境变量 `ACKS_OFFICE_THEMES` 列出的目录 →
用户主题目录（`<数据目录>/themes`）→ 内置主题。同名时排在前面的生效，所以可以用同名主题覆盖内置主题，
并在 `extends` 里继续继承原来的内置主题。

新建：`theme init my-brand --extends slate --accent "#0B6E4F" --brand "品牌名"`，
生成后即可 `--theme my-brand` 使用。

## tokens.json

**颜色**（`color.*`，值为 `#RRGGBB`）：

| 名称 | 用途 |
|---|---|
| `ink` / `muted` | 正文与标题 / 次要文字、说明、页眉页脚 |
| `rule` / `rule_mid` / `rule_strong` | 细线 / 中等线（表头下方等）/ 重线（合计行、封面信息） |
| `surface` | 浅底：表头、提示块、代码块 |
| `paper` / `slide_paper` | Word、PDF 页面底色 / 幻灯片底色 |
| `accent` | 强调色：编号、关键数字、链接、强调线，面积尽量小 |
| `negative` | 负数 |
| `dark`、`on_dark`、`on_dark_soft`、`on_dark_muted`、`accent_on_dark` | 深色页（如章节页）的底色与其上的文字、强调色 |
| `chart` / `chart_muted` | 图表系列色（数组）/ 非重点系列 |
| `sheet_rule` | Excel 数据行之间的细线 |

版式里引用颜色时写名称（如 `"head_rule": "rule_mid"`），也可直接写 `#RRGGBB`。

**字体**：`font.families` 定义五个字体槽位 `sans`、`serif`、`display`、`quote`、`mono`，各有 `cn`（中文）
与 `en`（西文）；`font.roles` 把八个角色（`body`、`title`、`heading`、`label`、`number`、`display`、
`quote`、`code`）映射到槽位与字重（`{"family": "serif", "weight": 900}`，字重 ≥ 600 为粗体）。

**字号、行距、间距**：`type.doc.*`、`type.slide.*`、`type.sheet.*` 为各处字号（磅），
`leading.*` 为行距倍数，`space.*` 为段前段后与内边距（磅）。全部字段与默认值见程序包内的
`acks_office/themes/builtin/_base/tokens.json`。

**版式变体**：

| 字段 | 可选值 | 说明 |
|---|---|---|
| `layout.doc.cover` | `title-block` / `standard` / `issue` | 标题块 / 独立封面（信息在页底）/ 期刊封面（期号、导语、本期目录） |
| `layout.doc.chapter` | `inline` / `opener` | 普通标题 / 每章一页章节页（大号编号、标签、导语） |
| `layout.doc.heading_numbers` | `true` / `false` | 标题自动编号（1、1.1、1.1.1），`inline` 时生效 |
| `layout.doc.columns`、`column_gap_cm`、`drop_cap` | 1–3 栏、栏间距、`true` / `false` | 章节页之后的正文分栏与首字下沉 |
| `layout.doc.callout` | `tint` / `rule` | 提示块：浅底 / 上下细线 |
| `layout.doc.quote` | `bar` / `pull` | 引文：左侧竖线 / 上下线加大号引号 |
| `layout.doc.caption` | `plain` / `accent` | 表题、图题的编号：普通 / 强调色斜体 |
| `layout.doc.table.head` | `fill` / `label` | 表头：浅底 / 小号标签加中等线 |
| `layout.doc.table.total` | `rules` / `bar` | 合计行：上下各一条细线 / 上方一条粗线（颜色为 `total_rule`） |
| `layout.doc.page` | `size`（A4 / A5 / Letter）、`margin_cm`、`header_cm`、`footer_cm` | 纸张与页边距 |
| `layout.slide.cover` | `simple` / `split` / `issue` | PPT 封面：简洁 / 左标题右关键数字 / 期刊式 |
| `layout.slide.section` | `light` / `dark` | 章节页：浅底 / 深色底 |
| `layout.slide.content` | `stacked` / `columns` | 内容页：上下排列 / 左窄栏标题、右宽栏正文 |
| `layout.sheet.head`、`layout.sheet.total` | `fill` / `label`、`rules` / `bar` | Excel 表头与合计行 |
| `layout.sheet.*` 其他 | `negative`、`highlight`、`tab_color`、`gridlines`、`freeze_header` | 负数色、重点单元格框色、工作表标签色、网格线、冻结表头 |

## chrome.json

- `brand`、`classification`、`publication`：品牌名、密级、刊名（可被文档元数据覆盖）；`logo.light` / `logo.dark`：资源路径。
- `labels`：表、图、目录、期号、提示块（说明、提示、重要、警告、注意）、合计等文字，可改成其他语言。
- `doc.header` / `doc.footer` / `slide.footer`：`left`、`right` 两段模板，可用占位符 `{title}`、`{short_title}`、
  `{subtitle}`、`{brand}`、`{publication}`、`{issue}`、`{season}`、`{date}`、`{author}`、`{version}`、
  `{classification}`、`{chapter}`、`{page}`、`{pages}`。模板按「 · 」分段，一段里的占位符都没有值时整段省略。
- `doc.cover_meta`：封面信息项，如 `["date", "author", "version"]`。

## fonts.json

按字体家族写：`license`（如 `OFL-1.1`、`system`；商用字体写实际授权，只引用、不下载）、
`fallback`（备选字体，按顺序找本机已安装的）、`install`（可下载时的字体键，见 `fonts list`）。

## 校验

`theme validate` 检查：必填字段与取值范围、版式变体、颜色引用；文字对比度达到 WCAG AA
（正文 4.5:1，大字与线条 3:1）；字体是否安装、授权是否允许分发、中文字体是否覆盖常用汉字；
页眉页脚模板里的占位符与 logo 文件。有错误时主题仍可加载，但出品可能不符合预期，应先修正。

## 帮用户定制自己的主题

1. 先问清楚：品牌名称与 logo；主色与强调色（色值，或「沉稳商务」「杂志感」这样的方向）；
   中英文字体偏好（优先推荐可免费商用的开源字体，如思源黑体、思源宋体、霞鹜文楷、Source Sans 3）；
   使用场景（对内汇报、对外方案、印刷还是屏幕）；页眉页脚要显示什么。
2. 选一个接近的内置主题作为父主题，`theme init` 生成主题包，再按上文改 `tokens.json` / `chrome.json`。
3. `theme validate` 检查，`theme preview` 生成样张：把 HTML 样张展示给用户确认，示例文件可转成图片核对。
4. 用户确认后再用这个主题出品；缺字体时按 `fonts install` 的说明征得同意后安装。

## Slate / Folio 样张定稿

内置的 Slate 与 Folio 依据「主题样张：Slate 与 Folio」（2026-10-06）实现。样张中的品牌、人物、日期与
经营数据都是演示内容，不是主题默认值。

| 用途 | Slate | Folio |
|---|---|---|
| 墨色 / 次要文字 | `#1F2933` / `#52606D` | `#141414` / `#5E5A54` |
| 分隔线 / 浅底 | `#CBD2D9` / `#F5F7FA` | `#CBC6BC` / `#F3F1EC`（纸色，用作幻灯片底色；Word 与 Excel 为白底） |
| 强调色 | 钢蓝 `#1F5FAE` | 朱砂 `#C8321F` |
| 深色章节底 | — | 靛青 `#1F2A44` |
| 图表色 | `#1F5FAE`、`#8794A1`、`#4C9A8A`、`#C08A2E`、`#8B6BB1` | `#141414`、`#5E5A54`、`#9C968D`、`#C8321F` |
| 提示块 | 浅底矩形，上方一行钢蓝粗体小标签，无左侧色条、圆角和阴影 | 上下细线 |
| Excel | 表头浅底加粗；合计行上下深色细线 | 表头小号标签；合计行上方朱砂粗线 |

## PPT 切换效果

`acks_office.pptx.add_transition_effects(输入, 输出, effect=…, duration=…)` 把每页的切换写入 PPTX：
`fade`、`push`、`wipe`、`split`、`cover`、`pull`、`dissolve`、`cut`、`zoom`、`random`。
不传输出路径时写到 `<原名>_transitions.pptx`，不改原文件；再次设置会替换原有切换，不重复叠加。
`duration` 映射为快（≤ 0.5 秒）、中（≤ 1 秒）、慢三档。该能力通过 Python API 调用，CLI 没有对应子命令。
