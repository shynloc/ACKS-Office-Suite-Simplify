# 更新记录

格式参照 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循
[语义化版本](https://semver.org/lang/zh-CN/)。

## [3.0.0] - 2026-10-10

主题引擎：设计从写死的常量变成运行时加载的主题包，同一套引擎按主题出品 Word、PDF、PPT、Excel。
升级说明见 [迁移指南](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v3.0.0/docs/migration-3.0.md)。

### 不兼容变更

- **默认主题改为 `neutral`**（中性、只用系统字体、不带品牌）。`theme="default"` 成为 `neutral` 的别名，
  2.x 的「简单样式」随之移除。
- **ACKS 样式移出开源仓库**：`theme="acks"` 不再内置，`acks_office.design_system` 与 `office_suite.design_system`
  一并移除；ACKS 品牌改为单独安装的 `acks` 主题包。
- **函数式接口** `acks_office.create / extract / convert / add_watermark / merge`：出错时抛出异常，
  不再返回 `{"success": False}`。`OfficeSuite` 与 `office_suite` 导入名保留到 4.0，使用时提示弃用。
- **不覆盖原文件**：`add_watermark` 默认写到 `<原名>_watermarked.<扩展名>`，`add_transition_effects`
  默认写到 `<原名>_transitions.pptx`；要改原文件时显式传入原路径。
- **合并更严格**：任一输入文件不存在时报错（CLI 为 `FILE_NOT_FOUND`），不再跳过后只合并一部分。
- Word 的结果不再返回估算的页数（PDF 仍返回实际页数）；PDF 结果的 `font` 改为 `fonts`（嵌入的字体家族）。
- 最低 Python 版本升到 3.10。

### 新增

- **主题引擎**：主题包由 `theme.json`、`tokens.json`、`chrome.json`、`fonts.json` 组成，沿 `extends` 继承，
  可放在用户主题目录、`ACKS_OFFICE_THEMES` 指定的目录或任意路径（`--theme 路径`）。
  内置 `neutral`、`slate`（商务）、`folio`（杂志）三套主题，取值来自定稿样张。
- **Word**：三种封面（标题块 / 独立封面 / 期刊封面与本期目录）、章节自动编号或章节页、正文分栏与首字下沉、
  主题化的表格（表头与合计行变体）、提示块、引文、表题图题，页眉页脚来自主题模板。
- **PDF** 由引擎直接排版，版式与 Word 一致，字体嵌入文件；中文排版补上行首禁则、拆分段落与两端对齐的处理。
- **PPT**：封面、章节页、内容页、数据页（关键数字 + 条形图）、表格页、大数字页、引文页、图片页，
  标题按实际字宽自动缩放，中西文分别设置字体，可加演讲者备注。
- **Excel**：标题行、表头与合计行变体、千位分隔与负数色、重点单元格、说明行、多张工作表与图表；
  数据可以是二维数组、对象数组或多张工作表的说明；读取时识别标题行。
- **Markdown front matter** 与 `--meta 键=值`：副标题、作者、日期、版本、期号、导语等元数据；
  标题可带 `{label="…"}`，表格前后的「表：…」作为表题。
- **主题命令**：`theme list / show / validate / init / preview`。校验对比度、字体安装与授权、常用汉字覆盖；
  `init` 生成继承内置主题的主题包；`preview` 生成 HTML 样张与四种格式的示例文件。
- **字体**：主题用到的七款开源字体可下载（地址固定、校验哈希，可变字体生成静态字重并改好名称）；
  `fonts install --theme <主题> [--system]` 安装主题缺的全部字体，`--system` 同时装到当前用户的字体目录。
  缺主题字体时自动改用本机备选字体，并给出 `FONT_SUBSTITUTED` 提醒。
- `fonts list` 与 `doctor` 按主题报告字体情况；doctor 计划给出可选的安装主题字体步骤。

### 修复

- 中文 macOS 上用 LibreOffice 转换时，中文字体被逐字替换成其他字体（例如手写体）。

## [2.1.2] - 2026-10-06

### 文档

- 移除第三方动态版本徽章，避免图片缓存继续显示旧版本；版本以正文和发行页面的版本号为准。
- GitHub / PyPI 同步发布新的文档快照，功能与 2.1.1 一致。

## [2.1.1] - 2026-10-06

### 文档与发布

- GitHub / PyPI 共用 README，安装入口更新为 PyPI，文档链接改为同版本的 GitHub 完整地址。
- 按实际实现说明 PDF 排版与字体条件、主题和品牌范围、CLI / Python API 区别及 LibreOffice 转换边界。
- 工作流示例移除未执行的调度、模板、回调和外部通知配置；使用示例按返回值报告结果。
- 补充文档维护规则；CI 与生产发布检查版本、CLI 入口清单和文档链接，生产发布须使用匹配版本标签。
- 程序功能与 2.1.0 一致；本版本用于同步发行包中的项目说明、元数据和示例。

## [2.1.0] - 2026-10

面向 AI Agent 的版本：可在 macOS、Windows、Linux 上不依赖本机 Office 出品，首次使用时自检环境，
并以技能（Skill）形式接入各类 Agent。

### 新增

- **命令行 `acks-office`**（也可 `python -m acks_office`）：`create`、`extract`、`convert`、`watermark`、
  `merge`、`fonts`、`doctor`、`version`。加 `--json` 输出统一结构 `{ok, data, artifacts, warnings, error}`，
  错误带 `code` 与 `hint`；退出码 0 成功、1 失败、2 参数错误；默认不覆盖已有文件（`--overwrite` 才覆盖）。
- **`doctor` 环境检查**（只读）：识别宿主 Agent，给出能力等级（none / L0 / L1 / L2），报告依赖、字体、
  LibreOffice 与本机其他 Office 程序（Word、WPS、Keynote 等只报告，不调用）、网络，并列出补齐计划；
  每一步都是可直接执行的参数列表，执行前需征得用户同意。
- **Agent 技能 `skills/acks-office/`**：SKILL.md（符合 Agent Skills 规范）、入口脚本 `scripts/acks.py`、
  参考文档。入口脚本自动使用专用虚拟环境、优先使用技能自带的代码，找不到程序时给出安装计划。
  `tools/build_skill_bundle.py` 生成可直接解压使用的技能包（含程序代码）。
- **PDF 中文字体**：依次使用参数 `font`、环境变量 `ACKS_OFFICE_PDF_FONT`、下载的开源字体、系统自带中文字体，
  并嵌入 PDF；都没有时回退到 STSong-Light 并给出 `FONT_FALLBACK` 提醒。
  `acks-office fonts install noto-sans-sc` 下载 Noto Sans SC（SIL OFL 1.1，可免费商用），校验 SHA-256。
- **Markdown**：Word 与 PDF 支持粗体、斜体、删除线、行内代码、链接、嵌套列表、任务清单、表格（数字列右对齐）、
  引用、GitHub 提示块（`> [!NOTE]` 等）、代码块、分隔线、图片。
- **品牌参数**：Word 与 PPT 的 `brand_name=""` 去掉全部品牌字样，传入其他名称即换成该品牌；
  `footer_label` 指定页脚文字；Word 新增封面副标题 `subtitle`。
- **PPT 切换效果** `add_transition_effects` 真正写入文件：fade、push、wipe、split、cover、pull、dissolve、
  cut、zoom、random；不支持的效果报错。
- `extract_data` 支持 `.pptx`。
- LibreOffice 自动查找：环境变量 `ACKS_OFFICE_SOFFICE` → PATH → 各系统标准安装位置（如 macOS 的
  `/Applications/LibreOffice.app`）。
- 邮件密码可从环境变量 `OFFICE_EMAIL_PASSWORD`（或 `ACKS_OFFICE_EMAIL_PASSWORD`）读取；邮件带 `Date` 与
  `Message-ID`；不存在的附件列在结果的 `missing_attachments` 里。
- 打包改用 `pyproject.toml`，可 `pip install` 安装并获得 `acks-office` 命令；版本号只在
  `acks_office/__init__.py` 维护。
- 持续集成：Ubuntu、macOS、Windows × Python 3.9、3.12。

### 修复

- **密送泄露**：Bcc 地址以前写进了邮件头，所有收件人都能看到密送名单；现在只用于投递。
- **邮件传输安全**：以前不校验邮件服务器证书（Python smtplib 的默认行为），服务器不支持 STARTTLS 时还会
  明文发送密码；现在校验证书与主机名，并拒绝明文登录。
- **PDF 中文**：以前只引用不嵌入的 STSong-Light，部分阅读器显示为方块或空白；现在嵌入中文字体。
- PDF 正文中的 `<`、`>`、`&` 以前会被当作标记，导致报错或丢字；现在原样显示。
- PDF 结果中的 `pages` 以前是估算值，现在是实际页数。
- Word 提取文本以前漏掉表格；现在按原文顺序输出，表格每行一行、单元格用 ` | ` 分隔。
- Word 合并以前会丢失图片和链接（关系 ID 失效）；现在一并复制。
- ACKS 主题 Word 的中文标题在封面上重复出现两次；PPT 中文封面标题前多出一个空格。
- `default` 主题 Word：每个有序列表从自己的起始号重新编号（以前接着上一个列表往下编）；
  任务项不再叠加项目符号；小图片不再被放大到整个版心宽度。
- LibreOffice 转换：不在 PATH 上时找不到（macOS 默认安装即如此）；用户正开着 LibreOffice 时转换静默失败；
  没有生成文件也返回成功。现在使用独立的临时配置目录，并检查确实生成了输出文件。
- PPT 切换效果以前什么也没写，却返回成功。
- 设计令牌一致性测试显式按 UTF-8 读取源码，修复 Windows 默认编码导致的解码失败。
- 集成测试日志使用 ASCII，避免 Windows 重定向输出时因 emoji 与中文编码失败而中断。

### 行为变化（升级前请留意）

- **包名改为 `acks_office`**，发行名 `acks-office`。`import office_suite` 仍可用，会给出
  `DeprecationWarning`，3.0 移除。
- Markdown 段落内的单个换行保留为换行；空行分段（以前每一行都是一个段落）。
- Word 封面标题只出现一次；副标题用 `subtitle` 传入。PPT 页面的副标题取每页的 `subtitle`，不再重复标题。
- 页脚文字：默认品牌显示「ACKS Studio · 文档设计规范 v2」；自定义品牌显示品牌名；`brand_name=""`
  时显示文档标题；`footer_label` 优先。PPT 封面右上角的「CONFIDENTIAL · 2026」只在默认品牌下显示
  （可用 `meta_right` 指定）。
- Word 合并以第一份文档为底，沿用其样式与页面设置；后续文档另起一页。不存在的输入文件列在 `skipped` 里。
- PDF 默认嵌入中文字体（子集），文件会比以前大一些。
- 发邮件默认校验服务器证书并要求加密连接；内网自签名证书或不支持加密的服务器需显式传 `allow_insecure=True`。
- PDF 水印按每页实际尺寸居中（以前固定按 A4 计算，横版或非 A4 页面会偏位）。
- `requirements.txt` 改为安装本包（`.`），依赖只在 `pyproject.toml` 维护；移除 `setup.py`。
- 依赖：PyPDF2 换成其后继 pypdf（接口相同），新增 markdown-it-py；去掉代码中从未使用的可选依赖
  （Jinja2、click、rich、python-magic、chardet、tqdm、python-dotenv）。
- 最低 Python 版本 3.9。

## [2.0.1] - 2026-09-25

### 修复

- Excel `extract_data`：默认调用报错；读取参数被忽略；类型推断改坏数据（工号、身份证号）；
  结果不能转 JSON；`.xls` 无法读取（补充依赖 `xlrd`）。
- design_system 测试的导入方式（此前 29 项中 26 项失败）。

### 行为变化

- 日期类单元格返回 ISO 8601 字符串，空单元格返回 `None`；文本型数字保持字符串。

## [2.0.0] - 2026-08-23

- 集成 ACKS Studio 文档设计规范 v2（Word / PPT / Excel 主题），`theme` 参数切换 `acks` 与 `default`。
- 工作流引擎 `execute_workflow`。
- 项目更名为 ACKS Office Suite Simplify。

## [1.0.0] - 2026-04-09

- 首个版本：Word、Excel、PPT、PDF 的生成、转换、水印与邮件发送。

[3.0.0]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v2.1.2...v3.0.0
[2.1.2]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v2.1.1...v2.1.2
[2.1.1]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v2.1.0...v2.1.1
[2.1.0]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v2.0.1...v2.1.0
[2.0.1]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v2.0.0...v2.0.1
[2.0.0]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v1.0.0
