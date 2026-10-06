# ACKS Office Suite Simplify

> 面向 AI Agent 和 Python 程序的办公工具：生成、读取 Word / Excel / PPT / PDF，提供命令行、环境检查和技能包。

文档对应版本：**2.1.2**。GitHub 与 PyPI 使用本文件作为项目说明；历史版本的说明随发行包保留。

<p align="center">
  <a href="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml"><img alt="tests" src="https://github.com/shynloc/ACKS-Office-Suite-Simplify/actions/workflows/tests.yml/badge.svg?branch=main"></a>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white">
  <img alt="Platform" src="https://img.shields.io/badge/平台-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey">
  <a href="https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/LICENSE"><img alt="License" src="https://img.shields.io/badge/License-MIT-green"></a>
</p>

由 **ACKS Studio** 出品。安装包名 `acks-office`，导入名 `acks_office`。
生成与读取文件由 Python 库完成，不需要 Microsoft Office 或 WPS；格式转换需要 LibreOffice。

## 当前能力

| 格式 / 入口 | 已实现 | 范围与条件 |
|---|---|---|
| Word `.docx` | Markdown 正文、表格、提示块、代码、图片、链接；提取正文与表格；合并、水印 | `acks` / `default` 两种主题 |
| PDF | Markdown 排版、提取已有文字、合并、水印；Python API 按页拆分 | 有可用中文 TrueType 字体时嵌入，否则回退并报告 `FONT_FALLBACK`；扫描文字需外部 OCR |
| PPT `.pptx` | 封面与内容页、文字提取、品牌与页脚参数；Python API 添加切换效果 | `acks` / `default` 两种主题；切换效果通过 Python API 调用 |
| Excel | 生成 `.xlsx`，读取 `.xlsx` / `.xls`，保留文本型编号、日期转 ISO 字符串 | `default` 主题可添加柱状图 |
| `convert` | 委托 LibreOffice 转换文件，如 DOCX → PDF | 转换方向取决于引擎的导入 / 导出过滤器 |
| `doctor` | 只读检查 Python、依赖、字体、LibreOffice、宿主线索，给出补齐计划 | 不执行安装；其他 Office 程序只检测、不调用 |
| Python API 扩展 | SMTP 邮件、批量操作、顺序执行工作流配置 | 邮件需 SMTP 配置；API 数据源 / YAML 示例使用 `workflow` 可选依赖 |

直接读取支持 `.docx`、`.pdf`、`.pptx`、`.xlsx`、`.xls`。
旧版 `.doc` / `.ppt` 可先交给 LibreOffice 转换。格式转换不保证任意两种格式都能互转。

## 安装与升级

推荐在专用虚拟环境中从 [PyPI](https://pypi.org/project/acks-office/) 安装：

```bash
python -m pip install acks-office==2.1.2
acks-office doctor --json
```

升级到最新发行版：

```bash
python -m pip install --upgrade acks-office
```

下载开源中文字体需要字体可选依赖：

```bash
python -m pip install "acks-office[fonts]==2.1.2"
acks-office fonts install noto-sans-sc
```

各系统虚拟环境、依赖、字体和 LibreOffice 的准备步骤见
[安装指南](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/INSTALL_GUIDE.md)。
也可从 [GitHub Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v2.1.2)
下载 wheel 或技能包；功能不依赖发布渠道。

## 三种入口

### Agent 技能

技能包由 `SKILL.md`、Python 入口脚本和程序代码组成，供支持技能规范及本地命令执行的 Agent 使用。
宿主识别依据环境变量、目录等线索给出推测；接入方式与路径以宿主实际配置为准。

1. 从 [GitHub Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v2.1.2)
   下载 `acks-office-skill-2.1.2.zip`。
2. 解压到宿主技能目录，得到 `acks-office/SKILL.md`。
   Claude Code、Codex、OpenClaw、WorkBuddy、Hermes 的目录示例见安装指南。
3. Agent 运行 `scripts/acks.py doctor --json`，说明补齐计划并取得用户同意后，
   安装依赖到专用虚拟环境。入口脚本随后使用该环境。

技能包含程序代码；依赖安装由 Agent 或用户执行，脚本本身不会自动安装。

### 命令行

子命令：`create`、`extract`、`convert`、`watermark`、`merge`、`theme`、`fonts`、`doctor`、`version`。
以下输入文件由用户准备，数据格式见
[命令参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/skills/acks-office/references/commands.md)。

```bash
acks-office create word -o 报告.docx --title "经营回顾" --content-file 正文.md
acks-office create pdf -o 报告.pdf --title "经营回顾" --content-file 正文.md
acks-office create pptx -o 汇报.pptx --slides-file slides.json --brand-name "栖木咖啡"
acks-office create excel -o 数据.xlsx --data-file data.csv
acks-office extract 报告.docx
acks-office convert 报告.docx --to pdf -o 转换版.pdf
acks-office watermark 报告.pdf --text "内部资料"
acks-office merge a.pdf b.pdf -o 合并.pdf
```

每个子命令可加 `--json`，返回 `{ok, data, artifacts, warnings, error}`；
退出码 0 成功、1 执行失败、2 参数错误。默认不覆盖已有输出，需显式加 `--overwrite`。
命令不在 PATH 上时，可用 `python -m acks_office`。

### Python 库

```python
from acks_office import OfficeSuite

suite = OfficeSuite(theme="acks")
result = suite.create(
    "word", title="经营回顾",
    content="# 概览\n\n营收 **增长 24%**。\n\n"
            "| 区域 | 营收 |\n|---|--:|\n| 华东 | 5888 |",
    output_path="报告.docx", brand_name="",
)
if not result["success"]:
    raise RuntimeError(result["error"])

data = suite.extract_data("报告.docx")
```

`OfficeSuite` 的操作返回 `success`；失败时返回 `error`。检查返回值后再报告成功。
底层格式模块可能直接抛出异常。Python API 与 CLI 的覆盖策略不同，见
[安全说明](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/SECURITY.md)。

PDF 拆分与 PPT 切换通过格式模块 API 调用：

```python
from acks_office.pdf import split_pdf
from acks_office.pptx import add_transition_effects

split_pdf("报告.pdf", "拆分页")
add_transition_effects("汇报.pptx", "汇报_淡入.pptx", effect="fade", duration=0.5)
```

PPT 切换支持 fade、push、wipe、split、cover、pull、dissolve、cut、zoom、random；
时长参数映射为快 / 中 / 慢三档，实际播放由演示软件决定。

## 主题、品牌与字体

- Word / PPT / Excel 的 `acks` 主题使用内置设计令牌，`default` 使用各格式的基础样式。
- Word / PPT 在 `acks` 主题下支持 `brand_name`、`footer_label`；传 `brand_name=""` 去掉品牌字样。
- PDF 使用独立的 ReportLab 排版，当前不会因为 `theme` 或品牌参数切换样式。
- Word / PPT 写入字体名称，打开文件的电脑缺字体时由显示软件替换。
- PDF 字体选择依次为参数 `font`、环境变量 `ACKS_OFFICE_PDF_FONT`、本工具下载的字体、可用系统字体。
  能否嵌入以 doctor 检测为准，不能仅凭操作系统名称判断。
  没有可用字体时回退到未嵌入的 STSong-Light，并返回 `FONT_FALLBACK`。

正文写法见
[Markdown 参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/skills/acks-office/references/markdown.md)，
环境等级与字段见
[doctor 参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/skills/acks-office/references/doctor.md)。

## 工作流范围

`OfficeSuite.execute_workflow(config)` 按顺序执行步骤，调度由调用方负责。
加载 YAML 或使用 API 数据源时，可安装 `python -m pip install "acks-office[workflow]"`。
示例见
[workflow_example.yaml](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/examples/workflow_example.yaml)。

当前不执行 `schedule`、`enabled`、`on_success`、`on_failure` 等根配置，也不读取 Word 模板文件；
`retry` 顺序步骤只返回跳过，`notification` 只打印通知，不发送到 Slack / 企业微信。
这些能力需要调用方自行实现。

## 后续设计方向与兼容层

自定义主题文件、Slate / Folio 主题、`theme init` / `theme preview` 属于后续规划，
当前发行版没有这些入口。样张定稿见
[设计参考](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/skills/acks-office/references/design-system.md)。
后续功能是否上线，以代码、测试和变更记录为准。

`import office_suite` 是仍可用的旧名兼容层，会发出弃用提示；新代码使用 `acks_office`。
程序提供技能与 CLI，不包含内置 MCP 服务或邮件配置加密功能。

## 验证与安全

最低 Python 3.9；CI 覆盖 Ubuntu / macOS / Windows × Python 3.9 / 3.12。
无 LibreOffice 的 runner 跳过实际转换测试，生成与读取测试仍运行。

文档在本机处理。doctor 联网检查可用 `--no-network` 关闭；
字体下载校验固定版本 SHA-256；SMTP 默认验证证书并拒绝明文登录。
版本变化见
[CHANGELOG](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/CHANGELOG.md)，
安全与环境变量见
[SECURITY.md](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/SECURITY.md)。

## 开发与文档维护

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
python -m pip install -e ".[dev]"
python tools/check_release_docs.py
python -m pytest
python test_integration.py
python tools/build_skill_bundle.py
```

README 是 GitHub 与 PyPI 的共同说明来源。功能新增、移除或弃用时同步更新文档，
将规划和可用功能分开，维护流程见
[CONTRIBUTING.md](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/CONTRIBUTING.md)。

[MIT 许可证](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/LICENSE)。
设计思路受 [MiniMax Office Skill](https://www.minimaxi.com/) 启发。
