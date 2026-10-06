# 安装指南

文档对应版本：**2.1.2**。

## 系统要求

- Python 3.9 或更高；CI 验证 Python 3.9 / 3.12。
- macOS、Windows、Linux；生成与读取不需要 Microsoft Office 或 WPS。
- LibreOffice 用于转换；字体是否可用，以 doctor 检测结果为准。

## Agent 技能

从 [2.1.2 Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v2.1.2)
下载 `acks-office-skill-2.1.2.zip`，解压到宿主实际配置的技能目录，得到 `acks-office/SKILL.md`。

目录示例：Claude Code `~/.claude/skills/`、Codex `~/.codex/skills/`、
OpenClaw `~/.openclaw/skills/` 或工作区 `skills/`、WorkBuddy `~/.workbuddy/skills/`、
Hermes `~/.hermes/skills/`。以宿主配置为准。

```bash
python3 ~/.claude/skills/acks-office/scripts/acks.py doctor --json
```

Windows 使用 `python` 或 `py -3`，按实际技能路径调用。
技能包包含程序代码；依赖安装需要用户同意，由 Agent 或用户执行 doctor 的计划。
安装到专用虚拟环境后，入口脚本会使用该环境；脚本自身不安装依赖。

## PyPI 安装（专用虚拟环境）

macOS / Linux：

```bash
python3 -m venv ~/.venvs/acks-office
~/.venvs/acks-office/bin/python -m pip install acks-office==2.1.2
~/.venvs/acks-office/bin/acks-office doctor --json
```

Windows PowerShell：

```powershell
py -3 -m venv "$env:LOCALAPPDATA\acks-office-cli"
& "$env:LOCALAPPDATA\acks-office-cli\Scripts\python.exe" -m pip install acks-office==2.1.2
& "$env:LOCALAPPDATA\acks-office-cli\Scripts\acks-office.exe" doctor --json
```

在已激活的环境中，以下 `python` 指该环境的解释器：

```bash
python -m pip install --upgrade acks-office
python -m acks_office doctor --json
```

可选依赖：

```bash
python -m pip install "acks-office[fonts]==2.1.2"
python -m pip install "acks-office[workflow]==2.1.2"
```

系统 Python 提示 `externally-managed-environment` 时使用虚拟环境。
命令不在 PATH 上时，使用虚拟环境内的完整路径或 `python -m acks_office`。

## GitHub wheel 与源码

从 [2.1.2 Release](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases/tag/v2.1.2)
下载 wheel，在专用环境安装：

```bash
python -m pip install ./acks_office-2.1.2-py3-none-any.whl
```

源码开发：

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
python -m pip install -e ".[dev]"
python tools/check_release_docs.py
python -m pytest
python test_integration.py
```

可复现的发行内容以版本标签及发行包为准，主分支可能含开发改动。

## 字体与转换

不能仅凭操作系统判断中文字体可嵌入。doctor 的 `fonts.pdf_cjk.embedded` 为 false 时，
在用户同意后下载开源中文字体：

```bash
python -m pip install "acks-office[fonts]==2.1.2"
acks-office fonts install noto-sans-sc
```

也可用 `--font` / `ACKS_OFFICE_PDF_FONT` 指定包含中文字形的 TrueType `.ttf` / `.ttc`。
Word / PPT 使用主题字体名，缺字体时由打开文档的软件替换显示。

需要转换时，按系统选择对应 LibreOffice 安装命令：

```bash
brew install --cask libreoffice                          # macOS
winget install -e --id TheDocumentFoundation.LibreOffice # Windows
sudo apt-get install -y libreoffice                      # Ubuntu / Debian
sudo dnf install -y libreoffice                          # Fedora / OpenCloudOS
```

非标准位置可用 `ACKS_OFFICE_SOFFICE` 指定 soffice 完整路径。
转换方向取决于 LibreOffice 的支持。

## 验证与常见问题

运行 `acks-office doctor --json` 查看报告与补齐计划。
`none` 表示 Python / 基础依赖不满足，`L0` 表示基础依赖齐全，
`L1` 还要求 PDF 嵌入字体和主题字体齐全，`L2` 还要求 LibreOffice 可调用。
这些是环境检查等级，不保证所有阅读器中的显示完全一致。

- PDF 字体缺失：查看 `fonts.pdf_cjk`，按计划补齐；扫描图片文字需要外部 OCR。
- `ENGINE_UNAVAILABLE`：准备 LibreOffice，或直接生成目标格式。
- 国内包镜像可能滞后于 PyPI，刚发布版本可使用官方索引。字体镜像用
  `ACKS_OFFICE_DOWNLOAD_MIRROR` 配置，仍验证 SHA-256。
- 从旧版升级：`office_suite` 兼容层仍可用，新代码用 `acks_office`。
  变化见 [CHANGELOG](https://github.com/shynloc/ACKS-Office-Suite-Simplify/blob/v2.1.2/CHANGELOG.md)。
