# 安装指南

## 系统要求

- Python 3.9 或更高版本
- macOS、Windows、Linux 均可；**不需要安装 Microsoft Office 或 WPS**
- 可选：LibreOffice（仅格式转换需要）

## 方式一：Agent 技能

1. 从 [Releases](https://github.com/shynloc/ACKS-Office-Suite-Simplify/releases) 下载 `acks-office-skill-<版本>.zip`；
2. 解压到 Agent 的技能目录（如 `~/.claude/skills/`、`~/.codex/skills/`、`~/.openclaw/skills/`、
   `~/.workbuddy/skills/`、`~/.hermes/skills/`），得到 `acks-office/SKILL.md`；
3. 首次使用时 Agent 会运行环境检查，并在你同意后把依赖装进专用虚拟环境。

也可以手动检查：

```bash
python3 ~/.claude/skills/acks-office/scripts/acks.py doctor
```

## 方式二：pip 安装

```bash
python -m pip install "acks-office @ https://github.com/shynloc/ACKS-Office-Suite-Simplify/archive/refs/tags/v2.1.0.zip"
acks-office doctor
```

安装后可用 `acks-office` 命令或 `python -m acks_office`（命令不在 PATH 上时）。

- 需要下载开源中文字体时，额外安装 fonttools：`pip install "acks-office[fonts] @ <同上地址>"`
- 系统 Python 受包管理器保护（提示 externally-managed-environment）时，请使用虚拟环境：

  ```bash
  python3 -m venv ~/.venvs/acks-office
  ~/.venvs/acks-office/bin/python -m pip install "acks-office @ <同上地址>"
  ```

## 方式三：源码

```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
python -m pip install -e ".[dev]"
```

## 验证

```bash
acks-office doctor          # 查看能力等级与补齐建议
python -m pytest            # 源码安装时：单元测试
python test_integration.py  # 源码安装时：集成测试
```

能力等级：`none` 缺依赖；`L0` 可生成全部格式；`L1` 字体齐全、排版与设计一致；`L2` 另可格式转换。

## 可选组件

### 中文字体（PDF）

macOS、Windows 自带可嵌入的中文字体。Linux 可安装文泉驿，或下载开源的 Noto Sans SC：

```bash
sudo apt-get install -y fonts-wqy-microhei     # Ubuntu / Debian
acks-office fonts install noto-sans-sc          # 或：下载 Noto Sans SC 到用户数据目录（需 fonttools）
```

### LibreOffice（格式转换）

```bash
brew install --cask libreoffice                                # macOS
winget install -e --id TheDocumentFoundation.LibreOffice       # Windows
sudo apt-get install -y libreoffice                            # Ubuntu / Debian
sudo dnf install -y libreoffice                                # Fedora / OpenCloudOS
```

装在非标准位置时，用环境变量 `ACKS_OFFICE_SOFFICE` 指定 soffice 的路径。

## 常见问题

**PDF 中文显示为方块或被替换**：运行 `acks-office doctor` 查看 `fonts.pdf_cjk`；`embedded` 为 false 时按上文安装字体，
或用 `--font` / `ACKS_OFFICE_PDF_FONT` 指定一个 TrueType 中文字体。

**转换报 ENGINE_UNAVAILABLE**：没有找到 LibreOffice，按上文安装或设置 `ACKS_OFFICE_SOFFICE`。

**依赖下载慢**：可使用国内镜像，例如 `pip install -i https://pypi.tuna.tsinghua.edu.cn/simple …`；
字体下载可设置 `ACKS_OFFICE_DOWNLOAD_MIRROR` 为 raw.githubusercontent.com 的镜像地址。

**从 2.0 升级**：`import office_suite` 仍可用，建议改为 `import acks_office`；行为变化见 [CHANGELOG](CHANGELOG.md)。
