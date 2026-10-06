# doctor：环境报告

`doctor --json` 只读：不安装、不下载、不修改文件。`data` 的结构（`schema: "doctor/1"`）：

| 字段 | 说明 |
|---|---|
| `version` | acks-office 版本 |
| `host` | `guess`：推测的宿主 Agent（如 `claude-code`、`openclaw`、`workbuddy`、`hermes`、`codex`），`evidence`：依据（技能所在路径、Python 路径、环境变量） |
| `level` | 能力等级，见下表 |
| `system` | 操作系统、架构、可用的包管理器（brew、winget、apt、uv…）、是否管理员 |
| `python` | 版本、路径、是否满足最低版本、是否在虚拟环境、是否受系统保护（PEP 668） |
| `dependencies` | 每个依赖的版本要求、已装版本、能否导入 |
| `fonts.pdf_cjk` | PDF 将使用的中文字体：`family`、`embedded`（是否嵌入 PDF）、`source`（`argument` / `env` / `installed` / `system` / `builtin`） |
| `fonts.theme` | 主题字体 `present` / `missing` |
| `office_apps` | 第一项是 LibreOffice（`path`、`callable`、`version`）；其后是 Word、PowerPoint、Excel、WPS、Keynote 等，只报告是否安装，`callable` 为 `optional_engine_disabled`（2.x 不调用它们） |
| `agents` | 本机发现的 Agent 及其技能目录 |
| `network` | `pypi`、`github` 是否可达（`--no-network` 时为 null） |
| `paths` | `data_dir`：用户数据目录（字体、专用虚拟环境），`writable`：能否写入；`skill_dir`：技能所在目录（经技能入口运行时） |
| `plan` | 补齐计划，见下文 |

## 能力等级

| 等级 | 条件 | 能做什么 |
|---|---|---|
| `none` | Python 版本不够或缺依赖 | 只能运行 doctor |
| `L0` | 依赖齐全 | 生成、读取全部格式；PDF 中文可能未嵌入，或 Word/PPT 主题字体缺失 |
| `L1` | L0 + PDF 中文字体可嵌入 + 主题字体齐全 | 满足字体检测条件；最终显示仍由打开文件的软件决定 |
| `L2` | L1 + LibreOffice 可调用 | 另可格式转换（如 docx → pdf） |

## 补齐计划

每项含 `id`、`why`，以及：

- `steps`：可直接执行的参数列表（不经过 shell，跨 bash / PowerShell / cmd 都一样）；
- `alternative_steps`：首选步骤失败时的备选（例如系统 Python 不带 venv 模块时改用 uv）；
- `target`：安装到哪里；`links`：需要用户手动下载的地址；`note`：补充说明；
- `optional`：可选项，只在任务需要时提出；`needs_admin`：需要管理员权限，交给用户自己执行；
- `needs_consent`：总为 true，**执行前必须征得用户同意**。

| id | 何时出现 | 内容 |
|---|---|---|
| `upgrade_python` | Python 低于 3.9 | 请用户自行安装新版 Python |
| `install_dependencies` | 缺依赖 | 经技能运行时：建专用虚拟环境并装依赖，入口脚本之后自动使用它；直接用命令行时：装进当前 Python |
| `install_cjk_font` | PDF 中文字体不能嵌入 | 安装 fonttools（如缺）并下载 Noto Sans SC（SIL OFL 1.1，可免费商用） |
| `install_theme_fonts` | 主题字体缺失 | 给出 Google Fonts 链接，安装到系统需用户同意 |
| `install_libreoffice` | 没有可用的 LibreOffice（可选） | 按系统给出 brew / winget / choco / scoop / apt / dnf / pacman 命令 |

入口脚本找不到程序本体时会直接返回 `error.code = PACKAGE_MISSING`，`data.plan` 里是
`install_package`：建专用虚拟环境，从 GitHub 发布包安装与技能同版本的 acks-office。
