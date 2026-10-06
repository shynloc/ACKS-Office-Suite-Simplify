# 安全说明

文档对应版本：**2.1.2**。

本文说明该版本的数据、网络与文件处理行为；后续版本以对应版本的说明和实现为准。

## 数据与网络

- 文档的生成、读取、合并、水印都在本机完成，不上传任何内容，也不收集使用数据。
- 只有以下情况会联网：
  | 场景 | 访问的地址 | 说明 |
  |---|---|---|
  | `doctor` | `pypi.org`、`raw.githubusercontent.com` | 只发一个 HEAD 请求判断能否连通；`--no-network` 可关闭 |
  | `fonts install` | `raw.githubusercontent.com`（或 `ACKS_OFFICE_DOWNLOAD_MIRROR`） | 地址固定到具体提交，下载后校验 SHA-256，不一致即放弃 |
  | `send_email` | 你配置的 SMTP 服务器 | 见下文 |
  | 工作流 `data_extract` 的 API 数据源 | 你配置的地址 | 仅在工作流里配置了才会访问 |

## 安装与修改

- `doctor` 只读：不安装、不下载、不修改任何文件。它给出的补齐计划每一项都标有 `needs_consent: true`，
  技能说明要求 Agent 先向用户说明并取得同意后才执行。
- 经技能运行时，依赖安装到用户数据目录下的专用虚拟环境，不改动系统 Python；需要管理员权限的步骤
  （如 `sudo apt-get install libreoffice`）标为 `needs_admin`，交由用户自己执行。
- 命令行默认不覆盖已有文件（需加 `--overwrite`）；`watermark` 默认写到新文件。
  Python API 中 `add_watermark`、`add_transition_effects` 不传 `output_path` 时会覆盖原文件，与旧版一致。

## 邮件

- 密码或授权码请通过环境变量 `OFFICE_EMAIL_PASSWORD`（或 `ACKS_OFFICE_EMAIL_PASSWORD`）提供，不要写进代码、
  配置文件或工作流 YAML；本工具不会把密码写到磁盘或日志里。
- 连接时校验服务器证书与主机名；端口 465/994 使用 SSL，其他端口要求 STARTTLS，服务器不支持加密时拒绝登录。
  只有内网自签名证书或确实不支持加密的服务器才应传 `allow_insecure=True`。
- 密送（Bcc）地址只用于投递，不会出现在邮件头里。

## 处理来路不明的文件

- `extract`、`merge`、`watermark` 用 python-docx、pypdf 等纯 Python 库解析文件，不执行文档里的宏或脚本。
- `convert` 调用 LibreOffice 的无界面模式，每次使用独立的临时配置目录。LibreOffice 本身曾出现过解析类漏洞，
  转换不受信任的文件前请保持 LibreOffice 为最新版本。
- Markdown 正文里的原始 HTML 不会被解释，按普通文字输出。

## 环境变量

| 变量 | 作用 |
|---|---|
| `OFFICE_EMAIL_PASSWORD` / `ACKS_OFFICE_EMAIL_PASSWORD` | 邮箱密码或授权码 |
| `ACKS_OFFICE_HOME` | 用户数据目录（字体、专用虚拟环境） |
| `ACKS_OFFICE_SOFFICE` | LibreOffice 可执行文件路径 |
| `ACKS_OFFICE_PDF_FONT` / `ACKS_OFFICE_PDF_FONT_BOLD` | PDF 中文字体文件 |
| `ACKS_OFFICE_DOWNLOAD_MIRROR` | 字体下载镜像 |

本工具不会自动读取 `.env` 文件；`.env.template` 仅作为模板，请在 shell 中设置或由你的程序加载。
填好真实值的文件不要提交到 Git（`.gitignore` 已忽略 `.env`）。

## 报告安全问题

请在仓库的 Security 页面使用 “Report a vulnerability” 私下报告。若该入口不可用，可提交 Issue 说明受影响的
功能，但不要附带可直接利用的细节。
