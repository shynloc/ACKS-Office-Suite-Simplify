"""环境探测（只读）：宿主线索、能力等级、缺项与补齐计划。不安装、不修改任何东西。

只依赖标准库，所以在 python-docx、reportlab 等依赖还没装时也能运行，并把缺什么报告出来。
"""

import importlib.metadata as metadata
import importlib.util
import os
import platform
import shutil
import subprocess
import sys
import sysconfig
import urllib.error
import urllib.request
from pathlib import Path
from typing import Dict, List, Optional

from . import __version__, fonts
from .utils import data_dir, find_soffice

SCHEMA = "doctor/1"
MIN_PYTHON = (3, 9)
# (发行包, 导入名, 版本要求)：与 pyproject.toml 的 dependencies 保持一致（有测试核对）
DEPENDENCIES = [("python-docx", "docx", ">=1.1.0"), ("openpyxl", "openpyxl", ">=3.1.2"),
                ("python-pptx", "pptx", ">=0.6.23"), ("reportlab", "reportlab", ">=4.0.0"),
                ("pypdf", "pypdf", ">=4.0.0"), ("pandas", "pandas", ">=2.0.0"), ("xlrd", "xlrd", ">=2.0.1"),
                ("Pillow", "PIL", ">=9.0.0"), ("markdown-it-py", "markdown_it", ">=3.0.0")]
FONTTOOLS = "fonttools>=4.40"
LIBREOFFICE_DOWNLOAD = "https://www.libreoffice.org/download/download-libreoffice/"

# 本机常见 Agent 的主目录与技能目录
AGENT_HOMES = [
    ("openclaw", "~/.openclaw", ["~/.openclaw/skills", "~/.openclaw/workspace/skills"]),
    ("workbuddy", "~/.workbuddy", ["~/.workbuddy/skills"]),
    ("hermes", "~/.hermes", ["~/.hermes/skills"]),
    ("claude-code", "~/.claude", ["~/.claude/skills"]),
    ("codex", "~/.codex", ["~/.codex/skills"]),
    ("deepseek-harness", "~/.dsh", []),
    ("doubao-work", "~/DoubaoWork", ["~/DoubaoWork/skills"]),
    ("agents-shared", "~/.agents", ["~/.agents/skills"]),
]
_PATH_HINTS = [("/.openclaw/", "openclaw"), ("/.workbuddy/", "workbuddy"), ("/.hermes/", "hermes"),
               ("/.claude/", "claude-code"), ("/.codex/", "codex"), ("/.dsh/", "deepseek-harness"),
               ("/dsh/", "deepseek-harness"), ("/doubaowork/", "doubao-work"),
               ("/mnt/skills/", "claude-cloud"), ("/mnt/user-data/", "claude-cloud"), ("/home/oai/", "chatgpt")]
_ENV_PREFIXES = [("CODEX_", "codex"), ("OPENCLAW_", "openclaw"), ("HERMES_", "hermes"), ("WORKBUDDY_", "workbuddy")]


def _host(skill_dir: Optional[str]) -> Dict:
    guesses, evidence = [], []
    for path in filter(None, [skill_dir, os.environ.get("ACKS_OFFICE_SKILL_DIR")]):
        low = str(Path(path).expanduser()).replace("\\", "/").lower() + "/"
        for marker, name in _PATH_HINTS:
            if marker in low:
                guesses.append(name)
                evidence.append(f"skill_dir={path}")
                break
    exe = sys.executable.replace("\\", "/").lower()
    for marker, name in _PATH_HINTS:
        if marker in exe:
            guesses.append(name)
            evidence.append(f"python={sys.executable}")
            break
    for key in sorted(os.environ):
        if key == "CLAUDECODE":
            guesses.append("claude-code")
            evidence.append("env:CLAUDECODE")
        for prefix, name in _ENV_PREFIXES:
            if key.startswith(prefix):
                guesses.append(name)
                evidence.append(f"env:{key}")
    if os.environ.get("AI_AGENT"):
        evidence.append(f"env:AI_AGENT={os.environ['AI_AGENT']}")
    return {"guess": guesses[0] if guesses else None, "evidence": list(dict.fromkeys(evidence))}


def _system() -> Dict:
    managers = [m for m in ("brew", "winget", "choco", "scoop", "apt", "dnf", "yum", "pacman", "uv", "pipx")
                if shutil.which(m)]
    try:
        admin = os.geteuid() == 0 if hasattr(os, "geteuid") else bool(
            __import__("ctypes").windll.shell32.IsUserAnAdmin())
    except Exception:
        admin = None
    return {"os": platform.system(), "release": platform.release(), "machine": platform.machine(),
            "package_managers": managers, "is_admin": admin}


def _dependencies() -> List[Dict]:
    deps = []
    for dist, module, spec in DEPENDENCIES:
        try:
            version = metadata.version(dist)
        except metadata.PackageNotFoundError:
            version = None
        deps.append({"name": dist, "requires": spec, "version": version,
                     "ok": importlib.util.find_spec(module) is not None})
    return deps


def _externally_managed() -> bool:
    """PEP 668：系统 Python 受包管理器保护时，pip 不能直接往里装（虚拟环境里不受影响）。"""
    stdlib = sysconfig.get_paths().get("stdlib", "")
    return (sys.prefix == sys.base_prefix and bool(stdlib)
            and os.path.isfile(os.path.join(stdlib, "EXTERNALLY-MANAGED")))


def _python(python_ok: bool) -> Dict:
    return {"version": platform.python_version(), "executable": sys.executable, "ok": python_ok,
            "min": f"{MIN_PYTHON[0]}.{MIN_PYTHON[1]}", "in_venv": sys.prefix != sys.base_prefix,
            "externally_managed": _externally_managed()}


def runtime_venv() -> Path:
    """技能专用的 Python 虚拟环境；技能入口 scripts/acks.py 发现它后会自动改用。"""
    return data_dir() / "venv"


def _venv_python(venv: Path) -> Path:
    return venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def _self_command() -> List[str]:
    """调用本工具的方式：经技能入口运行时用入口脚本，否则用模块方式（不依赖 PATH）。"""
    skill = os.environ.get("ACKS_OFFICE_SKILL_DIR")
    if skill and (Path(skill) / "scripts" / "acks.py").is_file():
        return [sys.executable, str(Path(skill) / "scripts" / "acks.py")]
    return [sys.executable, "-m", "acks_office"]


def _win_app_path(exe: str) -> Optional[str]:
    try:
        import winreg
    except ImportError:
        return None
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}") as key:
                return winreg.QueryValue(key, None)
        except OSError:
            continue
    return None


def _other_office_apps() -> List[Dict]:
    """Office、WPS、iWork：只报告是否安装，2.x 不调用这些程序。"""
    if sys.platform == "darwin":
        found = {
            "microsoft-word": ["/Applications/Microsoft Word.app"],
            "microsoft-powerpoint": ["/Applications/Microsoft PowerPoint.app"],
            "microsoft-excel": ["/Applications/Microsoft Excel.app"],
            "wps": ["/Applications/wpsoffice.app"],
            "keynote": ["/Applications/Keynote.app"],
            "pages": ["/Applications/Pages.app"],
            "numbers": ["/Applications/Numbers.app"],
        }
    elif os.name == "nt":
        local = os.environ.get("LOCALAPPDATA", "")
        found = {
            "microsoft-word": [_win_app_path("WINWORD.EXE")],
            "microsoft-powerpoint": [_win_app_path("POWERPNT.EXE")],
            "microsoft-excel": [_win_app_path("EXCEL.EXE")],
            "wps": [os.path.join(local, "Kingsoft", "WPS Office", "ksolaunch.exe")],
        }
    else:
        found = {"wps": ["/usr/bin/wps", "/opt/kingsoft/wps-office/office6/wps"]}
    apps = []
    for name, paths in found.items():
        path = next((p for p in paths if p and os.path.exists(p)), None)
        apps.append({"name": name, "path": path, "found": bool(path),
                     "callable": "optional_engine_disabled" if path else False})
    return apps


def _libreoffice(probe: bool) -> Dict:
    path = find_soffice()
    entry = {"name": "libreoffice", "path": path, "found": bool(path), "callable": False, "version": None}
    if path and probe:
        try:
            out = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=60)
            entry["callable"] = out.returncode == 0
            entry["version"] = (out.stdout or "").strip().splitlines()[0] if out.stdout.strip() else None
        except Exception as exc:
            entry["error"] = str(exc)
    return entry


def _agents() -> List[Dict]:
    agents = []
    for name, home, skills in AGENT_HOMES:
        home_path = Path(home).expanduser()
        if home_path.is_dir():
            dirs = [str(Path(d).expanduser()) for d in skills if Path(d).expanduser().is_dir()]
            agents.append({"name": name, "home": str(home_path), "skills_dirs": dirs})
    return agents


# 依赖从 PyPI 安装；开源字体从 GitHub 下载（不通时可设 ACKS_OFFICE_DOWNLOAD_MIRROR）
NETWORK_CHECKS = {"pypi": "https://pypi.org/simple/pip/", "github": "https://raw.githubusercontent.com/"}


def _reachable(url: str, timeout: float = 3.0) -> bool:
    try:
        urllib.request.urlopen(urllib.request.Request(url, method="HEAD"), timeout=timeout)
        return True
    except urllib.error.HTTPError:
        return True  # 能收到 HTTP 响应即说明网络可达
    except Exception:
        return False


def _writable(path: Path) -> bool:
    for candidate in [path, *path.parents]:
        if candidate.exists():
            return os.access(candidate, os.W_OK)
    return False


def _google_fonts_link(family: str) -> str:
    return "https://fonts.google.com/specimen/" + family.replace(" ", "+")


def _outside_runtime() -> bool:
    """经技能入口运行，但还不在专用虚拟环境里。"""
    return (bool(os.environ.get("ACKS_OFFICE_SKILL_DIR"))
            and Path(sys.prefix).resolve() != runtime_venv().resolve())


def _install_runtime(missing: List[str]) -> Dict:
    """缺依赖时的安装步骤。经技能运行时装进专用虚拟环境，不碰系统 Python；
    直接使用已安装的命令行时，装进当前 Python。"""
    venv = runtime_venv()
    if _outside_runtime():
        # 新建的虚拟环境看不到当前 Python 里已装的包，所以装全部依赖
        specs = [f"{dist}{spec}" for dist, _, spec in DEPENDENCIES]
        py = str(_venv_python(venv))
        item = {"target": str(venv),
                "steps": [[sys.executable, "-m", "venv", str(venv)], [py, "-m", "pip", "install", *specs]]}
        if shutil.which("uv"):  # 有些 Linux 发行版的 Python 没带 venv 模块
            item["alternative_steps"] = [["uv", "venv", "--python", sys.executable, str(venv)],
                                         ["uv", "pip", "install", "--python", py, *specs]]
        return item
    specs = [f"{dist}{spec}" for dist, _, spec in DEPENDENCIES if dist in missing]
    item = {"target": sys.prefix, "steps": [[sys.executable, "-m", "pip", "install", *specs]]}
    if _externally_managed():
        item["note"] = "当前 Python 受系统包管理器保护（PEP 668），请改用虚拟环境或 pipx 安装 acks-office"
    return item


def _libreoffice_steps(managers: List[str]) -> List[List[str]]:
    if sys.platform == "darwin":
        return [["brew", "install", "--cask", "libreoffice"]] if "brew" in managers else []
    if os.name == "nt":
        if "winget" in managers:
            return [["winget", "install", "-e", "--id", "TheDocumentFoundation.LibreOffice"]]
        if "choco" in managers:
            return [["choco", "install", "libreoffice-fresh", "-y"]]
        if "scoop" in managers:
            return [["scoop", "bucket", "add", "extras"], ["scoop", "install", "libreoffice"]]
        return []
    for manager, cmd in (("apt", ["sudo", "apt-get", "install", "-y", "libreoffice"]),
                         ("dnf", ["sudo", "dnf", "install", "-y", "libreoffice"]),
                         ("pacman", ["sudo", "pacman", "-S", "--noconfirm", "libreoffice-fresh"])):
        if manager in managers:
            return [cmd]
    return []


def _plan(python_ok: bool, deps: List[Dict], pdf_font: Dict, theme: Dict, libreoffice: Dict) -> List[Dict]:
    """补齐计划：每一项都需要用户同意后才能执行；steps 是可直接执行的参数列表，不经过 shell。"""
    plan: List[Dict] = []
    if not python_ok:
        plan.append({"id": "upgrade_python", "why": f"需要 Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} 或更高版本",
                     "links": {"python": "https://www.python.org/downloads/"},
                     "network": True, "installs": True, "needs_consent": True})
    missing = [d["name"] for d in deps if not d["ok"]]
    if missing and python_ok:
        plan.append({"id": "install_dependencies", "why": f"缺少依赖：{', '.join(missing)}",
                     **_install_runtime(missing), "network": True, "installs": True, "needs_consent": True})
    if not pdf_font.get("embedded"):
        steps = [_self_command() + ["fonts", "install", "noto-sans-sc"]]
        if missing and _outside_runtime():  # 字体将在新建的专用虚拟环境里安装
            steps.insert(0, [str(_venv_python(runtime_venv())), "-m", "pip", "install", FONTTOOLS])
        elif importlib.util.find_spec("fontTools") is None:
            steps.insert(0, [sys.executable, "-m", "pip", "install", FONTTOOLS])
        plan.append({"id": "install_cjk_font", "why": "没有可嵌入 PDF 的中文字体，PDF 中文将由阅读器替换显示",
                     "steps": steps, "license": "SIL Open Font License 1.1（可免费商用）",
                     "network": True, "installs": True, "needs_consent": True})
    if theme["missing"]:
        plan.append({"id": "install_theme_fonts",
                     "why": f"缺少主题字体：{', '.join(theme['missing'])}；Word/PPT 打开时会被替换成其他字体",
                     "links": {family: _google_fonts_link(family) for family in theme["missing"]},
                     "note": "需要安装到系统或用户字体目录，请先征得用户同意",
                     "network": True, "installs": True, "needs_consent": True})
    if libreoffice["callable"] is not True:
        steps = _libreoffice_steps(_system()["package_managers"])
        plan.append({"id": "install_libreoffice", "optional": True,
                     "why": "转 PDF、渲染预览与公式重算需要 LibreOffice；不装也能生成全部格式",
                     "steps": steps, "links": {"libreoffice": LIBREOFFICE_DOWNLOAD},
                     "needs_admin": any(step[0] in ("sudo", "winget", "choco") for step in steps),
                     "network": True, "installs": True, "needs_consent": True})
    return plan


def run(skill_dir: Optional[str] = None, network: bool = True, probe: bool = True) -> Dict:
    """生成环境报告。network=False 跳过联网检查；probe=False 不运行 LibreOffice --version。"""
    python_ok = sys.version_info >= MIN_PYTHON
    deps = _dependencies()
    deps_ok = python_ok and all(d["ok"] for d in deps)
    try:
        pdf_font = fonts.pdf_fonts().describe()
    except Exception as exc:
        pdf_font = {"error": str(exc), "embedded": False}
    theme = fonts.theme_font_status()
    libreoffice = _libreoffice(probe)
    fonts_ok = bool(pdf_font.get("embedded")) and not theme["missing"]

    if not deps_ok:
        level = "none"
    elif not fonts_ok:
        level = "L0"
    else:
        level = "L2" if libreoffice["callable"] is True else "L1"

    plan = _plan(python_ok, deps, pdf_font, theme, libreoffice)
    home = data_dir()
    return {
        "schema": SCHEMA,
        "version": __version__,
        "host": _host(skill_dir),
        "level": level,
        "system": _system(),
        "python": _python(python_ok),
        "dependencies": deps,
        "fonts": {"pdf_cjk": pdf_font, "theme": theme},
        "office_apps": [libreoffice] + _other_office_apps(),
        "agents": _agents(),
        "network": {name: _reachable(url) if network else None for name, url in NETWORK_CHECKS.items()},
        "paths": {"data_dir": str(home), "writable": _writable(home),
                  "skill_dir": skill_dir or os.environ.get("ACKS_OFFICE_SKILL_DIR")},
        "plan": plan,
    }
