#!/usr/bin/env python3
"""acks-office 技能入口：找到合适的 Python 和 acks_office 代码，再交给命令行。

用法：python3 scripts/acks.py <命令> [参数] --json（Windows 上用 python 或 py -3）

1. 已建好专用虚拟环境（<数据目录>/venv）时，改用它运行；
2. 优先用技能自带的代码（发布包里的 lib/，或仓库里的 src/），其次用已安装的 acks-office 包；
3. 都没有时输出安装计划，由 Agent 征得用户同意后执行。

本文件只用标准库，并保持 Python 3.6 也能运行，这样环境不满足时也能给出明确提示。
"""
import json
import os
import subprocess
import sys
from pathlib import Path

VERSION = "2.1.0"
MIN_PYTHON = (3, 9)
SKILL_DIR = Path(__file__).resolve().parent.parent
ARCHIVE = "https://github.com/shynloc/ACKS-Office-Suite-Simplify/archive/refs/tags/v{}.zip".format(VERSION)


def data_dir():
    """与 acks_office.utils.data_dir 保持一致。"""
    override = os.environ.get("ACKS_OFFICE_HOME")
    if override:
        return Path(override).expanduser()
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "acks-office"
    if os.name == "nt":
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "acks-office"
    return Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "acks-office"


def venv_python(venv):
    return venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def report(argv, code, message, hint, plan):
    if "--json" in argv:
        env = {"ok": False, "data": {"plan": plan}, "artifacts": [], "warnings": [],
               "error": {"code": code, "message": message, "hint": hint}}
        ascii_only = os.name == "nt" and not sys.stdout.isatty()
        print(json.dumps(env, ensure_ascii=ascii_only, indent=2))
    else:
        print("错误：" + message)
        print("提示：" + hint)
        for item in plan:
            for step in item.get("steps", []):
                print("  " + " ".join('"{}"'.format(a) if " " in a else a for a in step))
    return 1


def main(argv):
    os.environ.setdefault("ACKS_OFFICE_SKILL_DIR", str(SKILL_DIR))
    venv = data_dir() / "venv"
    py = venv_python(venv)
    in_venv = Path(sys.prefix).resolve() == venv.resolve()
    if py.is_file() and not in_venv and not os.environ.get("ACKS_OFFICE_NO_VENV"):
        try:
            return subprocess.call([str(py), str(Path(__file__).resolve())] + argv)
        except OSError:
            pass  # 虚拟环境损坏（例如创建它的 Python 已被卸载）时，用当前 Python 继续

    if sys.version_info < MIN_PYTHON:
        return report(argv, "PYTHON_TOO_OLD",
                      "当前 Python 是 {}，需要 {}.{} 或更高版本".format(sys.version.split()[0], *MIN_PYTHON),
                      "请用户安装新版 Python 后，用新的 Python 运行本脚本",
                      [{"id": "upgrade_python", "links": {"python": "https://www.python.org/downloads/"},
                        "needs_consent": True}])

    for src in (SKILL_DIR / "lib", SKILL_DIR.parent.parent / "src"):
        if (src / "acks_office" / "__init__.py").is_file():
            sys.path.insert(0, str(src))
            break
    try:
        from acks_office.cli import main as cli_main
    except ImportError:
        install = [str(py), "-m", "pip", "install", ARCHIVE]
        steps = [install] if in_venv else [[sys.executable, "-m", "venv", str(venv)], install]
        return report(argv, "PACKAGE_MISSING", "没有找到 acks-office 程序代码",
                      "征得用户同意后，按 data.plan 的步骤安装到专用虚拟环境（不影响系统 Python），然后重新运行",
                      [{"id": "install_package", "target": str(venv), "steps": steps,
                        "network": True, "installs": True, "needs_consent": True}])
    return cli_main(argv)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
