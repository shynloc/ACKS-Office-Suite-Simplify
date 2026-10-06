"""技能包：SKILL.md 符合 Agent Skills 规范、入口脚本可用、版本号一致、打包内容正确。"""

import importlib.util
import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

from acks_office import __version__, utils

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills" / "acks-office"
yaml = pytest.importorskip("yaml")


def _frontmatter():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    match = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    assert match, "SKILL.md 需要以 YAML frontmatter 开头"
    return yaml.safe_load(match.group(1)), text[match.end():]


def _load_wrapper():
    spec = importlib.util.spec_from_file_location("acks_wrapper", SKILL / "scripts" / "acks.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _run(script, *argv, env=None, flags=()):
    env = dict(os.environ, **(env or {}))
    env.pop("PYTHONPATH", None)
    out = subprocess.run([sys.executable, *flags, str(script), *argv, "--json"],
                         capture_output=True, text=True, encoding="utf-8", env=env)
    return out.returncode, json.loads(out.stdout)


def test_frontmatter_follows_agent_skills_spec():
    meta, body = _frontmatter()
    assert set(meta) <= {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
    assert meta["name"] == SKILL.name
    assert re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", meta["name"]) and len(meta["name"]) <= 64
    assert 0 < len(meta["description"]) <= 1024
    assert len(meta["compatibility"]) <= 500
    assert all(isinstance(k, str) and isinstance(v, str) for k, v in meta["metadata"].items())
    assert meta["metadata"]["version"] == __version__
    assert len(body.splitlines()) < 500


def test_referenced_files_exist():
    _, body = _frontmatter()
    links = re.findall(r"\]\(((?:references|scripts)/[^)]+)\)", body)
    assert links and all((SKILL / link).is_file() for link in links)


def test_wrapper_matches_package():
    wrapper = _load_wrapper()
    assert wrapper.VERSION == __version__
    assert wrapper.data_dir() == utils.data_dir()
    assert f"v{__version__}.zip" in wrapper.ARCHIVE


def test_wrapper_runs_repo_code(tmp_path):
    code, env = _run(SKILL / "scripts" / "acks.py", "version", env={"ACKS_OFFICE_HOME": str(tmp_path)})
    assert code == 0 and env["data"]["version"] == __version__


def test_wrapper_without_package_prints_install_plan(tmp_path):
    skill = tmp_path / "skills" / "acks-office" / "scripts"
    skill.mkdir(parents=True)
    script = skill / "acks.py"
    script.write_bytes((SKILL / "scripts" / "acks.py").read_bytes())

    # -S：不加载 site-packages，模拟只拷贝了技能目录、还没装程序本体的环境
    code, env = _run(script, "doctor", env={"ACKS_OFFICE_HOME": str(tmp_path / "home")}, flags=("-S",))

    assert code == 1 and env["error"]["code"] == "PACKAGE_MISSING"
    (item,) = env["data"]["plan"]
    venv = tmp_path / "home" / "venv"
    assert item["needs_consent"] and item["target"] == str(venv)
    assert item["steps"][0] == [sys.executable, "-m", "venv", str(venv)]
    assert item["steps"][1][-1].endswith(f"v{__version__}.zip")


def test_changelog_mentions_this_version():
    assert re.search(rf"^## \[?{re.escape(__version__)}\]?", (ROOT / "CHANGELOG.md").read_text(encoding="utf-8"), re.M)


def test_bundle_contains_skill_and_code(tmp_path):
    out = subprocess.run([sys.executable, str(ROOT / "tools" / "build_skill_bundle.py"), "--out", str(tmp_path)],
                         capture_output=True, text=True, encoding="utf-8")
    assert out.returncode == 0, out.stderr
    bundle = tmp_path / f"acks-office-skill-{__version__}.zip"
    names = zipfile.ZipFile(bundle).namelist()
    for required in ("acks-office/SKILL.md", "acks-office/LICENSE", "acks-office/scripts/acks.py",
                     "acks-office/references/commands.md", "acks-office/lib/acks_office/__init__.py",
                     "acks-office/lib/acks_office/design_system/tokens.json"):
        assert required in names
    assert not [n for n in names if "__pycache__" in n or "/tests/" in n or n.endswith(".pyc")]
    assert not [n for n in names if n.startswith("acks-office/lib/office_suite")]

    zipfile.ZipFile(bundle).extractall(tmp_path / "unpacked")
    script = tmp_path / "unpacked" / "acks-office" / "scripts" / "acks.py"
    # 只靠技能包自带的代码（-S 不加载已安装的包）：版本命令可用，doctor 报告缺依赖并给出专用环境的安装步骤
    home = {"ACKS_OFFICE_HOME": str(tmp_path / "home")}
    code, env = _run(script, "version", env=home, flags=("-S",))
    assert code == 0 and env["data"]["version"] == __version__
    code, env = _run(script, "doctor", "--no-network", "--no-probe", env=home, flags=("-S",))
    assert code == 0 and env["data"]["level"] == "none"
    (item,) = [i for i in env["data"]["plan"] if i["id"] == "install_dependencies"]
    assert item["target"] == str(tmp_path / "home" / "venv")
    assert Path(env["data"]["paths"]["skill_dir"]).resolve() == script.parent.parent.resolve()
