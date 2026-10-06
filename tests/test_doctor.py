"""doctor：报告结构、宿主识别、补齐计划；只读不写。"""

import re
import sys
from pathlib import Path

import pytest

from acks_office import __version__, doctor

ROOT = Path(__file__).resolve().parent.parent
HOST_ENV = re.compile(r"^(CLAUDECODE|AI_AGENT|CODEX_|OPENCLAW_|HERMES_|WORKBUDDY_|ACKS_OFFICE_SKILL_DIR)")


@pytest.fixture
def clean_env(monkeypatch, tmp_path):
    """去掉本机 Agent 留下的环境变量，数据目录指向临时目录。"""
    import os
    for key in list(os.environ):
        if HOST_ENV.match(key):
            monkeypatch.delenv(key)
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    return tmp_path / "home"


def test_report_shape(clean_env):
    report = doctor.run(network=False, probe=False)

    assert report["schema"] == "doctor/1" and report["version"] == __version__
    assert report["level"] in ("none", "L0", "L1", "L2")
    assert {"host", "system", "python", "dependencies", "fonts", "office_apps", "agents",
            "network", "paths", "plan"} <= set(report)
    assert report["network"] == {"pypi": None, "github": None}
    assert report["office_apps"][0]["name"] == "libreoffice"
    assert report["paths"]["data_dir"] == str(clean_env)
    assert all(d["ok"] for d in report["dependencies"])  # 测试环境依赖齐全
    for item in report["plan"]:
        assert item["needs_consent"] is True and item["why"]
        for step in item.get("steps", []):
            assert step and all(isinstance(arg, str) for arg in step)


def test_doctor_does_not_create_anything(clean_env):
    doctor.run(network=False, probe=False)
    assert not clean_env.exists()


@pytest.mark.parametrize("skill_dir, host", [
    ("~/.openclaw/workspace/skills/acks-office", "openclaw"),
    ("~/.workbuddy/skills/acks-office", "workbuddy"),
    ("~/.claude/skills/acks-office", "claude-code"),
    ("/mnt/skills/user/acks-office", "claude-cloud"),
])
def test_host_from_skill_location(clean_env, skill_dir, host):
    found = doctor.run(skill_dir=skill_dir, network=False, probe=False)["host"]
    assert found["guess"] == host and found["evidence"][0] == f"skill_dir={skill_dir}"


def test_host_from_environment(clean_env, monkeypatch):
    monkeypatch.setenv("CODEX_HOME", "/somewhere")
    assert doctor._host(None) == {"guess": "codex", "evidence": ["env:CODEX_HOME"]}


def test_dependency_list_matches_pyproject():
    text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    block = re.search(r"^dependencies = \[(.*?)^\]", text, re.S | re.M).group(1)
    declared = re.findall(r'"([^"]+)"', block)
    assert declared == [f"{dist}{spec}" for dist, _, spec in doctor.DEPENDENCIES]


def _missing(monkeypatch, *names):
    real = doctor._dependencies

    def fake():
        return [dict(d, ok=d["name"] not in names) for d in real()]
    monkeypatch.setattr(doctor, "_dependencies", fake)


def test_skill_mode_installs_into_private_venv(clean_env, monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_SKILL_DIR", str(tmp_path / "skill"))
    _missing(monkeypatch, "reportlab", "pypdf")

    report = doctor.run(network=False, probe=False)

    assert report["level"] == "none"
    (item,) = [i for i in report["plan"] if i["id"] == "install_dependencies"]
    venv = clean_env / "venv"
    assert item["target"] == str(venv)
    assert item["steps"][0] == [sys.executable, "-m", "venv", str(venv)]
    # 新建的虚拟环境里什么都没有：装全部依赖，而不只是当前 Python 缺的两个
    assert item["steps"][1][1:] == ["-m", "pip", "install", *(f"{d}{s}" for d, _, s in doctor.DEPENDENCIES)]
    assert Path(item["steps"][1][0]).parent.parent == venv


def test_cli_mode_installs_into_current_python(clean_env, monkeypatch):
    _missing(monkeypatch, "xlrd")
    (item,) = [i for i in doctor.run(network=False, probe=False)["plan"] if i["id"] == "install_dependencies"]
    assert item["steps"] == [[sys.executable, "-m", "pip", "install", "xlrd>=2.0.1"]]


def test_font_plan_when_nothing_embeddable(clean_env, monkeypatch):
    from acks_office import fonts
    monkeypatch.setattr(fonts, "_installed_catalog_fonts", lambda: [])
    monkeypatch.setattr(fonts, "_system_cjk_candidates", lambda: [])
    monkeypatch.delenv("ACKS_OFFICE_PDF_FONT", raising=False)

    report = doctor.run(network=False, probe=False)

    assert report["level"] == "L0" and report["fonts"]["pdf_cjk"]["source"] == "builtin"
    (item,) = [i for i in report["plan"] if i["id"] == "install_cjk_font"]
    assert item["steps"][-1] == [sys.executable, "-m", "acks_office", "fonts", "install", "noto-sans-sc"]


def test_font_tools_go_into_the_new_venv(clean_env, monkeypatch, tmp_path):
    from acks_office import fonts
    monkeypatch.setattr(fonts, "_installed_catalog_fonts", lambda: [])
    monkeypatch.setattr(fonts, "_system_cjk_candidates", lambda: [])
    monkeypatch.delenv("ACKS_OFFICE_PDF_FONT", raising=False)
    monkeypatch.setenv("ACKS_OFFICE_SKILL_DIR", str(tmp_path / "skill"))
    _missing(monkeypatch, "python-docx")

    (item,) = [i for i in doctor.run(network=False, probe=False)["plan"] if i["id"] == "install_cjk_font"]
    assert Path(item["steps"][0][0]).parent.parent == clean_env / "venv"
    assert item["steps"][0][1:] == ["-m", "pip", "install", doctor.FONTTOOLS]


def test_libreoffice_plan_is_optional(clean_env, monkeypatch):
    monkeypatch.setattr(doctor, "find_soffice", lambda: None)
    report = doctor.run(network=False, probe=True)
    (item,) = [i for i in report["plan"] if i["id"] == "install_libreoffice"]
    assert item["optional"] is True and item["links"]["libreoffice"].startswith("https://")
    assert report["level"] in ("none", "L0", "L1")
