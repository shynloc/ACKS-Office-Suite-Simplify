"""theme init 新建主题、theme preview 生成样张。"""

import json

import openpyxl
import pytest

from acks_office import themes
from acks_office.cli import main
from acks_office.themes import ThemeError
from acks_office.themes.preview import build_preview
from acks_office.themes.scaffold import init_theme


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    themes.clear_cache()
    yield
    themes.clear_cache()


def test_init_creates_a_usable_theme(tmp_path):
    result = init_theme("qimu", extends="slate", accent="#0b6e4f", brand="栖木咖啡", title="栖木")
    assert result["path"] == str(tmp_path / "home" / "themes" / "qimu") and result["user_dir"]
    assert result["validation"]["ok"] and not any(w["code"] == "MISSING_FIELD" for w in result["validation"]["warnings"])
    theme = themes.load_theme("qimu")
    assert theme.title == "栖木" and theme.chain[:2] == ["qimu", "slate"]
    assert theme.color("accent") == "#0B6E4F" and theme.chrome_values()["brand"] == "栖木咖啡"
    assert theme.get("layout.doc.cover") == "standard"  # 其余沿用 slate

    with pytest.raises(ThemeError) as exc:
        init_theme("qimu")
    assert exc.value.code == "THEME_EXISTS"
    logo = tmp_path / "home" / "themes" / "qimu" / "logo.png"
    logo.write_bytes(b"png")
    init_theme("qimu", extends="folio", force=True)  # --force 只覆盖主题文件
    assert logo.exists() and themes.load_theme("qimu").chain[1] == "folio"


@pytest.mark.parametrize("kwargs, code", [({"name": "_hidden"}, "BAD_THEME_NAME"), ({"name": "Bad Name"}, "BAD_THEME_NAME"),
                                          ({"name": "ok", "accent": "green"}, "BAD_COLOR"),
                                          ({"name": "ok", "extends": "nope"}, "THEME_NOT_FOUND")])
def test_init_rejects_bad_input(kwargs, code):
    with pytest.raises(ThemeError) as exc:
        init_theme(**kwargs)
    assert exc.value.code == code


def test_preview_writes_specimen_and_samples(tmp_path):
    result = build_preview("folio", str(tmp_path / "out"), ("html", "xlsx", "docx"))
    files = result["files"]
    assert set(files) == {"html", "xlsx", "docx"} and result["validation"]["ok"]
    html = (tmp_path / "out" / "folio-specimen.html").read_text(encoding="utf-8")
    assert "Folio 对开" in html and "对比度" in html and "folio-sample.xlsx" in html
    assert "_base" not in html.split("<header>")[1].split("</header>")[0]
    assert "font-size:" in html and "style=\"font-family: '" in html  # 字体栈不打断 style 属性
    assert openpyxl.load_workbook(files["xlsx"]).sheetnames == ["Q3 分区域", "月度明细"]
    with pytest.raises(ValueError):
        build_preview("folio", str(tmp_path / "bad"), ("gif",))


def test_cli_init_and_preview(tmp_path, capsys, monkeypatch):
    monkeypatch.chdir(tmp_path)
    assert main(["theme", "init", "brand-x", "--extends", "neutral", "--accent", "#7A2E8E", "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["data"]["theme"] == "brand-x" and len(env["artifacts"]) == 4
    assert main(["theme", "preview", "brand-x", "--formats", "html,pdf", "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert set(env["data"]["files"]) == {"html", "pdf"} and (tmp_path / "brand-x-preview" / "brand-x-sample.pdf").exists()
    assert main(["theme", "preview", "brand-x", "--formats", "gif", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "INVALID_INPUT"
    assert main(["theme", "init", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "USAGE_ERROR"
