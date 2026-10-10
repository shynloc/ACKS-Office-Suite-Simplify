"""字体目录与安装：地址固定、哈希校验、可变字体生成静态字重并改名、装到用户字体目录、按主题安装、doctor 计划。"""

import hashlib
import json
import re

import pytest

from acks_office import doctor, fonts, themes
from acks_office.cli import main

pytest.importorskip("fontTools")


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    monkeypatch.setattr(fonts, "_FACES", None)
    themes.clear_cache()
    yield
    themes.clear_cache()


def _glyph():
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0))
    pen.lineTo((0, 500))
    pen.lineTo((400, 500))
    pen.lineTo((400, 0))
    pen.closePath()
    return pen.glyph()


def _font(path, family, variable):
    from fontTools.fontBuilder import FontBuilder
    fb = FontBuilder(1000, isTTF=True)
    fb.setupGlyphOrder([".notdef", "A"])
    fb.setupCharacterMap({0x41: "A"})
    fb.setupGlyf({".notdef": _glyph(), "A": _glyph()})
    fb.setupHorizontalMetrics({".notdef": (500, 0), "A": (500, 0)})
    fb.setupHorizontalHeader(ascent=800, descent=-200)
    fb.setupNameTable({"familyName": family, "styleName": "Regular"})
    fb.setupOS2(sTypoAscender=800, usWinAscent=800, usWinDescent=200)
    fb.setupPost()
    if variable:
        fb.setupFvar(axes=[("wght", 100, 400, 900, "Weight")], instances=[])
    fb.save(str(path))
    return path.read_bytes()


def _git_sha1(data):
    return hashlib.sha1(b"blob %d\x00" % len(data) + data).hexdigest()


@pytest.fixture
def tiny(tmp_path, monkeypatch):
    """一个本地的「字体目录」条目：一份可变字体 + 一份静态字体，用 file:// 地址。"""
    variable = _font(tmp_path / "Tiny[wght].ttf", "Tiny Var", True)
    static = _font(tmp_path / "TinyLight.ttf", "Tiny Var Light", False)
    (tmp_path / "OFL.txt").write_text("SIL Open Font License 1.1", encoding="utf-8")
    entry = {"family": "Tiny Var", "license": "SIL Open Font License 1.1", "file_stem": "TinyVar",
             "license_url": (tmp_path / "OFL.txt").as_uri(),
             "sources": [{"url": (tmp_path / "Tiny[wght].ttf").as_uri(), "size": len(variable),
                          "git_sha1": _git_sha1(variable),
                          "instances": {"Regular": {"wght": 400}, "Bold": {"wght": 700}, "Black": {"wght": 900}}},
                         {"url": (tmp_path / "TinyLight.ttf").as_uri(), "size": len(static), "static": "Light",
                          "sha256": hashlib.sha256(static).hexdigest()}]}
    monkeypatch.setitem(fonts.CATALOG, "tiny", entry)
    monkeypatch.setattr(fonts, "system_font_dir", lambda: tmp_path / "user-fonts")
    monkeypatch.setattr(fonts, "_register_windows_font", lambda path: None)  # 测试不改注册表
    monkeypatch.setattr(fonts.shutil, "which", lambda name: None)  # 不运行 fc-cache
    return entry


def test_catalog_is_pinned_and_covers_theme_fonts():
    for key, spec in fonts.CATALOG.items():
        assert spec["license_url"].startswith("https://")
        for source in spec["sources"]:
            assert re.search(r"/[0-9a-f]{40}/|/v?\d+\.\d+[A-Za-z]?/", source["url"]), source["url"]  # 固定版本
            assert source.get("sha256") or source.get("git_sha1"), key
            assert source["size"] > 0 and (("static" in source) != ("instances" in source)), key
    wanted = set()
    for name in fonts.builtin_themes():
        families = themes.load_theme(name).fonts["families"]
        wanted |= {spec["install"] for spec in families.values() if spec.get("install")}
    assert wanted and wanted <= set(fonts.CATALOG)


def test_install_variable_and_static_fonts(tiny, tmp_path):
    result = fonts.install_font("tiny", system=True)
    installed = {name.rsplit("/", 1)[-1].rsplit("\\", 1)[-1] for name in result["files"]}
    assert installed == {"TinyVar-Regular.ttf", "TinyVar-Bold.ttf", "TinyVar-Black.ttf", "TinyVar-Light.ttf",
                         "TinyVar-OFL.txt"}
    assert fonts.catalog_installed("tiny")
    assert sorted(p.name for p in (tmp_path / "user-fonts").iterdir()) == [
        "TinyVar-Black.ttf", "TinyVar-Bold.ttf", "TinyVar-Light.ttf", "TinyVar-Regular.ttf"]
    home = tmp_path / "home" / "fonts"
    (black,) = fonts.read_font_faces(str(home / "TinyVar-Black.ttf"))
    assert black.families == ("Tiny Var", "Tiny Var Black") and black.weight == 900 and not black.variable
    (bold,) = fonts.read_font_faces(str(home / "TinyVar-Bold.ttf"))
    assert bold.families == ("Tiny Var",) and bold.style == "Bold" and bold.weight == 700
    assert not list(home.glob("*.download"))
    # 装好后按名称能找到对应字重
    assert fonts.best_face("Tiny Var", 900).weight == 900


def test_hash_mismatch_aborts(tiny, monkeypatch):
    broken = dict(tiny, sources=[dict(tiny["sources"][0], git_sha1="0" * 40)])
    monkeypatch.setitem(fonts.CATALOG, "tiny", broken)
    with pytest.raises(RuntimeError, match="校验失败"):
        fonts.install_font("tiny")
    assert not fonts.catalog_installed("tiny")
    with pytest.raises(ValueError):
        fonts.install_font("nope")


def test_theme_font_status(monkeypatch):
    monkeypatch.setattr(fonts, "family_installed", lambda name: name in {"PingFang SC", "Helvetica Neue", "Menlo"})
    slate = fonts.theme_font_status("slate")
    assert slate["theme"] == "slate" and slate["substituted"]["Noto Sans SC"] == "PingFang SC"
    assert {"noto-sans-sc", "source-sans-3", "jetbrains-mono"} <= set(slate["install"])
    neutral = fonts.theme_font_status("neutral")
    assert "PingFang SC" in neutral["present"] and neutral["install"] == []


def test_cli_install_theme_fonts(monkeypatch, capsys):
    monkeypatch.setattr(fonts, "family_installed", lambda name: False)
    calls = []

    def fake_install(key, progress=None, system=False):
        calls.append((key, system))
        return {"key": key, "family": key, "license": "OFL", "files": []}

    monkeypatch.setattr(fonts, "install_font", fake_install)
    assert main(["fonts", "install", "--theme", "slate", "--system", "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["data"]["theme"] == "slate" and {key for key, _ in calls} == set(fonts.theme_font_status("slate")["install"])
    assert all(system for _, system in calls)
    assert main(["fonts", "install", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "INVALID_INPUT"
    assert main(["fonts", "install", "--theme", "nope", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "THEME_NOT_FOUND"


def test_doctor_reports_theme_fonts(monkeypatch):
    monkeypatch.setattr(fonts, "family_installed", lambda name: False)
    report = doctor.run(network=False, probe=False)
    assert set(report["fonts"]["themes"]) >= {"neutral", "slate", "folio"}
    (item,) = [i for i in report["plan"] if i["id"] == "install_theme_fonts"]
    assert item["optional"] and item["needs_consent"]
    assert any(step[-4:] == ["install", "--theme", "folio", "--system"] for step in item["steps"])
