"""主题引擎：查找与继承、取值、字体选择、模板、校验。"""

import json

import pytest

from acks_office import fonts, themes
from acks_office.cli import main
from acks_office.themes import ThemeError, load_theme, plain_template, render_template, validate_theme

BUILTIN = ("neutral", "slate", "folio")


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    """主题与字体缓存都指向临时目录；每个测试重新加载主题。"""
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    monkeypatch.delenv("ACKS_OFFICE_THEMES", raising=False)
    themes.clear_cache()
    yield
    themes.clear_cache()


def write_theme(directory, theme=None, tokens=None, chrome=None, fonts_=None):
    directory.mkdir(parents=True, exist_ok=True)
    for name, data in (("theme", theme), ("tokens", tokens), ("chrome", chrome), ("fonts", fonts_)):
        if data is not None:
            (directory / f"{name}.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return directory


def test_builtin_themes_are_listed_and_neutral_is_default():
    listed = {t["name"]: t for t in themes.list_themes()}
    assert set(BUILTIN) <= set(listed) and "_base" not in listed
    assert all(listed[name]["source"] == "builtin" for name in BUILTIN)
    assert themes.DEFAULT_THEME == "neutral" and load_theme(None).name == "neutral"


@pytest.mark.parametrize("name", BUILTIN)
def test_builtin_themes_validate_without_errors(name):
    report = validate_theme(name, check_installed=False)
    assert report["ok"], report["errors"]
    assert report["contrast"] and all(item["ok"] for item in report["contrast"])
    assert not [w for w in report["warnings"] if w["code"] == "UNKNOWN_FIELD"]


def test_inheritance_merges_tokens_and_chrome():
    slate, folio, base = load_theme("slate"), load_theme("folio"), load_theme("neutral")
    assert slate.chain == ["slate", "_base"]
    assert slate.color("accent") == "#1F5FAE" and base.color("accent") == "#2557D6"
    assert slate.get("type.doc.body") == base.get("type.doc.body")  # 未覆盖的字段来自基础主题
    assert folio.get("layout.doc.cover") == "issue" and slate.get("layout.doc.cover") == "standard"
    assert folio.chrome_get("doc.header.right") == "{chapter}"
    assert folio.label("table") == "表"  # chrome.labels 从基础主题继承


def test_colors_and_values():
    slate = load_theme("slate")
    assert slate.hex("accent") == "1F5FAE" and slate.rgb("accent") == (0x1F, 0x5F, 0xAE)
    assert slate.color("#abcdef") == "#ABCDEF"
    assert slate.chart_colors()[0] == "#1F5FAE"
    assert slate.size("type.doc.h1") == 18.0
    assert slate.get("type.doc.missing", None) is None
    with pytest.raises(ThemeError):
        slate.color("no-such-color")
    with pytest.raises(ThemeError):
        slate.get("type.doc.missing")


def test_user_theme_extends_builtin_and_can_shadow_it(tmp_path):
    user_dir = tmp_path / "home" / "themes"
    write_theme(user_dir / "brand", {"name": "brand", "title": "我的品牌", "extends": "slate"},
                {"color": {"accent": "#0B6E4F"}}, {"brand": "示例品牌"})
    write_theme(user_dir / "slate", {"name": "slate", "title": "改过的 Slate", "extends": "slate"},
                {"color": {"accent": "#123456"}})
    themes.clear_cache()

    brand = load_theme("brand")
    assert brand.source == "user" and brand.chain[:2] == ["brand", "slate"]
    assert brand.color("accent") == "#0B6E4F" and brand.color("ink") == "#1F2933"
    assert brand.chrome_values({"title": "季报"})["brand"] == "示例品牌"

    slate = load_theme("slate")  # 用户的同名主题生效，并继承内置的 slate
    assert slate.source == "user" and slate.color("accent") == "#123456"
    assert slate.get("layout.doc.cover") == "standard"
    shadowed = [t for t in themes.list_themes() if t["name"] == "slate"]
    assert [t["shadowed"] for t in shadowed] == [False, True]


def test_theme_by_path_with_relative_parent(tmp_path):
    parent = write_theme(tmp_path / "pkg" / "parent", {"name": "parent", "extends": "folio"},
                         {"color": {"accent": "#7A2E8E"}})
    child = write_theme(tmp_path / "pkg" / "child", {"name": "child", "extends": "../parent"})
    theme = load_theme(str(child))
    assert theme.source == "path" and theme.chain[:3] == ["child", "parent", "folio"]
    assert theme.color("accent") == "#7A2E8E" and theme.dirs[1] == parent.resolve()
    assert load_theme(str(child / "theme.json")).name == "child"


def test_missing_cyclic_and_broken_themes(tmp_path):
    with pytest.raises(ThemeError) as missing:
        load_theme("no-such-theme")
    assert missing.value.code == "THEME_NOT_FOUND" and "slate" in missing.value.hint
    with pytest.raises(ThemeError) as bad_name:
        load_theme("Bad Name")
    assert bad_name.value.code == "THEME_NOT_FOUND"

    write_theme(tmp_path / "a", {"name": "a", "extends": "../b"})
    write_theme(tmp_path / "b", {"name": "b", "extends": "../a"})
    with pytest.raises(ThemeError) as cycle:
        load_theme(str(tmp_path / "a"))
    assert cycle.value.code == "THEME_INVALID" and "循环" in cycle.value.message

    broken = write_theme(tmp_path / "broken", {"name": "broken"})
    (broken / "tokens.json").write_text("{not json", encoding="utf-8")
    with pytest.raises(ThemeError) as invalid:
        load_theme(str(broken))
    assert invalid.value.code == "THEME_INVALID"


def test_validation_reports_each_problem(tmp_path):
    bad = write_theme(tmp_path / "bad", {"name": "bad", "title": "坏主题", "extends": "slate", "version": "1",
                                         "license": "MIT"},
                      {"color": {"accent": "#EEEEEE", "ink": "red"},
                       "layout": {"doc": {"cover": "fancy"}, "sheet": {"negative": "no-color"}},
                       "type": {"doc": {"h1": "大", "body": 2}}, "spacing": {"x": 1},
                       "font": {"roles": {"body": {"family": "nope"}}}},
                      {"doc": {"footer": {"left": "{brand} · {nope}"}}, "logo": {"light": "assets/logo.png"}})
    report = validate_theme(str(bad), check_installed=False)
    codes = {e["code"] for e in report["errors"]}
    assert not report["ok"]
    assert {"BAD_COLOR", "BAD_VARIANT", "BAD_COLOR_REF", "WRONG_TYPE", "OUT_OF_RANGE", "LOW_CONTRAST",
            "BAD_FONT_ROLE", "BAD_PLACEHOLDER", "ASSET_MISSING"} <= codes
    assert any(w["code"] == "UNKNOWN_FIELD" and w["path"] == "spacing.x" for w in report["warnings"])


def test_contrast_ratio_matches_wcag_examples():
    assert themes.contrast_ratio("#000000", "#FFFFFF") == pytest.approx(21.0)
    assert themes.contrast_ratio("#777777", "#FFFFFF") == pytest.approx(4.48, abs=0.01)


@pytest.mark.parametrize("template, values, expected", [
    ("{brand} · {short_title}", {"brand": "栖木", "short_title": "季报"}, "栖木 · 季报"),
    ("{brand} · {short_title}", {"brand": "", "short_title": "季报"}, "季报"),
    ("{publication} · 第 {issue} 期", {"publication": "季度刊", "issue": ""}, "季度刊"),
    ("{publication} · 第 {issue} 期", {"publication": "季度刊", "issue": "07"}, "季度刊 · 第 07 期"),
    ("第 {page} 页", {}, "第 {page} 页"),
    ("{nothing}", {}, ""),
])
def test_template_drops_empty_segments(template, values, expected):
    assert plain_template(template, values) == expected


def test_template_marks_page_fields():
    assert render_template("{title} · {page}/{pages}", {"title": "季报"}) == [
        ("text", "季报 · "), ("field", "page"), ("text", "/"), ("field", "pages")]


def test_font_choice_uses_installed_candidates(monkeypatch):
    installed = {"PingFang SC", "Helvetica Neue"}
    monkeypatch.setattr(fonts, "family_installed", lambda name: name in installed)
    slate = load_theme("slate", use_cache=False)

    body = slate.font("body")
    assert (body.cn, body.en, body.bold) == ("PingFang SC", "Helvetica Neue", False)
    assert {n["code"] for n in body.substitutions} == {"FONT_SUBSTITUTED"}
    assert "fonts install noto-sans-sc" in body.substitutions[0]["message"]
    assert slate.font("heading").bold

    strict = load_theme("slate", use_cache=False).font("body", policy="theme")
    assert (strict.cn, strict.en) == ("Noto Sans SC", "Source Sans 3")
    assert {n["code"] for n in strict.substitutions} == {"FONT_MISSING"}

    installed.add("Noto Sans SC")
    assert load_theme("slate", use_cache=False).font("body").cn == "Noto Sans SC"


def test_font_faces_are_read_from_files_and_cached(monkeypatch, tmp_path, cjk_font):
    font_dir = tmp_path / "fonts"
    font_dir.mkdir()
    (font_dir / "AcksTest.ttf").write_bytes(open(cjk_font, "rb").read())
    monkeypatch.setattr(fonts, "font_dirs", lambda: [font_dir])
    monkeypatch.setattr(fonts, "_FACES", None)  # 测试结束后恢复真实的字体列表

    faces = fonts.font_faces(refresh=True)
    assert [(f.families, f.weight, f.italic, f.truetype) for f in faces] == [(("AcksTest",), 400, False, True)]
    assert fonts.family_installed("acks test") and fonts.best_face("AcksTest", 700).path.endswith("AcksTest.ttf")
    assert {ord(c) for c in "中文测试报告"} <= fonts.face_codepoints(faces[0])  # 测试字体含这些字
    assert fonts._cache_path().is_file()

    monkeypatch.setattr(fonts, "read_font_faces", lambda path: pytest.fail("应使用缓存"))
    assert fonts.font_faces(refresh=True) == faces


def test_common_hanzi_list():
    chars = fonts.common_hanzi()
    assert len(chars) == 3755 and chars[0] == "啊" and len(set(chars)) == 3755


def test_cli_theme_commands(capsys, tmp_path):
    assert main(["theme", "list", "--json"]) == 0
    listed = json.loads(capsys.readouterr().out)["data"]
    assert listed["default"] == "neutral" and {"neutral", "slate", "folio"} <= {t["name"] for t in listed["themes"]}

    assert main(["theme", "show", "folio", "--json"]) == 0
    shown = json.loads(capsys.readouterr().out)["data"]
    assert shown["variants"]["slide"]["section"] == "dark" and shown["colors"]["accent"] == "#C8321F"

    assert main(["theme", "validate", "slate", "--no-fonts", "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["data"]["ok"] is True

    bad = write_theme(tmp_path / "bad", {"name": "bad", "extends": "slate"}, {"layout": {"doc": {"cover": "x"}}})
    assert main(["theme", "validate", str(bad), "--no-fonts", "--json"]) == 1
    env = json.loads(capsys.readouterr().out)
    assert env["error"]["code"] == "THEME_INVALID" and env["data"]["errors"][0]["code"] == "BAD_VARIANT"

    assert main(["theme", "show", "nope", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "THEME_NOT_FOUND"
    assert main(["theme", "show", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "USAGE_ERROR"
