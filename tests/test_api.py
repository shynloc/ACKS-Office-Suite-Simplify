"""3.0 的函数式接口：出错抛异常；写文件都写到新文件；OfficeSuite 保留到 4.0 并提示弃用。"""

import pytest

import acks_office
from acks_office import themes
from acks_office.themes import ThemeError


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    themes.clear_cache()
    yield
    themes.clear_cache()


def test_create_extract_and_watermark_never_overwrite(tmp_path):
    word = tmp_path / "r.docx"
    assert acks_office.create("word", str(word), content="# 标题\n\n正文内容", theme="slate")["theme"] == "slate"
    assert "正文内容" in acks_office.extract(str(word))
    before = word.read_bytes()
    stamped = acks_office.add_watermark(str(word), "内部资料")
    assert stamped["output_path"] == str(tmp_path / "r_watermarked.docx") and word.read_bytes() == before

    pdf = tmp_path / "r.pdf"
    assert acks_office.create("pdf", str(pdf), title="Report", content="Body")["theme"] == "neutral"
    before = pdf.read_bytes()
    assert acks_office.add_watermark(str(pdf), "DRAFT")["output_path"] == str(tmp_path / "r_watermarked.pdf")
    assert pdf.read_bytes() == before

    sheet = tmp_path / "t.xlsx"
    acks_office.create("xlsx", str(sheet), data=[["名称", "数量"], ["甲", 1]])
    assert acks_office.extract(str(sheet)) == [{"名称": "甲", "数量": 1}]

    merged = tmp_path / "m.pdf"
    assert acks_office.merge([str(pdf), str(tmp_path / "r_watermarked.pdf")], str(merged))["merged_pages"] == 2


def test_errors_raise_instead_of_returning_dicts(tmp_path):
    with pytest.raises(ValueError):
        acks_office.create("gif", str(tmp_path / "x.gif"))
    with pytest.raises(ThemeError):
        acks_office.create("word", str(tmp_path / "x.docx"), content="x", theme="nope")
    with pytest.raises(FileNotFoundError):
        acks_office.extract(str(tmp_path / "missing.xlsx"))
    (tmp_path / "note.txt").write_text("x", encoding="utf-8")
    with pytest.raises(ValueError):
        acks_office.extract(str(tmp_path / "note.txt"))
    with pytest.raises(ValueError):
        acks_office.add_watermark(str(tmp_path / "note.txt"), "x")
    with pytest.raises(ValueError):  # 不会把文件转换成自己
        acks_office.convert(str(tmp_path / "note.txt"), "txt")
    pdf = tmp_path / "a.pdf"
    acks_office.create("pdf", str(pdf), content="x")
    with pytest.raises(FileNotFoundError):  # 缺一个文件就不合并
        acks_office.merge([str(pdf), str(tmp_path / "missing.pdf")], str(tmp_path / "m.pdf"))
    assert not (tmp_path / "m.pdf").exists()
    acks_office.create("word", str(tmp_path / "b.docx"), content="x")
    with pytest.raises(ValueError):
        acks_office.merge([str(pdf), str(tmp_path / "b.docx")], str(tmp_path / "m.pdf"))


def test_office_suite_is_deprecated_but_still_works(tmp_path):
    with pytest.warns(DeprecationWarning, match="4.0"):
        suite = acks_office.OfficeSuite()
    result = suite.create("excel", title="T", data=[["a", "b"], [1, 2]], output_path=str(tmp_path / "t.xlsx"))
    assert result["success"] and result["theme"] == "neutral"
    assert not suite.create("gif", output_path=str(tmp_path / "x"))["success"]
    stamped = suite.add_watermark(str(tmp_path / "t.xlsx"), "x")
    assert not stamped["success"] and "水印" in stamped["error"]
