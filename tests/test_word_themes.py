"""主题驱动的 Word：三种封面、页眉页脚、编号、表格、提示块、Folio 的章节页与分栏、front matter。"""

import json
import zipfile

import pytest
from docx import Document
from docx.oxml.ns import qn

from acks_office import fonts, markdown_blocks as mdb, themes
from acks_office.cli import main
from acks_office.docx import create_word, extract_text
from acks_office.render.common import (attach_captions, document_meta, is_total_row, numeric_columns,
                                       split_front_matter)

REPORT = """---
title: 2026 年第三季度\\n经营回顾
short_title: 经营回顾
kicker: 经营分析报告
subtitle: 营收、门店与会员
brand: 栖木咖啡
classification: 内部资料
date: 2026 年 10 月 8 日
author: 经营分析部
version: v1.0
---

# 经营概览

第三季度营收 **1.28 亿元**，同比增长 24.0%。

## 分区域营收

表：分区域营收（单位：万元）

| 区域 | 营收 | 同比 |
|---|--:|--:|
| 华东 | 5,888 | +25.8% |
| 华南 | 3,456 | +22.0% |
| 合计 | 9,344 | +24.0% |

> [!TIP]
> 毛利率回落主要来自采购价上涨。

- 华东加密
- 稳住成本

> 一句引文。
"""

MAGAZINE = """---
title: 慢一点，\\n也走得远
publication: 栖木咖啡季度刊
issue: "07"
season: Autumn 2026
lede: 这是一份关于耐心的季报。
---

# 一个季度的账本 {label="生意 · THE BUSINESS"}

> 增长和成本，这个季度都在提醒我们：慢一点。

第三季度，营收第一次突破 1.2 亿元。

> 增长不靠打折，靠的是有人愿意再来一次。

| 区域 | 营收 |
|---|--:|
| 华东 | 5,888 |

# 选址逻辑 {label="门店 · STORES"}

选址看三件事。
"""


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    themes.clear_cache()
    yield
    themes.clear_cache()


def _render(tmp_path, theme, content=REPORT, **kwargs):
    out = tmp_path / f"{theme}.docx"
    result = create_word(None, content, str(out), theme=theme, **kwargs)
    return out, result, Document(str(out))


def _xml(path, part):
    with zipfile.ZipFile(path) as z:
        return "".join(z.read(n).decode("utf-8") for n in z.namelist() if n.startswith(f"word/{part}"))


def _texts(container):
    return [p.text for p in container.paragraphs]


def test_neutral_title_block_without_cover_page(tmp_path):
    out, result, doc = _render(tmp_path, "neutral")
    assert result["theme"] == "neutral"
    assert len(doc.sections) == 1 and doc.sections[0].different_first_page_header_footer
    body = _texts(doc)
    assert body[:3] == ["经营分析报告", "2026 年第三季度\n经营回顾", "营收、门店与会员"]
    assert "2026 年 10 月 8 日 · 经营分析部 · v1.0" in body
    assert doc.paragraphs[1].style.name == "Title"


def test_slate_cover_page_with_meta_in_footer(tmp_path):
    out, result, doc = _render(tmp_path, "slate")
    cover, body = doc.sections[0], doc.sections[1]
    assert len(doc.sections) == 2
    assert _texts(doc)[:3] == ["栖木咖啡\t内部资料", "", "经营分析报告"]
    cells = [c.text for c in cover.footer.tables[0]._cells]
    assert cells == ["日期", "2026 年 10 月 8 日", "编制", "经营分析部", "版本", "v1.0"]
    assert body.header.paragraphs[0].text == "2026 年第三季度经营回顾\t栖木咖啡"  # 标题合成一行
    footer = body.footer.paragraphs[0]
    assert footer.text.startswith("栖木咖啡 · 经营回顾\t") and "PAGE" in footer._p.xml
    assert cover.header.is_linked_to_previous  # 封面没有页眉


def test_slate_headings_are_numbered_with_word_numbering(tmp_path):
    out, _, doc = _render(tmp_path, "slate")
    heading = doc.styles["Heading 1"].element
    assert heading.find(qn("w:pPr")).find(qn("w:numPr")) is not None
    numbering = _xml(out, "numbering")
    assert 'w:val="%1.%2"' in numbering and "1F5FAE" in numbering  # 编号用强调色


def test_slate_table_caption_header_fill_and_total_row(tmp_path):
    out, _, doc = _render(tmp_path, "slate")
    assert "表 1　分区域营收（单位：万元）" in _texts(doc)
    (table,) = [t for t in doc.tables if t._cells[0].text == "区域"]
    header, total = table.rows[0], table.rows[-1]
    assert [c.text for c in header.cells] == ["区域", "营收", "同比"]
    assert 'w:fill="F5F7FA"' in header.cells[0]._tc.xml
    assert all(r.font.bold for c in total.cells for p in c.paragraphs for r in p.runs)
    assert header.cells[1].paragraphs[0].alignment == 2  # 数字列右对齐
    assert "区域 | 营收 | 同比" in extract_text(str(out))


def test_total_row_rules_follow_the_theme(tmp_path):
    def total_borders(theme):
        _, _, doc = _render(tmp_path, theme)
        (table,) = [t for t in doc.tables if t._cells[0].text == "区域"]
        borders = table.rows[-1].cells[1]._tc.tcPr.find(qn("w:tcBorders"))
        return {child.tag.split("}")[1]: (child.get(qn("w:sz")), child.get(qn("w:color"))) for child in borders}

    slate, folio = themes.load_theme("slate"), themes.load_theme("folio")
    assert total_borders("slate") == {"top": ("6", slate.hex("rule_strong")), "bottom": ("6", slate.hex("rule_strong"))}
    assert total_borders("folio") == {"top": ("12", folio.hex("accent"))}  # bar：上方一条粗线


def test_callout_and_quote_variants(tmp_path):
    _, _, slate = _render(tmp_path, "slate")
    callout = [t for t in slate.tables if "提示" in t._cells[0].text]
    assert callout and 'w:fill="F5F7FA"' in callout[0]._cells[0]._tc.xml
    quote = [p for p in slate.paragraphs if p.text == "一句引文。"][0]
    assert quote.style.name == "Quote" and "w:left" in quote._p.pPr.xml

    _, _, folio = _render(tmp_path, "folio", MAGAZINE + "\n> [!NOTE]\n> 说明内容\n")
    note = [t for t in folio.tables if t._cells[0].text.startswith("说明")]
    assert note and "w:fill" not in note[0]._cells[0]._tc.xml  # folio 的提示块是上下细线
    pull = [t for t in folio.tables if t._cells[0].text == "“"]
    assert pull and "增长不靠打折" in pull[0]._cells[1].text


def test_folio_issue_cover_chapters_columns_and_drop_cap(tmp_path):
    out, result, doc = _render(tmp_path, "folio", MAGAZINE)
    cover = doc.sections[0]
    assert _texts(doc)[0] == "栖木咖啡季度刊\tAUTUMN 2026"
    assert any(p.text == "No. 07" for p in doc.paragraphs)
    toc = cover.footer
    assert toc.paragraphs[0].text == "本期目录"
    rows = [[c.text for c in r.cells] for r in toc.tables[0].rows]
    assert [r[:2] for r in rows] == [["01", "生意　一个季度的账本"], ["02", "门店　选址逻辑"]]
    footer_xml = _xml(out, "footer")
    assert "PAGEREF _acks_chapter_1" in footer_xml and 'w:dirty="true"' in footer_xml
    document_xml = _xml(out, "document")
    assert 'w:name="_acks_chapter_2"' in document_xml
    assert 'w:dropCap="drop"' in document_xml and 'w:num="2"' in document_xml
    ledes = [p.text for p in doc.paragraphs if p.style.name == "Lede"]
    assert ledes == ["这是一份关于耐心的季报。", "增长和成本，这个季度都在提醒我们：慢一点。"]  # 封面导语、章节导语
    headers = {s.header.paragraphs[0].text for s in doc.sections if not s.header.is_linked_to_previous}
    assert "栖木咖啡季度刊 · 第 07 期\t生意" in headers and "栖木咖啡季度刊 · 第 07 期\t门店" in headers
    assert [p.text for p in doc.paragraphs if p.style.name == "Heading 1"] == ["一个季度的账本", "选址逻辑"]
    text = extract_text(str(out))
    assert "第三季度，营收第一次突破" in text.replace("第\n三", "第三")  # 首字下沉拆成了两段


def test_brand_and_footer_overrides(tmp_path):
    _, _, doc = _render(tmp_path, "slate", brand_name="", footer_label="内部资料 · 请勿外传")
    body = doc.sections[1]
    assert body.header.paragraphs[0].text == "2026 年第三季度经营回顾\t"
    assert body.footer.paragraphs[0].text.startswith("内部资料 · 请勿外传\t")
    assert _texts(doc)[0] == "\t内部资料"


def test_cover_can_be_turned_off(tmp_path):
    _, _, doc = _render(tmp_path, "slate", REPORT.replace("version: v1.0", "version: v1.0\ncover: false"))
    assert len(doc.sections) == 1 and doc.paragraphs[1].style.name == "Title"


def test_font_policy_and_missing_fonts(tmp_path, monkeypatch):
    monkeypatch.setattr(fonts, "family_installed", lambda name: False)
    out, result, _ = _render(tmp_path, "slate")
    codes = {w["code"] for w in result["warnings"]}
    assert codes == {"FONT_MISSING"}
    styles = _xml(out, "styles")
    assert 'w:eastAsia="Noto Sans SC"' in styles and 'w:ascii="Source Sans 3"' in styles

    monkeypatch.setattr(fonts, "family_installed", lambda name: name in ("PingFang SC", "Helvetica Neue"))
    themes.clear_cache()
    out, result, _ = _render(tmp_path, "slate")
    assert {w["code"] for w in result["warnings"]} == {"FONT_SUBSTITUTED", "FONT_MISSING"}
    assert 'w:eastAsia="PingFang SC"' in _xml(out, "styles")
    out, result, _ = _render(tmp_path, "slate", font_policy="theme")
    assert 'w:eastAsia="Noto Sans SC"' in _xml(out, "styles")


def test_theme_from_directory(tmp_path):
    custom = tmp_path / "brand"
    custom.mkdir()
    (custom / "theme.json").write_text(json.dumps({"name": "brand", "extends": "slate"}), encoding="utf-8")
    (custom / "tokens.json").write_text(json.dumps({"color": {"accent": "#0B6E4F"}}), encoding="utf-8")
    out, result, _ = _render(tmp_path, str(custom))
    assert result["theme"] == "brand" and "0B6E4F" in _xml(out, "numbering")


def test_missing_theme_raises(tmp_path):
    with pytest.raises(themes.ThemeError):
        create_word("标题", "正文", str(tmp_path / "x.docx"), theme="no-such-theme")


@pytest.mark.parametrize("text, meta, body", [
    ("---\ntitle: 季报\nissue: \"07\"\ncover: false\n---\n正文", {"title": "季报", "issue": "07", "cover": False}, "正文"),
    ("---\ntitle: 两行\\n标题\n---\n", {"title": "两行\n标题"}, ""),
    ("没有 front matter", {}, "没有 front matter"),
    ("---\nnot closed\n正文", {}, "---\nnot closed\n正文"),
])
def test_split_front_matter(text, meta, body):
    assert split_front_matter(text) == (meta, body)


def test_document_meta_precedence():
    meta, body = document_meta("参数标题", "---\ntitle: 正文标题\nauthor: 甲\n---\n内容",
                               {"author": "乙", "date": None})
    assert meta == {"title": "参数标题", "author": "乙"} and body == "内容"
    meta, _ = document_meta(None, "---\ntitle: 正文标题\n---\n", {})
    assert meta["title"] == "正文标题"


def test_captions_attach_before_or_after_tables():
    blocks = mdb.parse("表：之前\n\n| a |\n|---|\n| 1 |\n\n| b |\n|---|\n| 2 |\n\nTable: after\n\n普通段落")
    pairs = [(type(b).__name__, c) for b, c in attach_captions(blocks)]
    assert pairs == [("Table", "之前"), ("Table", "after"), ("Paragraph", None)]


def test_total_rows_and_numeric_columns():
    assert is_total_row(["合计", "1"]) and is_total_row(["Total", "1"]) and is_total_row(["合计（万元）", "1"])
    assert not is_total_row(["合计划", ""]) and not is_total_row([])
    assert numeric_columns(["区域", "营收", "备注"], [["华东", "5,888", "新店"], ["华南", "(1.2)", ""]],
                           [None, None, None]) == [1]


def test_cli_create_with_theme_and_meta(tmp_path, capsys, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "正文.md").write_text(MAGAZINE, encoding="utf-8")
    code = main(["create", "word", "-o", "刊物.docx", "--content-file", "正文.md", "--theme", "folio", "--json"])
    env = json.loads(capsys.readouterr().out)
    assert code == 0 and env["data"]["theme"] == "folio"
    doc = Document(str(tmp_path / "刊物.docx"))
    assert any(p.text == "No. 07" for p in doc.paragraphs)
    assert doc.core_properties.title == "慢一点，\n也走得远"

    code = main(["create", "word", "-o", "a.docx", "--content", "正文", "--theme", "slate",
                 "--meta", "kicker=测试报告", "--meta", "author=研发部", "--json"])
    assert code == 0 and json.loads(capsys.readouterr().out)["ok"]
    assert "测试报告" in _texts(Document(str(tmp_path / "a.docx")))
    assert main(["create", "word", "-o", "b.docx", "--content", "x", "--meta", "坏格式", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "INVALID_INPUT"
    assert main(["create", "word", "-o", "c.docx", "--content", "x", "--theme", "nope", "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "THEME_NOT_FOUND"
