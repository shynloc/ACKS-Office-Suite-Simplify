"""主题驱动的 PDF：三种封面、章节页与目录页码、页眉页脚、书签、段落拆分、字体回退与旧样式。"""

import json
import re

import pytest
from pypdf import PdfReader

from acks_office import fonts, themes
from acks_office.cli import main
from acks_office.pdf import create_pdf
from acks_office.render import pdffonts
from acks_office.render.pdf import render_pdf

REPORT = """---
title: 2026 年第三季度\\n经营回顾
kicker: 经营分析报告
subtitle: 营收、门店与会员
brand: Qimu Coffee
classification: Internal
date: 2026-10-08
author: Analytics
version: v1.0
---

# 经营概览 Overview

第三季度营收 **1.28 亿元**，同比增长 24.0%。Revenue grew 24.0% year over year.

## 分区域营收 Regions

表：分区域营收（单位：万元）

| 区域 | 营收 | 同比 |
|---|--:|--:|
| East | 5,888 | +25.8% |
| South | 3,456 | +22.0% |
| 合计 | 12,800 | +24.0% |

> [!TIP]
> Gross margin target: 59%.

- Item alpha
- Item beta

1. First step
2. Second step

```text
revenue = stores * average
```

![missing](nope.png)
"""

MAGAZINE = """---
title: Slow down
publication: Qimu Quarterly
issue: "07"
season: Autumn 2026
lede: A quarterly about patience.
---

# The ledger {label="Business · THE BUSINESS"}

> Growth without discounts.

Revenue grew 24 percent. New stores and returning members did the work.

> Customers come back when the coffee is good.

| Region | Revenue |
|---|--:|
| East | 5,888 |
| Total | 12,800 |

# Where the stores go {label="Stores · STORES"}

Site selection looks at commuters, offices and bookshops nearby.

# Why people come back {label="Members · MEMBERS"}

Stable quality, and staff who remember your order.
"""


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    themes.clear_cache()
    yield
    themes.clear_cache()


def _pages(path):
    return [re.sub(r"\s+", " ", page.extract_text() or "") for page in PdfReader(str(path)).pages]


def test_script_runs_and_reportlab_fixes():
    assert pdffonts.script_runs("营收 1.28 亿元，“同比”增长") == [
        ("cn", "营收"), ("en", " 1.28 "), ("cn", "亿元，“同比”增长")]
    assert pdffonts.script_runs("“quoted” 中文") == [("en", "“quoted” "), ("cn", "中文")]
    from reportlab.platypus import paragraph
    assert all(ch in paragraph.ALL_CANNOT_START for ch in "，；：？！")
    assert paragraph._split_blParaHard is pdffonts._join_split_lines
    assert paragraph._justifyDrawParaLineX is pdffonts._justify_line


@pytest.mark.parametrize("theme, cover_pages", [("neutral", 0), ("slate", 1)])
def test_report_layouts(tmp_path, theme, cover_pages):
    out = tmp_path / f"{theme}.pdf"
    result = render_pdf(None, REPORT, str(out), theme=theme)
    pages = _pages(out)
    assert result["theme"] == theme and result["pages"] == len(pages) >= 1
    assert isinstance(result["fonts"], list)
    text = " ".join(pages)
    for marker in ("Overview", "Regions", "5,888", "12,800", "Item alpha", "Second step", "revenue = stores"):
        assert marker in text, marker
    body = pages[cover_pages:]
    assert any("Qimu Coffee" in page for page in body)  # 页脚：品牌
    assert re.search(rf"\b{len(pages)}\b", pages[-1])  # 页脚：页码
    if cover_pages:  # standard 封面：品牌、密级与日期 / 编制 / 版本
        assert "Qimu Coffee" in pages[0] and "Internal" in pages[0] and "Analytics" in pages[0]
    codes = [w["code"] for w in result.get("warnings", [])]
    assert "IMAGE_NOT_FOUND" in codes
    titles = [item.title for item in PdfReader(str(out)).outline if not isinstance(item, list)]
    assert any("Overview" in t for t in titles)


def test_slate_numbers_headings(tmp_path):
    out = tmp_path / "slate.pdf"
    render_pdf(None, REPORT, str(out), theme="slate")
    text = " ".join(_pages(out))
    assert re.search(r"1\s+经营概览|1\s+.*Overview", text) and re.search(r"1\.1\s", text)


def test_folio_cover_toc_chapter_pages_and_headers(tmp_path):
    out = tmp_path / "folio.pdf"
    result = render_pdf(None, MAGAZINE, str(out), theme="folio")
    pages = _pages(out)
    assert result["pages"] == len(pages) == 4  # 封面 + 三个章节页
    cover = pages[0]
    assert "Qimu Quarterly" in cover and "07" in cover and "Slow down" in cover
    for number, chapter in enumerate(("Business", "Stores", "Members"), start=2):
        assert re.search(rf"{chapter} .*? {number}\b", cover), chapter  # 目录里的页码是章节实际所在的页
        assert chapter in pages[number - 1]  # 页眉：章节名
    assert "Growth without discounts." in pages[1] and "Total" in pages[1]
    assert "Qimu Quarterly" in pages[1] and re.search(r"\b2\b", pages[1])


def test_split_paragraphs_keep_spaces(tmp_path):
    english = ("Revenue grew strongly year over year thanks to the internationalization of membership programs "
               "and supply chain optimization and new stores opening across the eastern region. ") * 3
    mixed = "第三季度营收 1.28 亿元，同比增长 24.0%，其中华东地区贡献最大。Revenue grew strongly 门店达到 412 家。" * 6
    content = "# Chapter {label=\"One · ONE\"}\n\n" + "\n\n".join([english, mixed] * 6)
    expected = set(re.findall(r"[A-Za-z]+", english + mixed))
    for theme in ("folio", "slate"):
        out = tmp_path / f"split-{theme}.pdf"
        render_pdf("Split", content, str(out), theme=theme)
        words = set(re.findall(r"[A-Za-z]+", " ".join(_pages(out))))
        merged = {w for w in words - expected - {"Chapter", "One", "ONE", "Split", "evenue"} if len(w) > 3}
        assert not merged, (theme, sorted(merged)[:5])


def test_falls_back_to_reader_font_when_nothing_embeddable(tmp_path, monkeypatch):
    monkeypatch.setattr(fonts, "best_face", lambda *a, **k: None)
    out = tmp_path / "fallback.pdf"
    result = render_pdf(None, "# Title\n\nPlain text 中文。", str(out), theme="neutral")
    codes = {w["code"] for w in result["warnings"]}
    assert "FONT_NOT_EMBEDDED" in codes and result["fonts"] == []
    assert "Plain text" in " ".join(_pages(out))


def test_font_override_must_be_truetype(tmp_path):
    bogus = tmp_path / "bogus.ttf"
    bogus.write_bytes(b"not a font")
    with pytest.raises(ValueError):
        create_pdf(None, "text", str(tmp_path / "x.pdf"), font=str(bogus), theme="slate")


def test_default_pdf_theme_is_neutral(tmp_path):
    assert create_pdf("Plain", "Hello", str(tmp_path / "plain.pdf"))["theme"] == "neutral"
    assert create_pdf("Alias", "Hello", str(tmp_path / "alias.pdf"), theme="default")["theme"] == "neutral"


def test_cli_create_pdf_with_theme(tmp_path, capsys, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "report.md").write_text(MAGAZINE, encoding="utf-8")
    code = main(["create", "pdf", "-o", "report.pdf", "--content-file", "report.md", "--theme", "folio", "--json"])
    env = json.loads(capsys.readouterr().out)
    assert code == 0 and env["data"]["theme"] == "folio" and env["data"]["pages"] == 4
    assert PdfReader(str(tmp_path / "report.pdf")).metadata.title == "Slow down"
