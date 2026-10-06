"""PDF：中文字体嵌入与回退、Markdown 呈现、页数。"""

from pathlib import Path

import pytest
import reportlab
from pypdf import PdfReader

from acks_office import fonts
from acks_office.pdf import add_watermark, create_pdf


def _fonts(path):
    """{字体名: 是否嵌入了 TrueType 字形}"""
    found = {}
    for page in PdfReader(str(path)).pages:
        for ref in page["/Resources"]["/Font"].values():
            font = ref.get_object()
            descriptor = font.get("/FontDescriptor")
            if descriptor is None and "/DescendantFonts" in font:
                descriptor = font["/DescendantFonts"][0].get_object().get("/FontDescriptor")
            found[font["/BaseFont"]] = bool(descriptor) and "/FontFile2" in descriptor.get_object()
    return found


def _text(path):
    return "\n".join(page.extract_text() for page in PdfReader(str(path)).pages)


def test_chinese_is_embedded_with_given_font(tmp_path, cjk_font):
    out = tmp_path / "a.pdf"
    result = create_pdf("中文测试报告", "第一行\n\n第二行", str(out), font=cjk_font)

    assert result["theme"] == "neutral" and "AcksTest" in result["fonts"]
    assert not any(w["code"] == "FONT_NOT_EMBEDDED" for w in result.get("warnings", []))
    embedded = [name for name, ok in _fonts(out).items() if ok]
    assert any(name.endswith("AcksTest-Regular") for name in embedded)  # 中文用指定的字体并嵌入
    text = _text(out)
    assert all(part in text for part in ("中文测试报告", "第一行", "第二行"))


def test_font_from_environment_variable(tmp_path, cjk_font, monkeypatch):
    monkeypatch.setenv("ACKS_OFFICE_PDF_FONT", cjk_font)
    result = create_pdf("中文测试", "正文", str(tmp_path / "e.pdf"))
    assert "AcksTest" in result["fonts"]


def test_falls_back_to_reader_font_with_warning(tmp_path, monkeypatch):
    monkeypatch.setattr(fonts, "best_face", lambda *a, **k: None)  # 本机没有任何能嵌入的字体
    monkeypatch.delenv("ACKS_OFFICE_PDF_FONT", raising=False)
    out = tmp_path / "f.pdf"
    result = create_pdf("中文测试", "第一行", str(out))

    assert result["fonts"] == []
    assert "FONT_NOT_EMBEDDED" in [w["code"] for w in result["warnings"]]
    assert "中文测试" in _text(out)


def test_unusable_font_is_rejected(tmp_path):
    junk = tmp_path / "junk.ttf"
    junk.write_bytes(b"not a font")
    latin_only = Path(reportlab.__file__).parent / "fonts" / "Vera.ttf"  # 没有中文字形
    for path in (junk, latin_only):
        with pytest.raises(ValueError, match="字体无法用于 PDF"):
            fonts.pdf_fonts(str(path))


def test_markup_characters_are_text_not_tags(tmp_path, cjk_font):
    out = tmp_path / "m.pdf"
    create_pdf("中文测试", "a < b & c > d <b>粗体</b> &lt;", str(out), font=cjk_font)
    assert "a < b & c > d <b>粗体</b> <" in _text(out)  # 最后的 &lt; 是 Markdown 的字符实体


def test_markdown_structures(tmp_path, cjk_font):
    out = tmp_path / "s.pdf"
    content = ("## 区域营收\n\n| 区域 | 营收 |\n|---|--:|\n| 华东 | 5,888 |\n\n"
               "1. 步骤一\n2. 步骤二\n\n- [ ] 待办\n- [x] 完成\n\n> [!NOTE]\n> 段落\n\n```\ncode < 1\n```")
    result = create_pdf("中文测试", content, str(out), font=cjk_font)

    text = _text(out)
    # 提示块的标签来自主题（「说明」），测试字体里没有这两个字，这里只核对正文
    for expected in ("区域营收", "华东", "5,888", "1.", "步骤二", "□", "待办", "■", "完成", "段落", "code < 1"):
        assert expected in text
    assert "ZapfDingbats" not in "".join(_fonts(out))  # 任务标记用正文字体，不靠符号字体替换
    assert result["pages"] == 1


def test_reports_real_page_count(tmp_path, cjk_font):
    out = tmp_path / "long.pdf"
    result = create_pdf("中文测试", "\n\n".join(f"第{i}段落" for i in range(300)), str(out), font=cjk_font)
    assert result["pages"] == len(PdfReader(str(out)).pages) > 3


def test_image_and_missing_image(tmp_path, cjk_font, png):
    png(300, 150, name="pic.png")
    result = create_pdf("中文测试", "![图注](pic.png)\n\n![](nope.png)", str(tmp_path / "i.pdf"),
                        font=cjk_font, base_dir=str(tmp_path))
    assert [w["code"] for w in result["warnings"]] == ["IMAGE_NOT_FOUND"]
    assert len(PdfReader(str(tmp_path / "i.pdf")).pages[0].images) == 1


@pytest.mark.filterwarnings("error::DeprecationWarning")
def test_watermark_follows_each_page_size(tmp_path):
    from reportlab.pdfgen import canvas
    source = tmp_path / "mixed.pdf"
    c = canvas.Canvas(str(source))
    for size in ((595, 842), (842, 595)):  # 纵向 A4 + 横向 A4
        c.setPageSize(size)
        c.drawString(50, 50, "page")
        c.showPage()
    c.save()
    before = source.read_bytes()

    result = add_watermark(str(source), "INTERNAL", str(tmp_path / "wm.pdf"))

    assert result["watermarked_pages"] == 2 and source.read_bytes() == before
    pages = PdfReader(str(tmp_path / "wm.pdf")).pages
    assert [(float(p.mediabox.width), float(p.mediabox.height)) for p in pages] == [(595, 842), (842, 595)]
    assert all("INTERNAL" in p.extract_text() and "page" in p.extract_text() for p in pages)
