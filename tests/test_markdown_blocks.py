"""Markdown → 块结构：Word 与 PDF 共用的解析结果。"""

from acks_office import markdown_blocks as mdb


def test_inline_spans_keep_formatting_and_links():
    (para,) = mdb.parse("普通 **粗体** *斜体* ~~删除~~ `代码` [链接](https://example.com/?a=1&b=2)")
    spans = {s.text: s for s in para.spans}
    assert spans["粗体"].bold and spans["斜体"].italic and spans["删除"].strike and spans["代码"].code
    assert spans["链接"].link == "https://example.com/?a=1&b=2"
    assert mdb.plain(para.spans) == "普通 粗体 斜体 删除 代码 链接"


def test_single_newline_is_kept_as_line_break():
    (para,) = mdb.parse("第一行\n第二行")
    assert mdb.plain(para.spans) == "第一行\n第二行"


def test_lists_tasks_and_nesting():
    blocks = mdb.parse("3. 三\n4. 四\n\n- 父\n  - 子\n- [ ] 待办\n- [x] 完成")
    ordered, bullets = blocks
    assert ordered.ordered and ordered.start == 3 and len(ordered.items) == 2
    parent, todo, done = bullets.items
    assert mdb.plain(parent.children[0].items[0].spans) == "子"
    assert (todo.checked, mdb.plain(todo.spans)) == (False, "待办")
    assert (done.checked, mdb.plain(done.spans)) == (True, "完成")
    assert parent.checked is None


def test_table_with_alignment():
    (table,) = mdb.parse("| 区域 | 营收 |\n|:--|--:|\n| 华东 | 5,888 |\n| 华南 |")
    assert [mdb.plain(c) for c in table.header] == ["区域", "营收"]
    assert [[mdb.plain(c) for c in row] for row in table.rows] == [["华东", "5,888"], ["华南", ""]]
    assert table.aligns == ["left", "right"]


def test_alerts_quotes_code_rule_and_images():
    blocks = mdb.parse("> [!WARNING]\n> 小心\n\n> 普通引用\n\n```python\nx < 1\n```\n\n---\n\n![图注](a.png)")
    alert, quote, code, rule, image = blocks
    assert alert.alert == "warning" and mdb.plain(alert.blocks[0].spans) == "小心"
    assert quote.alert is None
    assert (code.text, code.lang) == ("x < 1", "python")
    assert isinstance(rule, mdb.Rule)
    assert (image.src, image.alt) == ("a.png", "图注")


def test_raw_html_is_plain_text():
    (para,) = mdb.parse("<b>不是标签</b>")
    assert mdb.plain(para.spans) == "<b>不是标签</b>"
