"""命令行：统一的 JSON 结构、退出码、不覆盖已有文件。"""

import json
import os
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

from acks_office import __version__, utils
from acks_office.cli import main

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def run(capsys, tmp_path, monkeypatch):
    """run(*参数) -> (退出码, JSON)；在临时目录里执行。"""
    monkeypatch.chdir(tmp_path)

    def call(*argv):
        code = main([*argv, "--json"])
        return code, json.loads(capsys.readouterr().out)
    return call


def test_version(run):
    code, env = run("version")
    assert code == 0 and env == {"ok": True, "data": {"version": __version__}, "artifacts": [],
                                 "warnings": [], "error": None}


def test_json_flag_may_come_first(capsys):
    assert main(["--json", "version"]) == 0
    assert json.loads(capsys.readouterr().out)["data"]["version"] == __version__


def test_usage_errors_are_json_with_exit_code_2(run):
    code, env = run("create", "word")  # 缺 -o
    assert code == 2 and not env["ok"] and env["error"]["code"] == "USAGE_ERROR"


def test_create_word_from_file_and_refuse_to_overwrite(run, tmp_path, png):
    png(name="pic.png")
    (tmp_path / "正文.md").write_text("# 概述\n\n**重点**\n\n![图](pic.png)", encoding="utf-8")

    code, env = run("create", "word", "-o", "报告.docx", "--content-file", "正文.md", "--brand-name", "")
    # 图片按正文文件所在目录找到（字体提醒取决于本机装了哪些字体，不计入）
    assert code == 0 and env["ok"] and not [w for w in env["warnings"] if not w["code"].startswith("FONT_")]
    (artifact,) = env["artifacts"]
    assert Path(artifact["path"]) == (tmp_path / "报告.docx").resolve() and artifact["size"] > 0

    code, env = run("create", "word", "-o", "报告.docx", "--content", "新内容")
    assert code == 1 and env["error"]["code"] == "OUTPUT_EXISTS" and env["error"]["hint"]

    code, env = run("create", "word", "-o", "报告.docx", "--content", "新内容", "--overwrite")
    assert code == 0


def test_content_file_in_gbk_and_with_bom(run, tmp_path):
    (tmp_path / "gbk.md").write_bytes("中文正文".encode("gbk"))
    (tmp_path / "bom.md").write_bytes("﻿# 标题".encode("utf-8"))
    for name in ("gbk.md", "bom.md"):
        code, env = run("create", "pdf", "-o", f"{name}.pdf", "--content-file", name)
        assert code == 0, env
    code, env = run("extract", "gbk.md.pdf")
    assert "中文正文" in env["data"]["text"]


def test_create_excel_from_csv_keeps_ids_as_text(run, tmp_path):
    (tmp_path / "data.csv").write_text("工号,姓名,销量,单价\n00123,张三,150,9.5\n123456789012345678,李四,-20,0\n",
                                       encoding="utf-8-sig")
    code, env = run("create", "excel", "-o", "数据.xlsx", "--data-file", "data.csv")
    assert code == 0, env

    ws = openpyxl.load_workbook(tmp_path / "数据.xlsx").active
    values = [[c.value for c in row if c.value is not None] for row in ws.iter_rows()]
    assert ["00123", "张三", 150, 9.5] in values
    assert ["123456789012345678", "李四", -20, 0] in values


def test_create_excel_from_json_and_bad_input(run, tmp_path):
    (tmp_path / "ok.json").write_text(json.dumps([["部门", "1月"], ["一部", 150]]), encoding="utf-8")
    (tmp_path / "bad.json").write_text("{not json", encoding="utf-8")
    (tmp_path / "flat.json").write_text(json.dumps([1, 2, 3]), encoding="utf-8")
    assert run("create", "excel", "-o", "ok.xlsx", "--data-file", "ok.json")[0] == 0
    for name in ("bad.json", "flat.json"):
        code, env = run("create", "excel", "-o", "x.xlsx", "--data-file", name)
        assert code == 1 and env["error"]["code"] == "INVALID_INPUT"
    code, env = run("create", "excel", "-o", "x.xlsx")
    assert env["error"]["code"] == "INVALID_INPUT"


def test_create_pptx_and_extract(run, tmp_path):
    slides = [{"title": "封面", "layout": "title"}, {"title": "要点", "content": "第一行\n第二行"}]
    (tmp_path / "slides.json").write_text(json.dumps(slides, ensure_ascii=False), encoding="utf-8")
    code, env = run("create", "pptx", "-o", "deck.pptx", "--slides-file", "slides.json", "--brand-name", "")
    assert code == 0 and env["data"]["slides_count"] == 2

    code, env = run("extract", "deck.pptx")
    assert env["data"]["format"] == "pptx" and "第二行" in env["data"]["slides"]["2"]

    (tmp_path / "empty.json").write_text("[]", encoding="utf-8")
    code, env = run("create", "pptx", "-o", "e.pptx", "--slides-file", "empty.json")
    assert env["error"]["code"] == "INVALID_INPUT"


def test_output_extension_must_match_type(run, tmp_path):
    code, env = run("create", "word", "-o", "报告.word", "--content", "正文")
    assert code == 1 and env["error"]["code"] == "INVALID_INPUT" and "报告.docx" in env["error"]["hint"]
    assert not (tmp_path / "报告.word").exists()
    assert run("create", "docx", "-o", "报告.DOCX", "--content", "正文")[0] == 0


def test_extract_rejects_unknown_formats(run, tmp_path):
    for name in ("a.txt", "a.doc"):
        (tmp_path / name).write_text("x", encoding="utf-8")
        code, env = run("extract", name)
        assert code == 1 and env["error"]["code"] == "UNSUPPORTED_FORMAT"
    assert "convert" in env["error"]["hint"]


def test_extract_excel_and_missing_file(run, tmp_path):
    (tmp_path / "d.json").write_text(json.dumps([["名称", "数量"], ["甲", 1]]), encoding="utf-8")
    run("create", "excel", "-o", "d.xlsx", "--data-file", "d.json")
    code, env = run("extract", "d.xlsx", "--sheet", "0")
    assert code == 0 and env["data"] == {"format": "xlsx", "rows": [{"名称": "甲", "数量": 1}]}

    code, env = run("extract", "nope.docx")
    assert code == 1 and env["error"]["code"] == "FILE_NOT_FOUND"


def test_watermark_writes_a_new_file(run, tmp_path):
    run("create", "pdf", "-o", "a.pdf", "--content", "正文")
    before = (tmp_path / "a.pdf").read_bytes()
    code, env = run("watermark", "a.pdf", "--text", "内部资料")
    assert code == 0 and Path(env["artifacts"][0]["path"]).name == "a_watermarked.pdf"
    assert (tmp_path / "a.pdf").read_bytes() == before


def test_merge_rules(run, tmp_path):
    run("create", "pdf", "-o", "a.pdf", "--content", "甲")
    run("create", "word", "-o", "b.docx", "--content", "乙")

    code, env = run("merge", "a.pdf", "missing.pdf", "-o", "m.pdf")  # 3.0：缺文件直接报错，不再只合并一部分
    assert code == 1 and env["error"]["code"] == "FILE_NOT_FOUND" and "missing.pdf" in env["error"]["message"]
    assert not (tmp_path / "m.pdf").exists()
    assert run("merge", "missing1.pdf", "missing2.pdf", "-o", "n.pdf")[1]["error"]["code"] == "FILE_NOT_FOUND"
    assert run("merge", "a.pdf", "b.docx", "-o", "x.pdf")[1]["error"]["code"] == "UNSUPPORTED_FORMAT"


def test_convert_without_libreoffice(run, tmp_path, monkeypatch):
    run("create", "word", "-o", "a.docx", "--content", "正文")
    monkeypatch.setattr(utils, "find_soffice", lambda: None)
    code, env = run("convert", "a.docx", "--to", "pdf")
    assert code == 1 and env["error"]["code"] == "ENGINE_UNAVAILABLE"
    assert not (tmp_path / "a.pdf").exists()


def test_convert_refuses_to_replace_source(run, tmp_path):
    run("create", "word", "-o", "a.docx", "--content", "正文")
    code, env = run("convert", "a.docx", "--to", "docx")
    assert env["error"]["code"] == "OUTPUT_EXISTS"


def test_doctor_and_fonts(run, monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    code, env = run("doctor", "--no-network", "--no-probe")
    assert code == 0 and env["data"]["schema"] == "doctor/1"
    code, env = run("fonts", "list")
    assert code == 0 and env["data"]["catalog"][0]["key"] == "noto-sans-sc"
    code, env = run("fonts", "install", "no-such-font")
    assert code == 1 and env["error"]["code"] == "FONT_INSTALL_FAILED"


def test_module_entry_point():
    env = dict(os.environ, PYTHONPATH=str(ROOT / "src"))
    out = subprocess.run([sys.executable, "-m", "acks_office", "--version"], capture_output=True, text=True, env=env)
    assert out.returncode == 0 and out.stdout.strip() == f"acks-office {__version__}"
