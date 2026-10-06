"""LibreOffice 查找与转换、用户数据目录。"""

import shutil
from pathlib import Path

import pytest

from acks_office import OfficeSuite, utils
from acks_office.docx import create_word


@pytest.fixture
def fake_soffice(tmp_path):
    path = tmp_path / "soffice"
    path.write_text("")
    return str(path)


def test_env_override_wins(monkeypatch, fake_soffice):
    monkeypatch.setenv("ACKS_OFFICE_SOFFICE", fake_soffice)
    assert utils.find_soffice() == fake_soffice


def test_env_override_pointing_nowhere_finds_nothing(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_SOFFICE", str(tmp_path / "missing"))
    assert utils.find_soffice() is None


def test_standard_install_location_is_used_when_not_on_path(monkeypatch, fake_soffice):
    monkeypatch.delenv("ACKS_OFFICE_SOFFICE", raising=False)
    monkeypatch.setattr(shutil, "which", lambda name: None)
    monkeypatch.setattr(utils, "soffice_candidates", lambda: ["/nonexistent/soffice", fake_soffice])
    assert utils.find_soffice() == fake_soffice


def test_path_lookup(monkeypatch, fake_soffice):
    monkeypatch.delenv("ACKS_OFFICE_SOFFICE", raising=False)
    monkeypatch.setattr(shutil, "which", lambda name: fake_soffice if name == "libreoffice" else None)
    assert utils.find_soffice() == fake_soffice


def test_candidates_are_absolute_paths():
    assert utils.soffice_candidates() and all(Path(p).is_absolute() for p in utils.soffice_candidates())


def test_missing_libreoffice_is_explained(monkeypatch, tmp_path):
    monkeypatch.setattr(utils, "find_soffice", lambda: None)
    with pytest.raises(RuntimeError, match="未找到 LibreOffice"):
        utils.convert_with_libreoffice(str(tmp_path / "a.docx"), "pdf", str(tmp_path / "a.pdf"))


def test_data_dir_override(monkeypatch, tmp_path):
    monkeypatch.setenv("ACKS_OFFICE_HOME", str(tmp_path / "home"))
    assert utils.data_dir() == tmp_path / "home"


@pytest.mark.skipif(utils.find_soffice() is None, reason="本机没有 LibreOffice")
def test_real_conversion_to_pdf(tmp_path):
    source = tmp_path / "报告.docx"
    create_word("季度报告", "正文", str(source))
    target = tmp_path / "out" / "报告.pdf"

    result = OfficeSuite().convert(str(source), to="pdf", output_path=str(target))

    assert result["success"] and result["engine"] == "libreoffice"
    assert target.read_bytes().startswith(b"%PDF")
    assert sorted(p.name for p in tmp_path.iterdir()) == ["out", "报告.docx"]  # 没有残留临时文件
