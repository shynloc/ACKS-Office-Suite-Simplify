"""旧包名 office_suite 仍可导入，指向同一套实现。"""

import importlib
import sys

import pytest

import acks_office


def _fresh_import(name):
    for key in [k for k in sys.modules if k == "office_suite" or k.startswith("office_suite.")]:
        del sys.modules[key]
    return importlib.import_module(name)


def test_old_package_warns_and_reexports():
    with pytest.warns(DeprecationWarning, match="acks_office"):
        office_suite = _fresh_import("office_suite")
    assert office_suite.OfficeSuite is acks_office.OfficeSuite
    assert office_suite.__version__ == acks_office.__version__


def test_old_submodule_paths_point_to_new_modules():
    with pytest.warns(DeprecationWarning):
        _fresh_import("office_suite")
    from office_suite.xlsx import extract_data
    from office_suite.design_system.tokens import add_cover_page
    import acks_office.xlsx
    import acks_office.design_system.tokens

    assert extract_data is acks_office.xlsx.extract_data
    assert add_cover_page is acks_office.design_system.tokens.add_cover_page
    assert sys.modules["office_suite.docx"] is sys.modules["acks_office.docx"]
