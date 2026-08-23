# -*- coding: utf-8 -*-
"""ACKS Studio 文档设计规范 v2.1 —— 集成进 office_suite。

使 office_suite 产出的 Word / PPT / Excel 文档直接符合 ACKS 设计规范。

来源：Thom & 银月共享知识库 05-Resources/文档设计规范-v2/
详见本目录 SOURCE.md 的溯源说明。

用法：
    from office_suite.design_system import tokens, slides, xlsx

    # Word
    from office_suite.design_system.tokens import (
        apply_default_styles, set_a4_page_margins,
        add_cover_page, add_section_heading, add_body_paragraph, ...
    )
"""
from . import tokens, slides, xlsx

__version__ = "2.1.0"
__all__ = ["tokens", "slides", "xlsx"]
