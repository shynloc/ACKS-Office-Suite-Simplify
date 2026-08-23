# -*- coding: utf-8 -*-
"""
ACKS Studio · Spreadsheet Design System v2.0
openpyxl token definitions and sheet-template factories.

Changes from v1.0:
  - Constants imported from xlsx_constants.py (auto-generated from tokens.json)
  - __all__ defines the public API
  - Sheet builder functions accept data parameters for reuse
  - Demo data isolated in save_demo() / build_workbook()

REQUIREMENTS
    pip install openpyxl

USAGE
    from xlsx import save_demo
    save_demo("ACKS-workbook.xlsx")

NEVER edit token values without updating tokens.json and running:
    python scripts/build_tokens.py
"""

from __future__ import annotations
from typing import Iterable, Mapping, Optional, Sequence

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, NamedStyle, PatternFill, Side,
)
from openpyxl.formatting.rule import (
    ColorScaleRule, DataBarRule, CellIsRule, FormulaRule,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Auto-generated constants from tokens.json
from .xlsx_constants import *  # noqa: F401,F403

__all__ = [
    "register_named_styles",
    "build_readme_sheet", "build_dashboard_sheet",
    "build_data_sheet", "build_model_sheet",
    "build_workbook", "save_demo",
]


# =============================================================
#                    STYLE PRIMITIVES
# =============================================================

def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def _border(*, top=None, bottom=None, left=None, right=None,
            color=XLSX_GRID) -> Border:
    def side(style):
        return Side(style=style, color=color) if style else Side(style=None)
    return Border(
        top=side(top), bottom=side(bottom),
        left=side(left), right=side(right),
    )

def _font(name=FONT_BODY, size=10, bold=False, color=XLSX_TEXT_PRIMARY,
          italic=False) -> Font:
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


# =============================================================
#                    NAMED STYLES (14)
# =============================================================

def register_named_styles(wb: Workbook) -> None:
    """Register all 14 ACKS_CELL_* NamedStyles on the workbook."""
    styles = [
        # HEADERS
        NamedStyle(
            name="ACKS_CELL_H1",
            fill=_fill(XLSX_HEADER_BG),
            font=_font(FONT_HEADER, 11, True, XLSX_BG_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_H2",
            fill=_fill(XLSX_SUBHEADER_BG),
            font=_font(FONT_HEADER, 10, True, XLSX_BG_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_H3",
            fill=_fill(XLSX_TOTAL_BG),
            font=_font(FONT_HEADER, 10, True, XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        # DATA
        NamedStyle(
            name="ACKS_CELL_DATA",
            font=_font(FONT_BODY, 10, color=XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_BAND",
            fill=_fill(XLSX_BAND),
            font=_font(FONT_BODY, 10, color=XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        # TOTAL
        NamedStyle(
            name="ACKS_CELL_TOTAL",
            fill=_fill(XLSX_TOTAL_BG),
            font=_font(FONT_HEADER, 10, True, XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center"),
            border=Border(
                top=Side(style="medium", color=XLSX_TEXT_PRIMARY),
                bottom=Side(style="thin",   color=XLSX_GRID),
                left=Side(style="thin",     color=XLSX_GRID),
                right=Side(style="thin",    color=XLSX_GRID),
            ),
            number_format=FMT_INT,
        ),
        # KPI
        NamedStyle(
            name="ACKS_CELL_KPI_LABEL",
            fill=_fill(XLSX_HEADER_BG),
            font=_font(FONT_NUMERIC, 9, True, XLSX_BG_PRIMARY),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_KPI_VALUE",
            font=_font(FONT_KPI, 22, True, XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="right", vertical="center", indent=1),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
            number_format=FMT_INT,
        ),
        NamedStyle(
            name="ACKS_CELL_KPI_POS",
            font=_font(FONT_NUMERIC, 10, True, XLSX_GOOD),
            alignment=Alignment(horizontal="right", vertical="center", indent=1),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
            number_format=FMT_PCT_DELTA,
        ),
        NamedStyle(
            name="ACKS_CELL_KPI_NEG",
            font=_font(FONT_NUMERIC, 10, True, XLSX_BAD),
            alignment=Alignment(horizontal="right", vertical="center", indent=1),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
            number_format=FMT_PCT_DELTA,
        ),
        # SEMANTIC
        NamedStyle(
            name="ACKS_CELL_ACCENT",
            fill=_fill(XLSX_PRIMARY_SOFT),
            font=_font(FONT_NUMERIC, 10, True, XLSX_PRIMARY_DEEP),
            alignment=Alignment(horizontal="right", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_GOOD",
            fill=_fill(XLSX_GOOD_SOFT),
            font=_font(FONT_NUMERIC, 10, True, XLSX_GOOD),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        NamedStyle(
            name="ACKS_CELL_WARN",
            fill=_fill(XLSX_BAD_SOFT),
            font=_font(FONT_NUMERIC, 10, True, XLSX_BAD),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
        # MODEL
        NamedStyle(
            name="ACKS_CELL_INPUT",
            fill=_fill(XLSX_INPUT_SOFT),
            font=_font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY),
            alignment=Alignment(horizontal="right", vertical="center"),
            border=Border(
                top=Side(style="dashed",    color=XLSX_INPUT),
                bottom=Side(style="dashed", color=XLSX_INPUT),
                left=Side(style="dashed",   color=XLSX_INPUT),
                right=Side(style="dashed",  color=XLSX_INPUT),
            ),
        ),
        NamedStyle(
            name="ACKS_CELL_FORMULA",
            fill=_fill(XLSX_BAND),
            font=_font(FONT_NUMERIC, 10, color=XLSX_TEXT_SECONDARY, italic=True),
            alignment=Alignment(horizontal="right", vertical="center"),
            border=_border(top="thin", bottom="thin", left="thin", right="thin"),
        ),
    ]

    existing = set()
    for s in wb.named_styles:
        if hasattr(s, 'name'):
            existing.add(s.name)
        elif isinstance(s, str):
            existing.add(s)
    for st in styles:
        if st.name not in existing:
            wb.add_named_style(st)


# =============================================================
#                    SHEET HELPERS
# =============================================================

def _setup_sheet(ws: Worksheet, *, tab_color: str = XLSX_TEXT_PRIMARY,
                 zoom: int = 110, freeze: Optional[str] = None,
                 hide_gridlines: bool = True):
    ws.sheet_view.showGridLines = not hide_gridlines
    ws.sheet_view.zoomScale = zoom
    ws.sheet_properties.tabColor = tab_color
    if freeze:
        ws.freeze_panes = freeze

def _set_widths(ws: Worksheet, widths: Sequence[float]):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _set_height(ws: Worksheet, row: int, height: float):
    ws.row_dimensions[row].height = height


# =============================================================
#                    SHEET TEMPLATES
# =============================================================

def build_readme_sheet(ws: Worksheet,
                       meta_rows: Optional[Sequence[tuple]] = None,
                       conventions: Optional[Sequence[tuple]] = None) -> None:
    """README sheet — workbook documentation."""
    _setup_sheet(ws, tab_color=XLSX_TEXT_SECONDARY, zoom=110)
    _set_widths(ws, [4, 28, 80])

    ws.merge_cells("B2:C2")
    ws["B2"] = "ACKS Studio · Workbook README"
    ws["B2"].style = "ACKS_CELL_H1"
    _set_height(ws, 2, 32)

    if meta_rows is None:
        meta_rows = [
            ("Version",      "2.0.0"),
            ("Issued",       "2026.05.18"),
            ("Author",       "ACKS Studio · Strategy & Finance"),
            ("Source",       "Data!A:Z · Pulled from BI on 2026.05.16"),
            ("Refresh",      "T+1 自动核对 · 每周一 09:00 全量复核"),
            ("Tab Order",    "README → Dashboard → Data → Model"),
            ("Sign-off",     "S. Lu (CEO) · L. Yang (Strategy)"),
        ]
    for i, (k, v) in enumerate(meta_rows, start=4):
        ws.cell(row=i, column=2, value=k).style = "ACKS_CELL_H3"
        ws.cell(row=i, column=3, value=v).style = "ACKS_CELL_DATA"
        _set_height(ws, i, 22)

    _set_height(ws, 12, 12)

    ws["B13"] = "Conventions · 约定"
    ws["B13"].style = "ACKS_CELL_H1"
    ws.merge_cells("B13:C13")
    _set_height(ws, 13, 28)

    if conventions is None:
        conventions = [
            ("Input cells",   "琥珀色虚线（ACKS_CELL_INPUT）— 唯一可手改区域"),
            ("Formula cells", "灰色斜体（ACKS_CELL_FORMULA）— 不要覆盖"),
            ("Total rows",    "顶部 2pt 黑色实线 + 浅灰底（ACKS_CELL_TOTAL）"),
            ("Delta arrows",  "▲ 正向（绿）· ▼ 负向（红）· 数字必须等宽对齐"),
            ("Tab colors",    "Dashboard = 橙 · Data = 黑 · Model = 灰"),
        ]
    for i, (k, v) in enumerate(conventions, start=14):
        ws.cell(row=i, column=2, value=k).style = "ACKS_CELL_H3"
        ws.cell(row=i, column=3, value=v).style = "ACKS_CELL_DATA"
        _set_height(ws, i, 22)


def build_dashboard_sheet(
    ws: Worksheet,
    kpis: Optional[Sequence[tuple]] = None,
    revenue_rows: Optional[Sequence[tuple]] = None,
) -> None:
    """Dashboard sheet — KPI cards + revenue breakdown.

    Args:
        kpis: List of (label, value, delta_string) tuples.
        revenue_rows: List of (no, name, y1, y2, yoy, status) tuples.
    """
    _setup_sheet(ws, tab_color=XLSX_PRIMARY, zoom=110, freeze="A6")
    _set_widths(ws, [4, 22, 14, 14, 14, 14, 14])

    ws.merge_cells("B2:G2")
    ws["B2"] = "ACKS Studio · Q1 2026 Dashboard"
    ws["B2"].style = "ACKS_CELL_H1"
    _set_height(ws, 2, 34)

    if kpis is None:
        kpis = [
            ("TOTAL REVENUE", "¥420.0M", "+0.220"),
            ("GROSS MARGIN",  "38.6%",   "+0.024"),
            ("ACTIVE CLIENTS","1,284",   "+0.18"),
            ("NPS SCORE",     "62",      "+1"),
        ]
    cols = ["C", "D", "E", "F"]
    for i, (lbl, val, dlt) in enumerate(kpis):
        col = cols[i]
        ws[f"{col}4"] = lbl
        ws[f"{col}4"].style = "ACKS_CELL_KPI_LABEL"
        ws[f"{col}5"] = val
        ws[f"{col}5"].style = "ACKS_CELL_KPI_VALUE"
        ws[f"{col}6"] = dlt
        ws[f"{col}6"].style = "ACKS_CELL_KPI_POS"
    _set_height(ws, 4, 22)
    _set_height(ws, 5, 42)
    _set_height(ws, 6, 22)

    title_row = 9
    ws.merge_cells(f"B{title_row}:G{title_row}")
    ws[f"B{title_row}"] = "分业务线营收 · Revenue by Business Line"
    ws[f"B{title_row}"].style = "ACKS_CELL_H1"
    _set_height(ws, title_row, 28)

    headers = ["No.", "业务线", "2024 (亿)", "2025 (亿)", "同比", "状态"]
    for i, h in enumerate(headers):
        c = ws.cell(row=title_row + 1, column=2 + i, value=h)
        c.style = "ACKS_CELL_H3"
        if h in ("2024 (亿)", "2025 (亿)", "同比"):
            c.alignment = Alignment(horizontal="right", vertical="center")

    if revenue_rows is None:
        revenue_rows = [
            ("01", "性能改装",  462, 588, 0.273,  "达标"),
            ("02", "空气动力",  218, 274, 0.257,  "达标"),
            ("03", "制动系统",  156, 183, 0.173,  "达标"),
            ("04", "轮毂悬挂",  312, 355, 0.138,  "未达标"),
        ]
    start_row = title_row + 2
    for i, (no, name, y1, y2, yoy, status) in enumerate(revenue_rows):
        r = start_row + i
        even = (i % 2 == 1)
        base = "ACKS_CELL_BAND" if even else "ACKS_CELL_DATA"

        ws.cell(row=r, column=2, value=no).style = base
        ws.cell(row=r, column=3, value=name).style = base

        c4 = ws.cell(row=r, column=4, value=y1); c4.style = base
        c4.number_format = FMT_INT
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)

        c5 = ws.cell(row=r, column=5, value=y2); c5.style = base
        c5.number_format = FMT_INT
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)

        c6 = ws.cell(row=r, column=6, value=yoy)
        if i == 0:
            c6.style = "ACKS_CELL_ACCENT"
        else:
            c6.style = "ACKS_CELL_KPI_POS"
        c6.number_format = FMT_PCT_DELTA

        c7 = ws.cell(row=r, column=7, value=("✓ " + status if status == "达标"
                                              else "✗ " + status))
        c7.style = "ACKS_CELL_GOOD" if status == "达标" else "ACKS_CELL_WARN"
        _set_height(ws, r, 22)

    # Total row
    tr = start_row + len(revenue_rows)
    ws.cell(row=tr, column=2, value="—").style = "ACKS_CELL_TOTAL"
    ws.cell(row=tr, column=3, value="合计").style = "ACKS_CELL_TOTAL"
    cT4 = ws.cell(row=tr, column=4, value=sum(r[2] for r in revenue_rows))
    cT4.style = "ACKS_CELL_TOTAL"; cT4.number_format = FMT_INT
    cT4.alignment = Alignment(horizontal="right", vertical="center")
    cT5 = ws.cell(row=tr, column=5, value=sum(r[3] for r in revenue_rows))
    cT5.style = "ACKS_CELL_TOTAL"; cT5.number_format = FMT_INT
    cT5.alignment = Alignment(horizontal="right", vertical="center")
    cT6 = ws.cell(row=tr, column=6, value=0.220)
    cT6.style = "ACKS_CELL_TOTAL"
    cT6.number_format = FMT_PCT_DELTA
    cT6.font = _font(FONT_NUMERIC, 10, True, XLSX_PRIMARY_DEEP)
    cT6.alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=tr, column=7, value="—").style = "ACKS_CELL_TOTAL"
    _set_height(ws, tr, 24)

    # Conditional formatting
    yoy_range = f"F{start_row}:F{start_row + len(revenue_rows) - 1}"
    ws.conditional_formatting.add(
        yoy_range,
        ColorScaleRule(
            start_type="num",  start_value=-0.10, start_color=XLSX_BAD_SOFT,
            mid_type="num",    mid_value=0.0,     mid_color=XLSX_BG_PRIMARY,
            end_type="num",    end_value=0.30,    end_color=XLSX_GOOD_SOFT,
        ),
    )
    y2_range = f"E{start_row}:E{start_row + len(revenue_rows) - 1}"
    ws.conditional_formatting.add(
        y2_range,
        DataBarRule(start_type="min", end_type="max", color=XLSX_PRIMARY,
                    showValue=True),
    )


def build_data_sheet(ws: Worksheet,
                     headers: Optional[Sequence[str]] = None,
                     rows: Optional[Sequence[tuple]] = None) -> None:
    """Flat data table for pivots and ETL.

    Args:
        headers: Column headers.
        rows: Data tuples matching header count.
    """
    _setup_sheet(ws, tab_color=XLSX_TEXT_PRIMARY, zoom=110, freeze="A2")
    _set_widths(ws, [12, 14, 12, 10, 12, 14, 12, 12])

    if headers is None:
        headers = ["报告期", "业务线", "客户类型", "区域",
                   "单数", "收入 (M)", "毛利率", "状态"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.style = "ACKS_CELL_H1"
        if h in ("单数", "收入 (M)", "毛利率"):
            c.alignment = Alignment(horizontal="right", vertical="center")
    _set_height(ws, 1, 24)

    if rows is None:
        rows = [
            ("2026 Q1", "性能改装",   "B2C", "华南", 186, 48.2, 0.412, "达标"),
            ("2026 Q1", "性能改装",   "B2B", "华南",  42, 18.4, 0.381, "达标"),
            ("2026 Q1", "空气动力",   "B2C", "华东",  98, 22.1, 0.356, "达标"),
            ("2026 Q1", "制动系统",   "B2C", "华北",  54,  9.8, 0.284, "未达标"),
            ("2026 Q1", "轮毂悬挂",   "B2C", "华东", 126, 31.5, 0.328, "未达标"),
            ("2026 Q1", "轮毂悬挂",   "B2C", "华南", 142, 35.8, 0.340, "达标"),
            ("2026 Q1", "AI 智能体",  "B2B", "全国",  23, 48.0, 0.684, "达标"),
            ("2026 Q1", "软件服务",   "B2B", "全国",  14, 54.0, 0.421, "达标"),
        ]
    for i, row in enumerate(rows, start=2):
        even = (i % 2 == 1)
        base = "ACKS_CELL_BAND" if even else "ACKS_CELL_DATA"
        for col, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=col, value=val)
            if col == 5:
                c.style = base; c.number_format = FMT_INT
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)
            elif col == 6:
                if row[1] == "AI 智能体":
                    c.style = "ACKS_CELL_ACCENT"
                else:
                    c.style = base
                    c.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)
                c.number_format = '"¥"#,##0.0"M"'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col == 7:
                c.style = base; c.number_format = FMT_PCT
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)
            elif col == 8:
                c.style = ("ACKS_CELL_GOOD" if val == "达标"
                           else "ACKS_CELL_WARN")
                c.value = ("✓ " + val if val == "达标" else "✗ " + val)
            else:
                c.style = base
        _set_height(ws, i, 22)

    # Conditional formatting on margin column
    margin_range = f"G2:G{len(rows) + 1}"
    ws.conditional_formatting.add(
        margin_range,
        CellIsRule(
            operator="lessThan",
            formula=["0.30"],
            fill=PatternFill("solid", fgColor=XLSX_BAD_SOFT),
            font=Font(color=XLSX_BAD, name=FONT_NUMERIC, size=10, bold=True),
        ),
    )


def build_model_sheet(ws: Worksheet,
                      model_rows: Optional[Sequence[tuple]] = None) -> None:
    """Financial forecast model with input + formula regions.

    Args:
        model_rows: List of (name, growth_rate, q1_actual) tuples.
    """
    _setup_sheet(ws, tab_color=XLSX_TEXT_TERTIARY, zoom=110, freeze="B5")
    _set_widths(ws, [4, 18, 14, 14, 14, 14, 14, 16])

    ws.merge_cells("B2:H2")
    ws["B2"] = "营收预测模型 · Q2 2026 — Q4 2026"
    ws["B2"].style = "ACKS_CELL_H1"
    _set_height(ws, 2, 28)

    ws.merge_cells("B3:C3")
    ws["B3"] = "输入 INPUT"
    ws["B3"].style = "ACKS_CELL_H3"
    ws.merge_cells("D3:H3")
    ws["D3"] = "输出 OUTPUT (FORMULA)"
    ws["D3"].style = "ACKS_CELL_H3"

    headers = ["业务线", "增速假设", "Q1 实际",
               "Q2 预测", "Q3 预测", "Q4 预测", "2026 全年"]
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=4, column=i, value=h)
        c.style = "ACKS_CELL_H3"
        if i >= 3:
            c.alignment = Alignment(horizontal="right", vertical="center")
    _set_height(ws, 4, 24)

    if model_rows is None:
        model_rows = [
            ("性能改装", 0.080, 186.0),
            ("空气动力", 0.065,  87.0),
            ("制动系统", 0.040,  54.0),
            ("轮毂悬挂", 0.035,  93.0),
        ]
    start_row = 5
    for i, (name, growth, q1) in enumerate(model_rows):
        r = start_row + i
        even = (i % 2 == 1)
        base = "ACKS_CELL_BAND" if even else "ACKS_CELL_DATA"

        ws.cell(row=r, column=2, value=name).style = base
        c_inp = ws.cell(row=r, column=3, value=growth)
        c_inp.style = "ACKS_CELL_INPUT"
        c_inp.number_format = FMT_PCT_DELTA
        c_q1 = ws.cell(row=r, column=4, value=q1)
        c_q1.style = base
        c_q1.number_format = '"¥"#,##0.0"M"'
        c_q1.alignment = Alignment(horizontal="right", vertical="center")
        c_q1.font = _font(FONT_NUMERIC, 10, color=XLSX_TEXT_PRIMARY)
        for j, col in enumerate(("E", "F", "G"), start=1):
            cell = ws[f"{col}{r}"]
            cell.value = f"=D{r}*(1+C{r})^{j}"
            cell.style = "ACKS_CELL_FORMULA"
            cell.number_format = '"¥"#,##0.0"M"'
        c_year = ws[f"H{r}"]
        c_year.value = f"=D{r}+E{r}+F{r}+G{r}"
        c_year.style = "ACKS_CELL_TOTAL"
        c_year.number_format = '"¥"#,##0.0"M"'
        c_year.alignment = Alignment(horizontal="right", vertical="center")
        _set_height(ws, r, 22)

    tr = start_row + len(model_rows)
    ws.cell(row=tr, column=2, value="合计").style = "ACKS_CELL_TOTAL"
    avg_cell = ws.cell(row=tr, column=3, value=f"=AVERAGE(C{start_row}:C{tr - 1})")
    avg_cell.style = "ACKS_CELL_TOTAL"
    avg_cell.number_format = FMT_PCT_DELTA
    avg_cell.alignment = Alignment(horizontal="right", vertical="center")
    avg_cell.font = _font(FONT_NUMERIC, 10, True, XLSX_TEXT_PRIMARY)
    for col_letter in ("D", "E", "F", "G", "H"):
        cell = ws[f"{col_letter}{tr}"]
        cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{tr - 1})"
        cell.style = "ACKS_CELL_TOTAL"
        cell.number_format = '"¥"#,##0.0"M"'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        if col_letter == "H":
            cell.font = _font(FONT_NUMERIC, 11, True, XLSX_PRIMARY_DEEP)
        else:
            cell.font = _font(FONT_NUMERIC, 10, True, XLSX_TEXT_PRIMARY)
    _set_height(ws, tr, 24)

    note_r = tr + 2
    ws.merge_cells(f"B{note_r}:H{note_r}")
    note_cell = ws[f"B{note_r}"]
    note_cell.value = ("注：黄色虚线为可调输入；灰色斜体为公式，请勿手改。"
                       "如需调整模型结构，请联系 L. Yang。")
    note_cell.font = _font(FONT_NOTE, 9, italic=True, color=XLSX_TEXT_TERTIARY)
    note_cell.alignment = Alignment(horizontal="left", vertical="center")
    _set_height(ws, note_r, 24)


# =============================================================
#                    DEMO BUILD
# =============================================================

def build_workbook() -> Workbook:
    """Construct the full ACKS demo workbook."""
    wb = Workbook()
    register_named_styles(wb)

    readme = wb.active
    readme.title = "README"
    build_readme_sheet(readme)

    build_dashboard_sheet(wb.create_sheet("Dashboard"))
    build_data_sheet(wb.create_sheet("Data"))
    build_model_sheet(wb.create_sheet("Model"))

    return wb


def save_demo(path: str = "ACKS-workbook.xlsx") -> str:
    wb = build_workbook()
    wb.save(path)
    return path


if __name__ == "__main__":
    out = save_demo()
    print(f"Saved → {out}")
