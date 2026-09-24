
import datetime
import os
from typing import Optional, Union, List, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
import pandas as pd

def create_excel(title: str, data: List[List[Any]], output_path: str,
                 create_chart: bool = False,
                 theme: str = "acks", **kwargs) -> Dict[str, Any]:
    """
    创建 Excel 表格。

    Args:
        title: 表格标题
        data: 二维数组数据，第一行为表头
        output_path: 输出路径
        create_chart: 是否生成柱状图
        theme: "acks"（符合 ACKS 设计规范，默认）或 "default"（简单样式）
    """
    if theme == "acks":
        return _create_excel_acks(title, data, output_path, create_chart, **kwargs)
    return _create_excel_default(title, data, output_path, create_chart, **kwargs)


def _create_excel_acks(title: str, data: List[List[Any]], output_path: str,
                       create_chart: bool, **kwargs) -> Dict[str, Any]:
    """用 ACKS 设计规范（design_system.xlsx）生成 Excel。"""
    from .design_system import xlsx as acks

    wb = openpyxl.Workbook()
    acks.register_named_styles(wb)
    ws = wb.active
    ws.title = title[:30]

    headers = list(data[0]) if data else []
    rows = [tuple(r) for r in data[1:]] if len(data) > 1 else []
    acks.build_data_sheet(ws, headers=headers, rows=rows)

    wb.save(output_path)

    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "rows": len(data),
        "columns": len(data[0]) if data else 0,
        "theme": "acks",
    }


def _create_excel_default(title: str, data: List[List[Any]], output_path: str,
                          create_chart: bool, **kwargs) -> Dict[str, Any]:
    """原简单样式（蓝色表头 + 柱状图）。"""
    # 创建工作簿和工作表
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title[:30]  # 工作表名最长30字符
    
    # 定义样式
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # 写入数据
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            
            # 表头样式
            if row_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
            else:
                cell.alignment = left_align if col_idx == 1 else center_align
            
            cell.border = thin_border
    
    # 自动调整列宽
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = min(adjusted_width, 50)  # 最大宽度50
    
    # 创建图表
    if create_chart and len(data) >= 2:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = title
        chart.y_axis.title = '数值'
        chart.x_axis.title = data[0][0]
        
        # 数据范围：第二列开始作为数值列
        data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(data), max_col=len(data[0]))
        categories = Reference(ws, min_col=1, min_row=2, max_row=len(data))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(categories)
        
        # 把图表放到数据下方
        ws.add_chart(chart, f"A{len(data) + 3}")
    
    # 保存文件
    wb.save(output_path)
    
    return {
        "output_path": output_path,
        "file_size": os.path.getsize(output_path),
        "rows": len(data),
        "columns": len(data[0]) if data else 0,
        "theme": "default",
    }

def extract_data(input_path: str, sheet_name: Optional[Union[str, int]] = None,
                 **kwargs) -> List[Dict[str, Any]]:
    """
    从Excel提取单个工作表的数据
    Args:
        input_path: 输入文件路径（.xlsx / .xls）
        sheet_name: 工作表名或从 0 开始的索引；不传或传 None 时读取第一个工作表
        **kwargs: 透传给 pandas.read_excel，如 header、dtype、usecols；
            未传 dtype / converters 时默认 dtype=object，即按单元格原值读取，不做类型推断
    Returns:
        字典列表，每一行是一个字典，key是表头。值可直接 JSON 序列化：
        空单元格为 None，日期/时间/时长为 ISO 8601 字符串
    """
    if isinstance(sheet_name, list):
        raise TypeError("sheet_name 只支持单个工作表名或索引，多个工作表请逐个读取")
    # 不让 pandas 推断列类型：否则工号 "00123" 会变成 123、身份证号丢精度，空单元格还会把整列 int/bool 变成 float。
    # 调用方自行指定 converters 时不加默认值，否则 pandas 会误报「同时指定了 converter 和 dtype」
    if "dtype" not in kwargs and "converters" not in kwargs:
        kwargs["dtype"] = object
    # pandas 的 sheet_name=None 表示读取全部工作表并返回 dict，这里固定为读取第一个工作表
    df = pd.read_excel(input_path, sheet_name=0 if sheet_name is None else sheet_name, **kwargs)
    return [{_to_json_value(k): _to_json_value(v) for k, v in row.items()}
            for row in df.to_dict('records')]


def _to_json_value(value: Any) -> Any:
    """单元格值转为可 JSON 序列化的值：空值 → None，日期/时间/时长 → ISO 8601 字符串。"""
    # 空值判断要放在最前：NaT 也是 datetime 的子类
    if pd.api.types.is_scalar(value) and pd.isna(value):
        return None
    if isinstance(value, datetime.timedelta):
        return pd.Timedelta(value).isoformat()
    if isinstance(value, (datetime.date, datetime.time)):  # 含 datetime 与 pd.Timestamp
        return value.isoformat()
    return value

def from_dataframe(df: pd.DataFrame, output_path: str, **kwargs) -> Dict[str, Any]:
    """
    从pandas DataFrame创建Excel
    """
    df.to_excel(output_path, index=False, **kwargs)
    
    # 美化表格
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    # 设置表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
    
    # 调整列宽
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[column].width = max_length + 2
    
    wb.save(output_path)
    
    return {"output_path": output_path, "rows": len(df), "columns": len(df.columns)}
