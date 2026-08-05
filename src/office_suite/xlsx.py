
import os
from typing import Optional, List, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
import pandas as pd

def create_excel(title: str, data: List[List[Any]], output_path: str, create_chart: bool = False, **kwargs) -> Dict[str, Any]:
    """
    创建Excel表格
    Args:
        title: 表格标题
        data: 二维数组数据，第一行为表头
        output_path: 输出路径
        create_chart: 是否生成柱状图
    """
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
        "columns": len(data[0]) if data else 0
    }

def extract_data(input_path: str, sheet_name: Optional[str] = None, **kwargs) -> List[Dict[str, Any]]:
    """
    从Excel提取数据
    Args:
        input_path: 输入文件路径
        sheet_name: 工作表名，默认第一个工作表
    Returns:
        字典列表，每一行是一个字典，key是表头
    """
    # 使用pandas读取更方便
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    return df.to_dict('records')

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
