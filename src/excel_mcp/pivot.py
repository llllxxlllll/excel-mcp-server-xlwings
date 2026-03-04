# -*- coding: utf-8 -*-
"""透视表模块 - 使用 xlwings 实现

本模块提供 Excel 透视表创建功能。
"""

from typing import Any, List, Optional
import uuid
import logging

from .xw_helper import get_workbook, get_sheet, save_workbook, column_string_from_index, WorkbookContext
from .data import read_excel_range
from .cell_utils import parse_cell_range
from .exceptions import ValidationError, PivotError

logger = logging.getLogger(__name__)


def create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "sum"
) -> dict[str, Any]:
    """创建透视表
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 包含源数据的工作表名称
        data_range: 源数据范围引用
        rows: 行标签字段
        values: 值字段
        columns: 列标签字段（可选）
        agg_func: 聚合函数 (sum, count, average, max, min)
        
    Returns:
        包含状态消息和透视表维度的字典
    """
    try:
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            sheet = ctx.get_sheet(sheet_name)
            
            # 解析范围
            if ':' not in data_range:
                raise ValidationError("Data range must be in format 'A1:B2'")
            
            try:
                start_cell, end_cell = data_range.split(':')
                start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            except ValueError as e:
                raise ValidationError(f"Invalid data range format: {str(e)}")
            
            if end_row is None or end_col is None:
                raise ValidationError("Invalid data range format: missing end coordinates")
            
            # 创建范围字符串
            data_range_str = f"{column_string_from_index(start_col)}{start_row}:{column_string_from_index(end_col)}{end_row}"
            
            # 清理字段名
            def clean_field_name(field: str) -> str:
                field = str(field).strip()
                for suffix in [" (sum)", " (average)", " (count)", " (min)", " (max)"]:
                    if field.lower().endswith(suffix):
                        return field[:-len(suffix)]
                return field
            
            # 读取源数据
            try:
                data_as_list = read_excel_range(filepath, sheet_name, start_cell, end_cell)
                if not data_as_list or len(data_as_list) < 2:
                    raise PivotError("Source data must have a header row and at least one data row.")
                
                headers = [str(h) for h in data_as_list[0]]
                data = [dict(zip(headers, row)) for row in data_as_list[1:]]
                
                if not data:
                    raise PivotError("No data rows found after header.")
                    
            except Exception as e:
                raise PivotError(f"Failed to read or process source data: {str(e)}")
            
            # 验证聚合函数
            valid_agg_funcs = ["sum", "average", "count", "min", "max"]
            if agg_func.lower() not in valid_agg_funcs:
                raise ValidationError(
                    f"Invalid aggregation function. Must be one of: {', '.join(valid_agg_funcs)}"
                )
            
            # 验证字段名
            if data:
                available_fields_raw = data[0].keys()
                available_fields = {clean_field_name(str(header)).lower() for header in available_fields_raw}
                
                for field_list, field_type in [(rows, "row"), (values, "value")]:
                    for field in field_list:
                        if clean_field_name(str(field)).lower() not in available_fields:
                            raise ValidationError(
                                f"Invalid {field_type} field '{field}'. "
                                f"Available fields: {', '.join(sorted(available_fields_raw))}"
                            )
            
            # 清理字段名
            cleaned_rows = [clean_field_name(field) for field in rows]
            cleaned_values = [clean_field_name(field) for field in values]
            
            # 创建透视表工作表
            pivot_sheet_name = f"{sheet_name}_pivot"
            
            # 删除已存在的透视表工作表
            sheet_names = [s.name for s in wb.sheets]
            if pivot_sheet_name in sheet_names:
                wb.sheets[pivot_sheet_name].delete()
            
            # 创建新工作表
            pivot_ws = wb.sheets.add(pivot_sheet_name, after=wb.sheets[-1])
            
            # 写入表头
            current_col = 1
            for field in cleaned_rows:
                cell = pivot_ws.range((1, current_col))
                cell.value = field
                cell.font.bold = True
                current_col += 1
            
            for field in cleaned_values:
                cell = pivot_ws.range((1, current_col))
                cell.value = f"{field} ({agg_func})"
                cell.font.bold = True
                current_col += 1
            
            # 获取每个行字段的唯一值
            field_values = {}
            for field in cleaned_rows:
                all_values = []
                for record in data:
                    value = str(record.get(field, ''))
                    all_values.append(value)
                field_values[field] = sorted(set(all_values))
            
            # 生成所有行字段值的组合
            row_combinations = _get_combinations(field_values)
            
            # 计算表格维度
            total_rows = len(row_combinations) + 1
            total_cols = len(cleaned_rows) + len(cleaned_values)
            
            # 写入数据行
            current_row = 2
            for combo in row_combinations:
                col = 1
                for field in cleaned_rows:
                    pivot_ws.range((current_row, col)).value = combo[field]
                    col += 1
                
                # 过滤数据
                filtered_data = _filter_data(data, combo, {})
                
                # 计算并写入聚合值
                for value_field in cleaned_values:
                    try:
                        value = _aggregate_values(filtered_data, value_field, agg_func)
                        pivot_ws.range((current_row, col)).value = value
                    except Exception as e:
                        raise PivotError(f"Failed to aggregate values for field '{value_field}': {str(e)}")
                    col += 1
                
                current_row += 1
            
            # 保存工作簿
            ctx.save()
            
            return {
                "message": "Summary table created successfully",
                "details": {
                    "source_range": data_range_str,
                    "pivot_sheet": pivot_sheet_name,
                    "rows": cleaned_rows,
                    "columns": columns or [],
                    "values": cleaned_values,
                    "aggregation": agg_func
                }
            }
        
    except (ValidationError, PivotError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create pivot table: {e}")
        raise PivotError(str(e))


def _get_combinations(field_values: dict[str, set]) -> list[dict]:
    """获取字段值的所有组合"""
    result = [{}]
    for field, values in list(field_values.items()):
        new_result = []
        for combo in result:
            for value in sorted(values):
                new_combo = combo.copy()
                new_combo[field] = value
                new_result.append(new_combo)
        result = new_result
    return result


def _filter_data(data: list[dict], row_filters: dict, col_filters: dict) -> list[dict]:
    """根据行和列过滤器过滤数据"""
    result = []
    for record in data:
        matches = True
        for field, value in row_filters.items():
            if record.get(field) != value:
                matches = False
                break
        for field, value in col_filters.items():
            if record.get(field) != value:
                matches = False
                break
        if matches:
            result.append(record)
    return result


def _aggregate_values(data: list[dict], field: str, agg_func: str) -> float:
    """使用指定函数聚合值"""
    values = [record[field] for record in data if field in record and isinstance(record[field], (int, float))]
    if not values:
        return 0
    
    if agg_func == "sum":
        return sum(values)
    elif agg_func == "average":
        return sum(values) / len(values)
    elif agg_func == "count":
        return len(values)
    elif agg_func == "min":
        return min(values)
    elif agg_func == "max":
        return max(values)
    else:
        return sum(values)
