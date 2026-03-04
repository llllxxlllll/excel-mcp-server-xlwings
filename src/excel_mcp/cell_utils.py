# -*- coding: utf-8 -*-
"""单元格工具模块

提供单元格引用解析和验证功能。
"""

import re

from .xw_helper import column_index_from_string


def parse_cell_range(
    cell_ref: str,
    end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """解析 Excel 单元格引用为行列索引
    
    Args:
        cell_ref: 起始单元格引用，如 'A1'
        end_ref: 结束单元格引用，如 'B10'（可选）
        
    Returns:
        (start_row, start_col, end_row, end_col) 元组
        如果没有 end_ref，则 end_row 和 end_col 为 None
    """
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    else:
        start_cell = cell_ref
        end_cell = None

    match = re.match(r"([A-Z]+)([0-9]+)", start_cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {start_cell}")
    col_str, row_str = match.groups()
    start_row = int(row_str)
    start_col = column_index_from_string(col_str)

    if end_cell:
        match = re.match(r"([A-Z]+)([0-9]+)", end_cell.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {end_cell}")
        col_str, row_str = match.groups()
        end_row = int(row_str)
        end_col = column_index_from_string(col_str)
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col


def validate_cell_reference(cell_ref: str) -> bool:
    """验证 Excel 单元格引用格式
    
    Args:
        cell_ref: 单元格引用，如 'A1', 'BC123'
        
    Returns:
        如果格式有效返回 True，否则返回 False
    """
    if not cell_ref:
        return False

    # 分离列和行部分
    col = row = ""
    for c in cell_ref:
        if c.isalpha():
            if row:  # 数字后面不能有字母
                return False
            col += c
        elif c.isdigit():
            row += c
        else:
            return False

    return bool(col and row)
