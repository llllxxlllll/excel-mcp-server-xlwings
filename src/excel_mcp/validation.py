# -*- coding: utf-8 -*-
"""验证模块 - 使用 xlwings 实现

本模块提供公式验证和范围验证功能。
"""

import logging
import re
from typing import Any, Tuple

from .xw_helper import get_workbook, get_sheet, column_string_from_index, WorkbookContext
from .cell_utils import parse_cell_range, validate_cell_reference
from .exceptions import ValidationError

logger = logging.getLogger(__name__)


def validate_formula(formula: str) -> Tuple[bool, str]:
    """验证 Excel 公式语法和安全性
    
    Args:
        formula: Excel 公式字符串
        
    Returns:
        (is_valid, message) 元组
    """
    if not formula.startswith("="):
        return False, "Formula must start with '='"

    # 移除 '=' 前缀进行验证
    formula_body = formula[1:]

    # 检查括号是否平衡
    parens = 0
    for c in formula_body:
        if c == "(":
            parens += 1
        elif c == ")":
            parens -= 1
        if parens < 0:
            return False, "Unmatched closing parenthesis"

    if parens > 0:
        return False, "Unclosed parenthesis"

    # 基本函数名验证
    func_pattern = r"([A-Z]+)\("
    funcs = re.findall(func_pattern, formula_body)
    unsafe_funcs = {"INDIRECT", "HYPERLINK", "WEBSERVICE", "DGET", "RTD"}

    for func in funcs:
        if func in unsafe_funcs:
            return False, f"Unsafe function: {func}"

    return True, "Formula is valid"


def validate_formula_in_cell_operation(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> dict[str, Any]:
    """验证 Excel 公式并与单元格内容比较
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 工作表名称
        cell: 单元格引用
        formula: 要验证的公式
        
    Returns:
        包含验证结果的字典
        
    Raises:
        ValidationError: 当验证失败时
    """
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)

            # 验证单元格引用
            if not validate_cell_reference(cell):
                raise ValidationError(f"Invalid cell reference: {cell}")

            # 验证公式语法
            is_valid, message = validate_formula(formula)
            if not is_valid:
                raise ValidationError(f"Invalid formula syntax: {message}")

            # 验证公式中的单元格引用
            cell_refs = re.findall(r'[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?', formula)
            for ref in cell_refs:
                if ':' in ref:  # 范围引用
                    start, end = ref.split(':')
                    if not (validate_cell_reference(start) and validate_cell_reference(end)):
                        raise ValidationError(f"Invalid cell range reference in formula: {ref}")
                else:  # 单个单元格引用
                    if not validate_cell_reference(ref):
                        raise ValidationError(f"Invalid cell reference in formula: {ref}")

            # 获取单元格当前公式
            cell_range = sheet.range(cell)
            current_formula = cell_range.formula

            # 如果单元格有公式（以 = 开头）
            if isinstance(current_formula, str) and current_formula.startswith('='):
                # 标准化公式进行比较
                provided = formula if formula.startswith('=') else f"={formula}"
                
                if current_formula != provided:
                    return {
                        "message": "Formula is valid but doesn't match cell content",
                        "valid": True,
                        "matches": False,
                        "cell": cell,
                        "provided_formula": formula,
                        "current_formula": current_formula
                    }
                else:
                    return {
                        "message": "Formula is valid and matches cell content",
                        "valid": True,
                        "matches": True,
                        "cell": cell,
                        "formula": formula
                    }
            else:
                return {
                    "message": "Formula is valid but cell contains no formula",
                    "valid": True,
                    "matches": False,
                    "cell": cell,
                    "provided_formula": formula,
                    "current_content": str(current_formula) if current_formula else ""
                }

    except ValidationError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to validate formula: {e}")
        raise ValidationError(str(e))


def validate_range_in_sheet_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> dict[str, Any]:
    """验证范围是否存在于工作表中并返回数据范围信息
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格
        end_cell: 结束单元格（可选）
        
    Returns:
        包含验证结果的字典
        
    Raises:
        ValidationError: 当验证失败时
    """
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 获取实际数据维度
            used_range = sheet.used_range
            if used_range is not None:
                data_max_row = used_range.last_cell.row
                data_max_col = used_range.last_cell.column
            else:
                data_max_row = 1
                data_max_col = 1
            
            # 验证范围
            try:
                start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            except ValueError as e:
                raise ValidationError(f"Invalid range: {str(e)}")
            
            # 如果未指定结束位置，使用起始位置
            if end_row is None:
                end_row = start_row
            if end_col is None:
                end_col = start_col
            
            # 验证边界
            is_valid, message = validate_range_bounds(
                sheet, start_row, start_col, end_row, end_col
            )
            if not is_valid:
                raise ValidationError(message)
            
            range_str = f"{start_cell}" if end_cell is None else f"{start_cell}:{end_cell}"
            data_range_str = f"A1:{column_string_from_index(data_max_col)}{data_max_row}"
            
            # 检查范围是否超出数据区域
            extends_beyond_data = (
                end_row > data_max_row or 
                end_col > data_max_col
            )
            
            return {
                "message": (
                    f"Range '{range_str}' is valid. "
                    f"Sheet contains data in range '{data_range_str}'"
                ),
                "valid": True,
                "range": range_str,
                "data_range": data_range_str,
                "extends_beyond_data": extends_beyond_data,
                "data_dimensions": {
                    "max_row": data_max_row,
                    "max_col": data_max_col,
                    "max_col_letter": column_string_from_index(data_max_col)
                }
            }
    except ValidationError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to validate range: {e}")
        raise ValidationError(str(e))


def validate_range_bounds(
    sheet,
    start_row: int,
    start_col: int,
    end_row: int | None = None,
    end_col: int | None = None,
) -> Tuple[bool, str]:
    """验证单元格范围是否在工作表边界内
    
    Args:
        sheet: xlwings Sheet 对象
        start_row: 起始行（1-based）
        start_col: 起始列（1-based）
        end_row: 结束行（可选）
        end_col: 结束列（可选）
        
    Returns:
        (is_valid, message) 元组
    """
    # 获取工作表的已使用范围
    used_range = sheet.used_range
    if used_range is not None:
        max_row = used_range.last_cell.row
        max_col = used_range.last_cell.column
    else:
        max_row = 1
        max_col = 1

    try:
        # 检查起始单元格边界
        if start_row < 1 or start_row > max_row:
            return False, f"Start row {start_row} out of bounds (1-{max_row})"
        if start_col < 1 or start_col > max_col:
            return False, (
                f"Start column {column_string_from_index(start_col)} "
                f"out of bounds (A-{column_string_from_index(max_col)})"
            )

        # 如果指定了结束单元格，检查其边界
        if end_row is not None and end_col is not None:
            if end_row < start_row:
                return False, "End row cannot be before start row"
            if end_col < start_col:
                return False, "End column cannot be before start column"
            if end_row > max_row:
                return False, f"End row {end_row} out of bounds (1-{max_row})"
            if end_col > max_col:
                return False, (
                    f"End column {column_string_from_index(end_col)} "
                    f"out of bounds (A-{column_string_from_index(max_col)})"
                )

        return True, "Range is valid"
    except Exception as e:
        return False, f"Invalid range: {e!s}"
