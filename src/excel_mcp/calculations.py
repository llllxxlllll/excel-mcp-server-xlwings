# -*- coding: utf-8 -*-
"""公式和计算模块 - 使用 xlwings 实现

本模块提供 Excel 公式应用功能。
"""

from typing import Any
import logging

from .xw_helper import get_workbook, get_sheet, save_workbook, WorkbookContext
from .cell_utils import validate_cell_reference
from .exceptions import ValidationError, CalculationError
from .validation import validate_formula

logger = logging.getLogger(__name__)


def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> dict[str, Any]:
    """应用 Excel 公式到单元格
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 工作表名称
        cell: 单元格引用，如 'A1'
        formula: Excel 公式，如 '=SUM(A1:A10)'
        
    Returns:
        包含操作结果的字典
        
    Raises:
        ValidationError: 当参数无效时
        CalculationError: 当公式应用失败时
    """
    try:
        # 验证单元格引用
        if not validate_cell_reference(cell):
            raise ValidationError(f"Invalid cell reference: {cell}")
        
        # 确保公式以 = 开头
        if not formula.startswith('='):
            formula = f'={formula}'
        
        # 验证公式语法
        is_valid, message = validate_formula(formula)
        if not is_valid:
            raise CalculationError(f"Invalid formula syntax: {message}")
        
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            try:
                # 使用 xlwings 应用公式到单元格
                cell_range = sheet.range(cell)
                cell_range.formula = formula
            except Exception as e:
                raise CalculationError(f"Failed to apply formula to cell: {str(e)}")
            
            try:
                # 保存工作簿
                ctx.save()
            except Exception as e:
                raise CalculationError(f"Failed to save workbook after applying formula: {str(e)}")
            
            return {
                "message": f"Applied formula '{formula}' to cell {cell}",
                "cell": cell,
                "formula": formula
            }
        
    except (ValidationError, CalculationError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formula: {e}")
        raise CalculationError(str(e))


def get_formula(
    filepath: str,
    sheet_name: str,
    cell: str
) -> dict[str, Any]:
    """获取单元格中的公式
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 工作表名称
        cell: 单元格引用，如 'A1'
        
    Returns:
        包含公式信息的字典
        
    Raises:
        ValidationError: 当参数无效时
        CalculationError: 当获取公式失败时
    """
    try:
        # 验证单元格引用
        if not validate_cell_reference(cell):
            raise ValidationError(f"Invalid cell reference: {cell}")
        
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 获取单元格
            cell_range = sheet.range(cell)
            
            # 获取公式和值
            formula = cell_range.formula
            value = cell_range.value
            
            # 判断是否有公式
            has_formula = isinstance(formula, str) and formula.startswith('=')
            
            return {
                "cell": cell,
                "has_formula": has_formula,
                "formula": formula if has_formula else None,
                "value": value
            }
        
    except (ValidationError, CalculationError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get formula: {e}")
        raise CalculationError(str(e))



def formula_operation(
    filepath: str,
    sheet_name: str,
    cell: str,
    action: str,
    formula: str = None
) -> dict[str, Any]:
    """统一的公式操作接口
    
    Unified formula operation interface.
    
    Args:
        filepath: Excel文件路径 / Excel file path
        sheet_name: 工作表名称 / Worksheet name
        cell: 单元格引用 / Cell reference
        action: 操作类型 / Action type
        formula: 公式（用于apply/validate）/ Formula (for apply/validate)
        
    Actions / 操作类型:
        - apply: 应用公式 (需要 formula)
                 Apply formula (requires formula)
        - validate: 验证公式语法 (需要 formula)
                    Validate formula syntax (requires formula)
        - get: 获取单元格公式
               Get cell formula
    """
    action = action.lower()
    
    if action == "apply":
        if not formula:
            raise ValidationError("apply 操作需要 formula 参数")
        return apply_formula(filepath, sheet_name, cell, formula)
    
    elif action == "validate":
        if not formula:
            raise ValidationError("validate 操作需要 formula 参数")
        # 确保公式以 = 开头
        if not formula.startswith('='):
            formula = f'={formula}'
        is_valid, message = validate_formula(formula)
        return {
            "message": message if is_valid else f"公式无效: {message}",
            "valid": is_valid,
            "formula": formula
        }
    
    elif action == "get":
        return get_formula(filepath, sheet_name, cell)
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: apply, validate, get"
        )
