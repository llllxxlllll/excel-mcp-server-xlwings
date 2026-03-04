# -*- coding: utf-8 -*-
"""数据验证模块 - 使用 xlwings 实现

提供获取单元格数据验证信息的功能。
"""

import logging
from typing import Any, Dict, List, Optional

import xlwings as xw

from .xw_helper import get_workbook, get_sheet, WorkbookContext
from .exceptions import SheetError

logger = logging.getLogger(__name__)


def get_data_validation_for_cell(
    filepath: str,
    sheet_name: str,
    cell_address: str
) -> Optional[Dict[str, Any]]:
    """获取指定单元格的数据验证信息
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        cell_address: 单元格地址，如 'A1', 'B2' 等
        
    Returns:
        包含验证元数据的字典，如果没有验证则返回 None
    """
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            cell = sheet.range(cell_address)
            
            # 使用 xlwings API 获取数据验证
            validation = _get_cell_validation_via_api(cell)
            
            if validation:
                return validation
            return None
        
    except Exception as e:
        logger.warning(f"获取单元格 {cell_address} 的验证信息失败: {e}")
        return None


def _get_cell_validation_via_api(cell: xw.Range) -> Optional[Dict[str, Any]]:
    """通过 xlwings API 获取单元格的数据验证信息
    
    Args:
        cell: xlwings Range 对象
        
    Returns:
        验证信息字典或 None
    """
    try:
        # 通过 Excel COM API 访问 Validation 对象
        api = cell.api
        validation = api.Validation
        
        # 检查是否有验证规则
        # Type 属性：1=整数, 2=小数, 3=列表, 4=日期, 5=时间, 6=文本长度, 7=自定义
        val_type = validation.Type
        
        if val_type == 0:  # 无验证
            return None
            
        validation_info = {
            "cell": cell.address.replace("$", ""),
            "has_validation": True,
            "validation_type": _get_validation_type_name(val_type),
            "allow_blank": validation.IgnoreBlank,
        }
        
        # 获取运算符
        try:
            operator = validation.Operator
            validation_info["operator"] = _get_operator_name(operator)
        except Exception:
            pass
        
        # 获取输入提示
        try:
            if validation.ShowInput:
                validation_info["prompt_title"] = validation.InputTitle
                validation_info["prompt"] = validation.InputMessage
        except Exception:
            pass
        
        # 获取错误提示
        try:
            if validation.ShowError:
                validation_info["error_title"] = validation.ErrorTitle
                validation_info["error_message"] = validation.ErrorMessage
        except Exception:
            pass
        
        # 获取公式/值
        try:
            formula1 = validation.Formula1
            if formula1:
                if val_type == 3:  # 列表类型
                    validation_info["allowed_values"] = _extract_list_values(formula1, cell.sheet)
                else:
                    validation_info["formula1"] = formula1
        except Exception:
            pass
        
        try:
            formula2 = validation.Formula2
            if formula2:
                validation_info["formula2"] = formula2
        except Exception:
            pass
        
        return validation_info
        
    except Exception as e:
        # 如果没有验证规则，会抛出异常
        if "Unable to get" in str(e) or "没有" in str(e):
            return None
        logger.debug(f"获取验证信息时出错: {e}")
        return None


def _get_validation_type_name(type_code: int) -> str:
    """将验证类型代码转换为名称
    
    Args:
        type_code: Excel 验证类型代码
        
    Returns:
        验证类型名称
    """
    type_names = {
        1: "whole",      # 整数
        2: "decimal",    # 小数
        3: "list",       # 列表
        4: "date",       # 日期
        5: "time",       # 时间
        6: "textLength", # 文本长度
        7: "custom",     # 自定义
    }
    return type_names.get(type_code, f"unknown({type_code})")


def _get_operator_name(operator_code: int) -> str:
    """将运算符代码转换为名称
    
    Args:
        operator_code: Excel 运算符代码
        
    Returns:
        运算符名称
    """
    operator_names = {
        1: "between",
        2: "notBetween",
        3: "equal",
        4: "notEqual",
        5: "greaterThan",
        6: "lessThan",
        7: "greaterThanOrEqual",
        8: "lessThanOrEqual",
    }
    return operator_names.get(operator_code, f"unknown({operator_code})")


def _extract_list_values(formula: str, sheet: xw.Sheet) -> List[str]:
    """从列表验证公式中提取允许的值
    
    Args:
        formula: 验证公式
        sheet: xlwings Sheet 对象
        
    Returns:
        允许值的列表
    """
    try:
        # 移除引号
        formula = formula.strip('"')
        
        # 处理逗号分隔的列表
        if ',' in formula and ':' not in formula:
            values = [val.strip().strip('"') for val in formula.split(',')]
            return [val for val in values if val]
        
        # 处理范围引用（如 "$A$1:$A$5" 或 "Sheet1!$A$1:$A$5"）
        elif ':' in formula or formula.startswith('$'):
            try:
                # 尝试解析范围并获取值
                range_ref = formula
                if formula.startswith('='):
                    range_ref = formula[1:]
                
                # 使用 xlwings 读取范围值
                range_values = sheet.range(range_ref).value
                
                if range_values is None:
                    return [f"Range: {formula} (empty)"]
                
                # 处理单个值或列表
                if isinstance(range_values, list):
                    # 可能是二维列表
                    actual_values = []
                    for item in range_values:
                        if isinstance(item, list):
                            actual_values.extend([str(v) for v in item if v is not None])
                        elif item is not None:
                            actual_values.append(str(item))
                    return actual_values if actual_values else [f"Range: {formula} (empty)"]
                else:
                    return [str(range_values)]
                    
            except Exception as e:
                logger.warning(f"无法解析范围 '{formula}': {e}")
                return [f"Range: {formula}"]
        
        # 单个值
        else:
            return [formula.strip('"')]
            
    except Exception as e:
        logger.warning(f"解析列表公式 '{formula}' 失败: {e}")
        return [formula]


def get_all_validation_ranges(filepath: str, sheet_name: str) -> List[Dict[str, Any]]:
    """获取工作表中所有的数据验证范围
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        
    Returns:
        包含验证范围信息的字典列表
    """
    validations = []
    
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 通过 Excel COM API 获取所有验证
            # 遍历使用的范围中的单元格，检查每个单元格的验证
            used_range = sheet.used_range
            if used_range is None:
                return validations
            
            # 获取已处理的验证范围，避免重复
            processed_ranges = set()
            
            # 遍历使用范围中的单元格
            for cell in used_range:
                try:
                    api = cell.api
                    validation = api.Validation
                    val_type = validation.Type
                    
                    if val_type == 0:  # 无验证
                        continue
                    
                    # 获取验证应用的范围
                    # 注意：Excel COM API 不直接提供验证范围，
                    # 我们使用单元格地址作为范围标识
                    cell_addr = cell.address.replace("$", "")
                    
                    # 检查是否已处理过相同的验证规则
                    try:
                        formula1 = validation.Formula1
                    except Exception:
                        formula1 = ""
                    
                    range_key = f"{val_type}_{formula1}"
                    if range_key in processed_ranges:
                        continue
                    processed_ranges.add(range_key)
                    
                    validation_info = {
                        "ranges": cell_addr,
                        "validation_type": _get_validation_type_name(val_type),
                        "allow_blank": validation.IgnoreBlank,
                    }
                    
                    if val_type == 3 and formula1:  # 列表类型
                        validation_info["allowed_values"] = _extract_list_values(formula1, sheet)
                    elif formula1:
                        validation_info["formula1"] = formula1
                    
                    validations.append(validation_info)
                    
                except Exception:
                    # 单元格没有验证规则
                    continue
                
    except Exception as e:
        logger.warning(f"获取验证范围失败: {e}")
        
    return validations


def get_data_validation_info(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """获取工作表的数据验证信息（用于 MCP 工具）
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        
    Returns:
        包含所有验证规则的字典
    """
    try:
        validations = get_all_validation_ranges(filepath, sheet_name)
        return {
            "status": "success",
            "sheet_name": sheet_name,
            "validation_count": len(validations),
            "validations": validations
        }
    except SheetError as e:
        return {
            "status": "error",
            "message": str(e)
        }
    except Exception as e:
        logger.error(f"获取数据验证信息失败: {e}")
        return {
            "status": "error",
            "message": f"获取数据验证信息失败: {e}"
        }
