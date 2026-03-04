# -*- coding: utf-8 -*-
"""工作表操作模块 - 使用 xlwings 实现

本模块提供工作表管理、合并单元格、行列操作等功能。
使用 WorkbookContext 上下文管理器自动处理文件锁定问题。
"""

import logging
from typing import Any, Dict, List, Optional

from .xw_helper import (
    get_workbook, get_sheet, save_workbook,
    column_string_from_index, column_index_from_string,
    WorkbookContext
)
from .cell_utils import parse_cell_range
from .exceptions import SheetError, ValidationError

logger = logging.getLogger(__name__)


def copy_sheet(filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
    """复制工作表"""
    try:
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            # 检查源工作表是否存在
            sheet_names = [s.name for s in wb.sheets]
            if source_sheet not in sheet_names:
                raise SheetError(f"Source sheet '{source_sheet}' not found")
            
            if target_sheet in sheet_names:
                raise SheetError(f"Target sheet '{target_sheet}' already exists")
            
            # 复制工作表
            source = wb.sheets[source_sheet]
            source.api.Copy(After=wb.sheets[-1].api)
            
            # 重命名新工作表
            new_sheet = wb.sheets[-1]
            new_sheet.name = target_sheet
            
            ctx.save()
            return {"message": f"Sheet '{source_sheet}' copied to '{target_sheet}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to copy sheet: {e}")
        raise SheetError(str(e))


def delete_sheet(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """删除工作表"""
    try:
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            sheet_names = [s.name for s in wb.sheets]
            if sheet_name not in sheet_names:
                raise SheetError(f"Sheet '{sheet_name}' not found")
            
            if len(wb.sheets) == 1:
                raise SheetError("Cannot delete the only sheet in workbook")
            
            # 删除工作表
            wb.sheets[sheet_name].delete()
            
            ctx.save()
            return {"message": f"Sheet '{sheet_name}' deleted"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete sheet: {e}")
        raise SheetError(str(e))


def rename_sheet(filepath: str, old_name: str, new_name: str) -> Dict[str, Any]:
    """重命名工作表"""
    try:
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            sheet_names = [s.name for s in wb.sheets]
            if old_name not in sheet_names:
                raise SheetError(f"Sheet '{old_name}' not found")
            
            if new_name in sheet_names:
                raise SheetError(f"Sheet '{new_name}' already exists")
            
            wb.sheets[old_name].name = new_name
            
            ctx.save()
            return {"message": f"Sheet renamed from '{old_name}' to '{new_name}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to rename sheet: {e}")
        raise SheetError(str(e))


def merge_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """合并单元格范围"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 获取范围并合并
            cell_range = sheet.range(f"{start_cell}:{end_cell}")
            cell_range.merge()
            
            ctx.save()
            return {"message": f"Range '{start_cell}:{end_cell}' merged in sheet '{sheet_name}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to merge range: {e}")
        raise SheetError(str(e))


def unmerge_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """取消合并单元格范围"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 获取范围并取消合并
            cell_range = sheet.range(f"{start_cell}:{end_cell}")
            cell_range.unmerge()
            
            ctx.save()
            return {"message": f"Range '{start_cell}:{end_cell}' unmerged successfully"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to unmerge range: {e}")
        raise SheetError(str(e))


def get_merged_ranges(filepath: str, sheet_name: str) -> List[str]:
    """获取工作表中的合并单元格范围"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 通过 API 获取合并范围
            merged_areas = []
            try:
                for area in sheet.api.UsedRange.MergeArea:
                    merged_areas.append(area.Address.replace('$', ''))
            except Exception:
                # 如果没有合并单元格，返回空列表
                pass
            
            return merged_areas
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get merged cells: {e}")
        raise SheetError(str(e))


def insert_row(filepath: str, sheet_name: str, start_row: int, count: int = 1) -> Dict[str, Any]:
    """插入行"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            if start_row < 1:
                raise ValidationError("Start row must be 1 or greater")
            if count < 1:
                raise ValidationError("Count must be 1 or greater")
            
            # 使用 xlwings API 插入行
            for _ in range(count):
                sheet.range(f"{start_row}:{start_row}").api.Insert()
            
            ctx.save()
            return {"message": f"Inserted {count} row(s) starting at row {start_row} in sheet '{sheet_name}'"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to insert rows: {e}")
        raise SheetError(str(e))


def insert_cols(filepath: str, sheet_name: str, start_col: int, count: int = 1) -> Dict[str, Any]:
    """插入列"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            if start_col < 1:
                raise ValidationError("Start column must be 1 or greater")
            if count < 1:
                raise ValidationError("Count must be 1 or greater")
            
            # 使用 xlwings API 插入列
            col_letter = column_string_from_index(start_col)
            for _ in range(count):
                sheet.range(f"{col_letter}:{col_letter}").api.Insert()
            
            ctx.save()
            return {"message": f"Inserted {count} column(s) starting at column {start_col} in sheet '{sheet_name}'"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to insert columns: {e}")
        raise SheetError(str(e))


def delete_rows(filepath: str, sheet_name: str, start_row: int, count: int = 1) -> Dict[str, Any]:
    """删除行"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            if start_row < 1:
                raise ValidationError("Start row must be 1 or greater")
            if count < 1:
                raise ValidationError("Count must be 1 or greater")
            
            # 使用 xlwings API 删除行
            end_row = start_row + count - 1
            sheet.range(f"{start_row}:{end_row}").api.Delete()
            
            ctx.save()
            return {"message": f"Deleted {count} row(s) starting at row {start_row} in sheet '{sheet_name}'"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete rows: {e}")
        raise SheetError(str(e))


def delete_cols(filepath: str, sheet_name: str, start_col: int, count: int = 1) -> Dict[str, Any]:
    """删除列"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            if start_col < 1:
                raise ValidationError("Start column must be 1 or greater")
            if count < 1:
                raise ValidationError("Count must be 1 or greater")
            
            # 使用 xlwings API 删除列
            start_letter = column_string_from_index(start_col)
            end_letter = column_string_from_index(start_col + count - 1)
            sheet.range(f"{start_letter}:{end_letter}").api.Delete()
            
            ctx.save()
            return {"message": f"Deleted {count} column(s) starting at column {start_col} in sheet '{sheet_name}'"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete columns: {e}")
        raise SheetError(str(e))


def copy_range_operation(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """复制单元格范围到另一位置"""
    try:
        with WorkbookContext(filepath) as ctx:
            source_ws = ctx.get_sheet(sheet_name)
            target_ws = ctx.get_sheet(target_sheet) if target_sheet else source_ws
            
            # 获取源范围
            source_range = source_ws.range(f"{source_start}:{source_end}")
            
            # 获取目标范围
            target_range = target_ws.range(target_start)
            
            # 复制
            source_range.copy(target_range)
            
            ctx.save()
            return {"message": "Range copied successfully"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to copy range: {e}")
        raise SheetError(str(e))


def delete_range_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """删除单元格范围并移动剩余单元格"""
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            if shift_direction not in ["up", "left"]:
                raise ValidationError(f"Invalid shift direction: {shift_direction}. Must be 'up' or 'left'")
            
            # 获取范围
            if end_cell:
                cell_range = sheet.range(f"{start_cell}:{end_cell}")
                range_str = f"{start_cell}:{end_cell}"
            else:
                cell_range = sheet.range(start_cell)
                range_str = start_cell
            
            # 删除并移动
            # xlDeleteShiftUp = -4162, xlDeleteShiftToLeft = -4159
            shift_const = -4162 if shift_direction == "up" else -4159
            cell_range.api.Delete(Shift=shift_const)
            
            ctx.save()
            return {"message": f"Range {range_str} deleted successfully"}
    except (ValidationError, SheetError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete range: {e}")
        raise SheetError(str(e))


def worksheet_operation(
    filepath: str,
    action: str,
    sheet_name: str = None,
    new_name: str = None,
    source_sheet: str = None,
) -> Dict[str, Any]:
    """统一的工作表操作接口
    
    Unified worksheet operation interface.
    
    Args:
        filepath: Excel文件路径 / Excel file path
        action: 操作类型 / Action type (create, copy, delete, rename, list)
        sheet_name: 工作表名称 / Worksheet name
        new_name: 新名称（用于create/rename）/ New name (for create/rename)
        source_sheet: 源工作表（用于copy）/ Source sheet (for copy)
        
    Actions / 操作类型:
        - create: 创建工作表 (需要 new_name)
        - copy: 复制工作表 (需要 source_sheet, new_name)
        - delete: 删除工作表 (需要 sheet_name)
        - rename: 重命名工作表 (需要 sheet_name, new_name)
        - list: 列出所有工作表
    """
    action = action.lower()
    
    if action == "create":
        if not new_name:
            raise ValidationError("create 操作需要 new_name 参数")
        from .workbook import create_sheet
        return create_sheet(filepath, new_name)
    
    elif action == "copy":
        if not source_sheet or not new_name:
            raise ValidationError("copy 操作需要 source_sheet 和 new_name 参数")
        return copy_sheet(filepath, source_sheet, new_name)
    
    elif action == "delete":
        if not sheet_name:
            raise ValidationError("delete 操作需要 sheet_name 参数")
        return delete_sheet(filepath, sheet_name)
    
    elif action == "rename":
        if not sheet_name or not new_name:
            raise ValidationError("rename 操作需要 sheet_name 和 new_name 参数")
        return rename_sheet(filepath, sheet_name, new_name)
    
    elif action == "list":
        with WorkbookContext(filepath) as ctx:
            sheets_info = []
            for i, sheet in enumerate(ctx.wb.sheets):
                sheets_info.append({
                    "name": sheet.name,
                    "index": i
                })
            return {
                "message": f"共 {len(sheets_info)} 个工作表",
                "sheets": sheets_info
            }
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: create, copy, delete, rename, list"
        )


def merge_cell_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_cell: str = None,
    end_cell: str = None
) -> Dict[str, Any]:
    """统一的单元格合并操作接口"""
    action = action.lower()
    
    if action == "merge":
        if not start_cell or not end_cell:
            raise ValidationError("merge 操作需要 start_cell 和 end_cell 参数")
        return merge_range(filepath, sheet_name, start_cell, end_cell)
    
    elif action == "unmerge":
        if not start_cell or not end_cell:
            raise ValidationError("unmerge 操作需要 start_cell 和 end_cell 参数")
        return unmerge_range(filepath, sheet_name, start_cell, end_cell)
    
    elif action == "list":
        merged = get_merged_ranges(filepath, sheet_name)
        return {
            "message": f"共 {len(merged)} 个合并区域",
            "merged_ranges": merged
        }
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: merge, unmerge, list"
        )


def row_column_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_index: int = None,
    count: int = 1
) -> Dict[str, Any]:
    """统一的行列操作接口"""
    action = action.lower()
    
    if start_index is None:
        raise ValidationError("需要 start_index 参数")
    
    if action == "insert_rows":
        return insert_row(filepath, sheet_name, start_index, count)
    
    elif action == "insert_cols":
        return insert_cols(filepath, sheet_name, start_index, count)
    
    elif action == "delete_rows":
        return delete_rows(filepath, sheet_name, start_index, count)
    
    elif action == "delete_cols":
        return delete_cols(filepath, sheet_name, start_index, count)
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: insert_rows, insert_cols, delete_rows, delete_cols"
        )


def range_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_cell: str = None,
    end_cell: str = None,
    target_cell: str = None,
    target_sheet: str = None,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """统一的范围操作接口"""
    action = action.lower()
    
    if action == "copy":
        if not start_cell or not end_cell or not target_cell:
            raise ValidationError("copy 操作需要 start_cell, end_cell, target_cell 参数")
        return copy_range_operation(
            filepath, sheet_name, start_cell, end_cell, 
            target_cell, target_sheet or sheet_name
        )
    
    elif action == "delete":
        if not start_cell or not end_cell:
            raise ValidationError("delete 操作需要 start_cell, end_cell 参数")
        return delete_range_operation(
            filepath, sheet_name, start_cell, end_cell, shift_direction
        )
    
    elif action == "validate":
        if not start_cell:
            raise ValidationError("validate 操作需要 start_cell 参数")
        from .validation import validate_range_in_sheet_operation
        range_str = f"{start_cell}:{end_cell}" if end_cell else start_cell
        return validate_range_in_sheet_operation(filepath, sheet_name, range_str)
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: copy, delete, validate"
        )
