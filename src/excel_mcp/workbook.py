# -*- coding: utf-8 -*-
"""工作簿操作模块 - 使用 xlwings 实现

本模块提供工作簿的创建、获取元数据和工作表管理功能。
"""

import logging
from pathlib import Path
from typing import Any, Optional

import xlwings as xw

from .exceptions import WorkbookError
from .xw_helper import (
    get_app,
    get_workbook,
    get_sheet,
    save_workbook,
    column_string_from_index,
    ExcelNotFoundError,
    WorkbookContext,
)

logger = logging.getLogger(__name__)


def create_workbook(filepath: str, sheet_name: str = "Sheet1") -> dict[str, Any]:
    """创建新的 Excel 工作簿
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 初始工作表名称，默认为 "Sheet1"
        
    Returns:
        包含创建结果的字典
        
    Raises:
        WorkbookError: 当创建失败时
    """
    try:
        path = Path(filepath).resolve()
        
        # 检查文件是否已存在
        if path.exists():
            raise WorkbookError(f"文件已存在: {filepath}")
        
        # 确保父目录存在
        path.parent.mkdir(parents=True, exist_ok=True)
        
        # 检查所有 Excel 实例中是否有未保存的空白工作簿可以复用
        wb = None
        target_app = None
        
        for app in xw.apps:
            for book in app.books:
                try:
                    book_path = book.fullname
                    # 如果 fullname 不包含路径分隔符，说明是未保存的新工作簿
                    if '\\' not in book_path and '/' not in book_path:
                        # 检查是否只有一个工作表且为空
                        if len(book.sheets) == 1:
                            sheet = book.sheets[0]
                            used_range = sheet.used_range
                            # 检查工作表是否为空
                            if used_range is None or used_range.value is None:
                                wb = book
                                target_app = app
                                break
                except Exception:
                    continue
            if wb is not None:
                break
        
        # 如果没有可复用的空白工作簿，获取或创建 Excel 应用并添加新工作簿
        if wb is None:
            target_app = get_app(visible=True)
            wb = target_app.books.add()
        else:
            # 确保复用的工作簿所在的 Excel 实例可见
            target_app.visible = True

        # 重命名默认工作表
        if wb.sheets:
            wb.sheets[0].name = sheet_name
        
        # 保存工作簿
        wb.save(str(path))
        
        return {
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
            "filepath": str(path)
        }
    except WorkbookError:
        raise
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise WorkbookError(f"Excel 未安装或无法启动: {e}")
    except Exception as e:
        logger.error(f"创建工作簿失败: {e}")
        raise WorkbookError(f"创建工作簿失败: {e!s}")


def get_or_create_workbook(filepath: str) -> xw.Book:
    """获取现有工作簿或创建新工作簿
    
    Args:
        filepath: 工作簿文件路径
        
    Returns:
        xlwings Book 对象
    """
    return get_workbook(filepath, create_if_missing=True)


def create_sheet(filepath: str, sheet_name: str) -> dict:
    """在工作簿中创建新工作表
    
    Args:
        filepath: 工作簿文件路径
        sheet_name: 新工作表名称
        
    Returns:
        包含创建结果的字典
        
    Raises:
        WorkbookError: 当创建失败时
    """
    try:
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            # 检查工作表是否已存在
            existing_sheets = [s.name for s in wb.sheets]
            if sheet_name in existing_sheets:
                raise WorkbookError(f"工作表 '{sheet_name}' 已存在")
            
            # 创建新工作表
            wb.sheets.add(sheet_name)
            ctx.save()
            
            return {"message": f"工作表 '{sheet_name}' 创建成功"}
    except WorkbookError:
        raise
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise WorkbookError(f"Excel 未安装或无法启动: {e}")
    except Exception as e:
        logger.error(f"创建工作表失败: {e}")
        raise WorkbookError(str(e))


def get_workbook_info(filepath: str, include_ranges: bool = False) -> dict[str, Any]:
    """获取工作簿元数据
    
    Args:
        filepath: 工作簿文件路径
        include_ranges: 是否包含每个工作表的使用范围
        
    Returns:
        包含工作簿元数据的字典
        
    Raises:
        WorkbookError: 当获取失败时
    """
    try:
        path = Path(filepath).resolve()
        if not path.exists():
            raise WorkbookError(f"文件不存在: {filepath}")
        
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            info = {
                "filename": path.name,
                "sheets": [s.name for s in wb.sheets],
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime
            }
            
            if include_ranges:
                # 获取每个工作表的使用范围
                ranges = {}
                for sheet in wb.sheets:
                    # 使用 xlwings 获取已使用范围
                    used_range = sheet.used_range
                    if used_range is not None:
                        last_row = used_range.last_cell.row
                        last_col = used_range.last_cell.column
                        if last_row > 0 and last_col > 0:
                            col_letter = column_string_from_index(last_col)
                            ranges[sheet.name] = f"A1:{col_letter}{last_row}"
                info["used_ranges"] = ranges
            
            return info
        
    except WorkbookError:
        raise
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise WorkbookError(f"Excel 未安装或无法启动: {e}")
    except Exception as e:
        logger.error(f"获取工作簿信息失败: {e}")
        raise WorkbookError(str(e))
