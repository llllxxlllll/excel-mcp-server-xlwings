# -*- coding: utf-8 -*-
"""表格模块 - 使用 xlwings 实现

本模块提供 Excel 表格创建功能。
"""

import uuid
import logging
from typing import Optional

from .xw_helper import get_workbook, get_sheet, save_workbook, WorkbookContext
from .exceptions import DataError

logger = logging.getLogger(__name__)


def create_excel_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> dict:
    """创建 Excel 原生表格
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        data_range: 表格数据范围（如 "A1:D5"）
        table_name: 表格唯一名称，如果未提供则自动生成
        table_style: 表格视觉样式
        
    Returns:
        包含成功消息和表格详情的字典
    """
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 如果未提供表格名称，生成唯一名称
            if not table_name:
                table_name = f"Table_{uuid.uuid4().hex[:8]}"
            
            # 检查表格名称是否已存在
            try:
                existing_tables = [t.name for t in sheet.api.ListObjects]
                if table_name in existing_tables:
                    raise DataError(f"Table name '{table_name}' already exists.")
            except Exception:
                pass  # 如果无法获取现有表格列表，继续创建
            
            # 获取数据范围
            data_rng = sheet.range(data_range)
            
            # 创建表格
            # 使用 Excel API 创建 ListObject（表格）
            table = sheet.api.ListObjects.Add(
                SourceType=1,  # xlSrcRange
                Source=data_rng.api,
                XlListObjectHasHeaders=1  # xlYes
            )
            
            # 设置表格名称
            table.Name = table_name
            
            # 应用样式
            try:
                table.TableStyle = table_style
            except Exception:
                # 如果样式不存在，使用默认样式
                pass
            
            # 保存工作簿
            ctx.save()
            
            return {
                "message": f"Successfully created table '{table_name}' in sheet '{sheet_name}'.",
                "table_name": table_name,
                "range": data_range
            }
        
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to create table: {e}")
        raise DataError(str(e))
