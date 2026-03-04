# -*- coding: utf-8 -*-
"""格式化模块 - 使用 xlwings 实现

本模块提供 Excel 单元格格式化功能。
"""

import logging
from typing import Any, Dict, Optional

from .xw_helper import get_workbook, get_sheet, save_workbook, hex_to_rgb, WorkbookContext
from .cell_utils import parse_cell_range, validate_cell_reference
from .exceptions import ValidationError, FormattingError

logger = logging.getLogger(__name__)


def format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    border_style: Optional[str] = None,
    border_color: Optional[str] = None,
    number_format: Optional[str] = None,
    alignment: Optional[str] = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Optional[Dict[str, Any]] = None,
    conditional_format: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """应用格式到单元格范围
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格引用
        end_cell: 结束单元格引用（可选）
        bold: 是否加粗
        italic: 是否斜体
        underline: 是否下划线
        font_size: 字体大小
        font_color: 字体颜色（十六进制）
        bg_color: 背景颜色（十六进制）
        border_style: 边框样式
        border_color: 边框颜色（十六进制）
        number_format: 数字格式
        alignment: 对齐方式
        wrap_text: 是否自动换行
        merge_cells: 是否合并单元格
        protection: 保护设置
        conditional_format: 条件格式设置
        
    Returns:
        包含操作状态的字典
    """
    try:
        # 验证单元格引用
        if not validate_cell_reference(start_cell):
            raise ValidationError(f"Invalid start cell reference: {start_cell}")
            
        if end_cell and not validate_cell_reference(end_cell):
            raise ValidationError(f"Invalid end cell reference: {end_cell}")
        
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 获取范围
            if end_cell:
                cell_range = sheet.range(f"{start_cell}:{end_cell}")
            else:
                cell_range = sheet.range(start_cell)
            
            # 应用字体格式
            if bold:
                cell_range.font.bold = True
            if italic:
                cell_range.font.italic = True
            if underline:
                cell_range.api.Font.Underline = 2  # xlUnderlineStyleSingle
            if font_size is not None:
                cell_range.font.size = font_size
            if font_color is not None:
                try:
                    # 移除可能的 # 或 FF 前缀
                    color = font_color.lstrip('#').lstrip('F').lstrip('f')
                    if len(color) == 6:
                        rgb = hex_to_rgb(color)
                        cell_range.font.color = rgb
                except Exception as e:
                    raise FormattingError(f"Invalid font color: {str(e)}")
            
            # 应用背景色
            if bg_color is not None:
                try:
                    color = bg_color.lstrip('#').lstrip('F').lstrip('f')
                    if len(color) == 6:
                        rgb = hex_to_rgb(color)
                        cell_range.color = rgb
                except Exception as e:
                    raise FormattingError(f"Invalid background color: {str(e)}")
            
            # 应用边框
            if border_style is not None:
                try:
                    # xlwings 边框样式映射
                    style_map = {
                        'thin': 1,
                        'medium': 2,
                        'thick': 5,
                        'double': 6,
                        'hair': 7,
                        'dotted': 4,
                        'dashed': 3,
                    }
                    xl_style = style_map.get(border_style.lower(), 1)
                    
                    # 设置边框颜色
                    if border_color:
                        color = border_color.lstrip('#').lstrip('F').lstrip('f')
                        if len(color) == 6:
                            rgb = hex_to_rgb(color)
                            color_value = rgb[0] + rgb[1] * 256 + rgb[2] * 65536
                        else:
                            color_value = 0  # 黑色
                    else:
                        color_value = 0  # 黑色
                    
                    # 应用四边边框
                    for edge in range(7, 13):  # xlEdgeLeft to xlInsideHorizontal
                        try:
                            cell_range.api.Borders(edge).LineStyle = xl_style
                            cell_range.api.Borders(edge).Color = color_value
                        except Exception:
                            pass  # 某些边框可能不适用于单个单元格
                            
                except Exception as e:
                    raise FormattingError(f"Invalid border settings: {str(e)}")
            
            # 应用数字格式
            if number_format is not None:
                cell_range.number_format = number_format
            
            # 应用对齐
            if alignment is not None:
                align_map = {
                    'left': -4131,    # xlLeft
                    'center': -4108,  # xlCenter
                    'right': -4152,   # xlRight
                    'justify': -4130, # xlJustify
                }
                xl_align = align_map.get(alignment.lower(), -4131)
                cell_range.api.HorizontalAlignment = xl_align
            
            # 应用自动换行
            if wrap_text:
                cell_range.api.WrapText = True
            
            # 合并单元格
            if merge_cells and end_cell:
                cell_range.merge()
            
            # 保存工作簿
            ctx.save()
            
            range_str = f"{start_cell}:{end_cell}" if end_cell else start_cell
            return {
                "message": f"Applied formatting to range {range_str}",
                "range": range_str
            }
        
    except (ValidationError, FormattingError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formatting: {e}")
        raise FormattingError(str(e))
