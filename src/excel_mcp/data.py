# -*- coding: utf-8 -*-
"""数据读写模块 - 使用 xlwings 实现

本模块提供 Excel 数据的读取和写入功能。
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
import logging

import xlwings as xw

from .exceptions import DataError
from .xw_helper import (
    get_workbook,
    get_workbook_ex,
    get_sheet,
    save_workbook,
    parse_range,
    column_string_from_index,
    parse_cell_reference,
    close_workbook,
    cleanup_excel_app,
    ExcelNotFoundError,
    WorkbookContext,
)

logger = logging.getLogger(__name__)


def read_excel_range(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> List[List[Any]]:
    """读取 Excel 范围数据
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格，默认 "A1"
        end_cell: 结束单元格，可选
        preview_only: 是否仅预览模式
        
    Returns:
        二维列表，包含单元格值
        
    Raises:
        DataError: 当读取失败时
    """
    try:
        with WorkbookContext(str(filepath)) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            # 解析范围
            if ':' in start_cell:
                start_cell, end_cell = start_cell.split(':')

            # 获取范围
            if end_cell:
                rng = parse_range(sheet, start_cell, end_cell)
            else:
                # 如果没有指定结束单元格，使用已使用范围
                used_range = sheet.used_range
                if used_range is None or used_range.value is None:
                    return []
                rng = used_range
            
            # 读取数据
            value = rng.value
            
            # 处理单个单元格的情况
            if not isinstance(value, (list, tuple)):
                return [[value]] if value is not None else []
            
            # 处理单行的情况
            if value and not isinstance(value[0], (list, tuple)):
                return [list(value)]
            
            # 转换为列表
            data = []
            for row in value:
                if row is not None:
                    row_data = list(row) if isinstance(row, tuple) else row
                    if any(v is not None for v in row_data):
                        data.append(row_data)
            
            return data
        
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise DataError(f"Excel 未安装或无法启动: {e}")
    except DataError:
        raise
    except Exception as e:
        logger.error(f"读取 Excel 范围失败: {e}")
        raise DataError(str(e))


def write_data(
    filepath: str,
    sheet_name: Optional[str],
    data: Optional[List[List]],
    start_cell: str = "A1",
) -> Dict[str, str]:
    """写入数据到 Excel 工作表
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称，如果为 None 则使用活动工作表
        data: 要写入的二维数据列表
        start_cell: 起始单元格，默认 "A1"
        
    Returns:
        包含操作结果的字典
        
    Raises:
        DataError: 当写入失败时
    """
    try:
        if not data:
            raise DataError("未提供要写入的数据")
        
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            
            # 确定工作表
            if not sheet_name:
                sheet = wb.sheets.active
                if sheet is None:
                    raise DataError("工作簿中没有活动工作表")
                sheet_name = sheet.name
            else:
                # 检查工作表是否存在，不存在则创建
                sheet_names = [s.name for s in wb.sheets]
                if sheet_name not in sheet_names:
                    wb.sheets.add(sheet_name)
                sheet = ctx.get_sheet(sheet_name)
            
            # 写入数据
            rng = sheet.range(start_cell)
            rng.value = data
            
            # 保存
            ctx.save()
            
            return {"message": f"数据已写入 {sheet_name}", "active_sheet": sheet_name}
        
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise DataError(f"Excel 未安装或无法启动: {e}")
    except DataError:
        raise
    except Exception as e:
        logger.error(f"写入数据失败: {e}")
        raise DataError(str(e))


def _infer_column_type(values: List[Any]) -> str:
    """推断列的数据类型
    
    Args:
        values: 列值列表（不含表头）
        
    Returns:
        推断的类型字符串
    """
    # 过滤掉 None 值
    non_null = [v for v in values if v is not None]
    if not non_null:
        return "empty"
    
    # 统计类型
    type_counts = {}
    for v in non_null:
        if isinstance(v, bool):
            t = "boolean"
        elif isinstance(v, (int, float)):
            t = "number"
        elif isinstance(v, str):
            t = "string"
        else:
            t = type(v).__name__
        type_counts[t] = type_counts.get(t, 0) + 1
    
    # 返回最常见的类型
    if type_counts:
        return max(type_counts, key=type_counts.get)
    return "mixed"


def _analyze_column(col_values: List[Any], col_name: Any, col_index: int) -> Dict[str, Any]:
    """分析单列的结构信息
    
    Args:
        col_values: 列值列表（不含表头）
        col_name: 列名
        col_index: 列索引
        
    Returns:
        列分析结果字典
    """
    non_null = [v for v in col_values if v is not None]
    total = len(col_values)
    non_null_count = len(non_null)
    
    # 计算唯一值
    try:
        unique_values = set(str(v) for v in non_null)
        unique_count = len(unique_values)
    except Exception:
        unique_count = non_null_count
        unique_values = set()
    
    # 判断是否可能是索引列（唯一值比例高）
    is_potential_index = (
        unique_count == non_null_count and 
        non_null_count > 0 and 
        non_null_count == total
    )
    
    result = {
        "column_index": col_index,
        "column_letter": column_string_from_index(col_index),
        "column_name": col_name,
        "data_type": _infer_column_type(col_values),
        "total_rows": total,
        "non_null_count": non_null_count,
        "null_count": total - non_null_count,
        "unique_count": unique_count,
        "is_potential_index": is_potential_index,
    }
    
    # 如果唯一值较少，列出所有可能值（用于枚举类型识别）
    if unique_count <= 10 and unique_values:
        result["unique_values"] = sorted(list(unique_values))[:10]
    
    return result


def _compress_large_data(
    rng,
    sheet_name: str,
    sample_head: int = 5,
    sample_tail: int = 3
) -> Dict[str, Any]:
    """压缩大数据集，返回结构信息和样例数据
    
    Args:
        rng: xlwings Range 对象
        sheet_name: 工作表名称
        sample_head: 头部样例行数
        sample_tail: 尾部样例行数
        
    Returns:
        压缩后的数据结构字典
    """
    start_row = rng.row
    start_col = rng.column
    
    # 读取所有值
    value = rng.value
    
    # 标准化为二维列表
    if not isinstance(value, (list, tuple)):
        rows = [[value]]
    elif not isinstance(value[0], (list, tuple)):
        rows = [list(value)]
    else:
        rows = [list(row) if row else [] for row in value]
    
    total_rows = len(rows)
    total_cols = max(len(row) for row in rows) if rows else 0
    
    # 假设第一行是表头
    headers = rows[0] if rows else []
    data_rows = rows[1:] if len(rows) > 1 else []
    
    # 分析每列结构
    columns_info = []
    for j in range(total_cols):
        col_name = headers[j] if j < len(headers) else None
        col_values = []
        for row in data_rows:
            val = row[j] if j < len(row) else None
            col_values.append(val)
        
        col_info = _analyze_column(col_values, col_name, start_col + j)
        columns_info.append(col_info)
    
    # 识别潜在索引列
    potential_indexes = [
        col["column_name"] for col in columns_info 
        if col.get("is_potential_index")
    ]
    
    # 提取样例数据
    sample_rows = []
    
    # 添加表头
    sample_rows.append({
        "row_type": "header",
        "row_number": start_row,
        "values": headers
    })
    
    # 添加头部样例
    head_count = min(sample_head, len(data_rows))
    for i in range(head_count):
        sample_rows.append({
            "row_type": "data",
            "row_number": start_row + 1 + i,
            "values": data_rows[i]
        })
    
    # 如果数据行数超过头部+尾部，添加省略标记和尾部样例
    if len(data_rows) > sample_head + sample_tail:
        omitted_count = len(data_rows) - sample_head - sample_tail
        sample_rows.append({
            "row_type": "omitted",
            "omitted_rows": omitted_count,
            "message": f"... 省略 {omitted_count} 行数据 ..."
        })
        
        # 添加尾部样例
        for i in range(sample_tail):
            idx = len(data_rows) - sample_tail + i
            sample_rows.append({
                "row_type": "data",
                "row_number": start_row + 1 + idx,
                "values": data_rows[idx]
            })
    elif len(data_rows) > head_count:
        # 数据不多，但还有剩余的行
        for i in range(head_count, len(data_rows)):
            sample_rows.append({
                "row_type": "data",
                "row_number": start_row + 1 + i,
                "values": data_rows[i]
            })
    
    return {
        "compressed": True,
        "compression_reason": f"数据量较大 ({total_rows} 行 x {total_cols} 列)，已压缩为结构摘要",
        "range": rng.address,
        "sheet_name": sheet_name,
        "summary": {
            "total_rows": total_rows,
            "data_rows": total_rows - 1,  # 减去表头
            "total_columns": total_cols,
            "start_cell": f"{column_string_from_index(start_col)}{start_row}",
            "end_cell": f"{column_string_from_index(start_col + total_cols - 1)}{start_row + total_rows - 1}",
        },
        "columns": columns_info,
        "potential_index_columns": potential_indexes,
        "sample_data": sample_rows
    }


def read_excel_range_with_metadata(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    include_validation: bool = True,
    compression_threshold: int = 100,
    sample_head: int = 5,
    sample_tail: int = 3
) -> Dict[str, Any]:
    """读取带元数据的 Excel 范围数据
    
    当数据量超过阈值时，自动压缩为结构摘要+样例数据，避免消耗过多 tokens。
    
    智能关闭策略：
    - 如果文件原本已在 Excel 中打开，读取后保持打开状态
    - 如果文件是本次静默打开的，读取后自动关闭释放锁定
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格
        end_cell: 结束单元格（可选）
        include_validation: 是否包含验证元数据
        compression_threshold: 压缩阈值（行数），超过此值自动压缩，默认 100
        sample_head: 压缩时头部样例行数，默认 5
        sample_tail: 压缩时尾部样例行数，默认 3
        
    Returns:
        包含结构化单元格数据和元数据的字典。
        当数据量大时返回压缩的结构摘要。
        
    Raises:
        DataError: 当读取失败时
    """
    wb = None
    was_already_open = True
    app = None
    
    try:
        # 使用扩展版获取工作簿，同时获取打开状态
        wb, was_already_open, app = get_workbook_ex(str(filepath))
        sheet = get_sheet(wb, sheet_name)
        
        # 解析范围
        if ':' in start_cell:
            start_cell, end_cell = start_cell.split(':')
        
        # 获取范围
        if end_cell:
            rng = parse_range(sheet, start_cell, end_cell)
        else:
            # 使用已使用范围
            used_range = sheet.used_range
            if used_range is None or used_range.value is None:
                return {
                    "range": f"{start_cell}:",
                    "sheet_name": sheet_name,
                    "cells": []
                }
            rng = used_range
        
        # 检查数据量，决定是否压缩
        row_count = rng.rows.count
        col_count = rng.columns.count
        
        if row_count > compression_threshold:
            # 数据量大，返回压缩的结构摘要
            result = _compress_large_data(
                rng, 
                sheet_name, 
                sample_head=sample_head,
                sample_tail=sample_tail
            )
            return result
        
        # 数据量小，返回完整数据
        # 构建范围字符串
        range_str = rng.address
        
        # 构建单元格数据
        cells = []
        
        # 获取范围的行列信息
        start_row = rng.row
        start_col = rng.column
        
        # 读取值
        value = rng.value
        
        # 处理单个单元格
        if not isinstance(value, (list, tuple)):
            cell_address = f"{column_string_from_index(start_col)}{start_row}"
            cell_data = {
                "address": cell_address,
                "value": value,
                "row": start_row,
                "column": start_col
            }
            if include_validation:
                cell_data["validation"] = {"has_validation": False}
            cells.append(cell_data)
        else:
            # 处理多个单元格
            rows = value if isinstance(value[0], (list, tuple)) else [value]
            for i, row in enumerate(rows):
                if row is None:
                    continue
                row_values = row if isinstance(row, (list, tuple)) else [row]
                for j, val in enumerate(row_values):
                    current_row = start_row + i
                    current_col = start_col + j
                    cell_address = f"{column_string_from_index(current_col)}{current_row}"
                    
                    cell_data = {
                        "address": cell_address,
                        "value": val,
                        "row": current_row,
                        "column": current_col
                    }
                    
                    if include_validation:
                        # xlwings 获取验证信息较复杂，暂时标记为无验证
                        cell_data["validation"] = {"has_validation": False}
                    
                    cells.append(cell_data)
        
        return {
            "compressed": False,
            "range": range_str,
            "sheet_name": sheet_name,
            "cells": cells
        }
        
    except ExcelNotFoundError as e:
        logger.error(f"Excel 未安装: {e}")
        raise DataError(f"Excel 未安装或无法启动: {e}")
    except DataError:
        raise
    except Exception as e:
        logger.error(f"读取 Excel 范围元数据失败: {e}")
        raise DataError(str(e))
    finally:
        # 智能关闭：仅当文件是本次新打开的才关闭
        if not was_already_open and wb is not None:
            try:
                wb.close()
                if app is not None:
                    app.quit()
            except Exception:
                pass
