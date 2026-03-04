# -*- coding: utf-8 -*-
"""xlwings 辅助模块 - 封装 xlwings 操作

本模块提供统一的 xlwings 操作接口，包括：
- Excel 应用程序实例管理
- 工作簿获取和创建
- 工作表操作
- 范围解析
"""

import xlwings as xw
from pathlib import Path
from typing import Optional, Tuple, Union

from .exceptions import WorkbookError, SheetError


class ExcelNotFoundError(Exception):
    """Excel 未安装或无法启动异常"""
    pass


def column_index_from_string(col_str: str) -> int:
    """将列字母转换为列索引（1-based）
    
    Args:
        col_str: 列字母，如 'A', 'B', 'AA'
        
    Returns:
        列索引，从1开始
    """
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def column_string_from_index(col_idx: int) -> str:
    """将列索引转换为列字母（1-based）
    
    Args:
        col_idx: 列索引，从1开始
        
    Returns:
        列字母，如 'A', 'B', 'AA'
    """
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result


# 全局配置：是否显示 Excel 窗口（测试时可设为 False）
EXCEL_VISIBLE = True


def get_app(visible: bool = None) -> xw.App:
    """获取或创建 Excel 应用程序实例
    
    如果已有 Excel 实例运行，则复用现有实例；
    否则创建新的 Excel 应用程序实例。
    
    Args:
        visible: 是否显示 Excel 窗口，默认使用全局配置 EXCEL_VISIBLE
        
    Returns:
        xlwings App 对象
        
    Raises:
        ExcelNotFoundError: 当 Excel 未安装或无法启动时
    """
    if visible is None:
        visible = EXCEL_VISIBLE
        
    try:
        # 尝试获取现有的 Excel 应用程序实例
        apps = xw.apps
        if apps:
            app = apps.active
            if app is not None:
                app.visible = visible
                return app
        
        # 创建新的 Excel 应用程序实例（不自动创建空白工作簿）
        app = xw.App(visible=visible, add_book=False)
        return app
    except Exception as e:
        error_msg = str(e).lower()
        if "excel" in error_msg or "com" in error_msg or "dispatch" in error_msg:
            raise ExcelNotFoundError(
                "无法启动 Excel 应用程序。请确保已安装 Microsoft Excel。"
                f"\n详细错误: {e}"
            )
        raise


def get_workbook(
    filepath: str,
    create_if_missing: bool = False,
    visible: bool = None
) -> xw.Book:
    """获取工作簿，支持连接已打开的文件
    
    如果文件已在 Excel 中打开，则连接到现有工作簿实例；
    否则打开文件或创建新工作簿。
    
    Args:
        filepath: 工作簿文件路径
        create_if_missing: 如果文件不存在是否创建新工作簿
        visible: 是否显示 Excel 窗口
        
    Returns:
        xlwings Book 对象
        
    Raises:
        WorkbookError: 当工作簿操作失败时
        ExcelNotFoundError: 当 Excel 未安装时
    """
    wb, _, _ = get_workbook_ex(filepath, create_if_missing, visible)
    return wb


def get_workbook_ex(
    filepath: str,
    create_if_missing: bool = False,
    visible: bool = None
) -> Tuple[xw.Book, bool, Optional[xw.App]]:
    """获取工作簿（扩展版），返回额外的状态信息
    
    三种情形：
    - 情形1: 文件已在 Excel 中打开 -> 直接绑定，不关闭
    - 情形2: 文件存在但未打开 -> 静默打开，操作完可关闭
    - 情形3: 文件不存在 -> 创建新文件
    
    Args:
        filepath: 工作簿文件路径
        create_if_missing: 如果文件不存在是否创建新工作簿
        visible: 是否显示 Excel 窗口（仅对新打开的文件生效）
        
    Returns:
        (wb, was_already_open, app) 元组:
        - wb: xlwings Book 对象
        - was_already_open: True 表示文件原本已打开，False 表示本次新打开
        - app: 新创建的 App 实例（仅当 was_already_open=False 时有值）
        
    Raises:
        WorkbookError: 当工作簿操作失败时
        ExcelNotFoundError: 当 Excel 未安装时
    """
    path = Path(filepath).resolve()
    basename = path.name
    
    # 使用全局配置作为默认值
    if visible is None:
        visible = EXCEL_VISIBLE
    
    try:
        # 情形1: 检查文件是否已在任何 Excel 实例中打开
        for app in xw.apps:
            for book in app.books:
                try:
                    # 优先用文件名匹配（更快），再用完整路径确认
                    if book.name == basename:
                        book_path = Path(book.fullname).resolve()
                        if book_path == path:
                            # 文件已打开，返回现有工作簿，标记为已打开
                            return book, True, None
                except Exception:
                    # 忽略无法获取路径的工作簿（如未保存的新工作簿）
                    continue
        
        # 情形2/3: 文件未打开
        if path.exists():
            # 情形2: 文件存在，静默打开（visible=False 避免弹窗）
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.open(str(path))
                return wb, False, app
            except Exception as e:
                # 打开失败，清理 app
                try:
                    app.quit()
                except Exception:
                    pass
                raise WorkbookError(f"无法打开工作簿 '{filepath}': {e}")
        elif create_if_missing:
            # 情形3: 文件不存在，创建新工作簿
            app = xw.App(visible=visible, add_book=False)
            try:
                book = app.books.add()
                book.save(str(path))
                return book, False, app
            except Exception as e:
                try:
                    app.quit()
                except Exception:
                    pass
                raise WorkbookError(f"无法创建工作簿 '{filepath}': {e}")
        else:
            raise WorkbookError(f"工作簿文件不存在: {filepath}")
            
    except ExcelNotFoundError:
        raise
    except WorkbookError:
        raise
    except Exception as e:
        error_msg = str(e).lower()
        if "excel" in error_msg or "com" in error_msg:
            raise ExcelNotFoundError(
                f"无法打开工作簿，Excel 可能未正确安装: {e}"
            )
        raise WorkbookError(f"无法打开工作簿 '{filepath}': {e}")


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet:
    """获取工作表
    
    Args:
        wb: xlwings Book 对象
        sheet_name: 工作表名称
        
    Returns:
        xlwings Sheet 对象
        
    Raises:
        SheetError: 当工作表不存在时
    """
    try:
        return wb.sheets[sheet_name]
    except KeyError:
        available_sheets = [s.name for s in wb.sheets]
        raise SheetError(
            f"工作表 '{sheet_name}' 不存在。"
            f"可用的工作表: {available_sheets}"
        )
    except Exception as e:
        raise SheetError(f"无法获取工作表 '{sheet_name}': {e}")


def save_workbook(wb: xw.Book, filepath: Optional[str] = None) -> None:
    """保存工作簿
    
    Args:
        wb: xlwings Book 对象
        filepath: 保存路径，如果为 None 则保存到原路径
        
    Raises:
        WorkbookError: 当保存失败时
    """
    try:
        if filepath:
            wb.save(filepath)
        else:
            wb.save()
    except Exception as e:
        raise WorkbookError(f"无法保存工作簿: {e}")


def parse_range(
    sheet: xw.Sheet,
    start_cell: str,
    end_cell: Optional[str] = None
) -> xw.Range:
    """解析范围字符串为 xlwings Range 对象
    
    Args:
        sheet: xlwings Sheet 对象
        start_cell: 起始单元格，如 'A1'
        end_cell: 结束单元格，如 'B10'，可选
        
    Returns:
        xlwings Range 对象
        
    Raises:
        ValueError: 当范围字符串无效时
    """
    try:
        if end_cell:
            return sheet.range(f"{start_cell}:{end_cell}")
        else:
            return sheet.range(start_cell)
    except Exception as e:
        range_str = f"{start_cell}:{end_cell}" if end_cell else start_cell
        raise ValueError(f"无效的范围 '{range_str}': {e}")


def parse_cell_reference(cell_ref: str) -> Tuple[int, int]:
    """解析单元格引用为行列索引（1-based）
    
    Args:
        cell_ref: 单元格引用，如 'A1', 'BC123'
        
    Returns:
        (row, col) 元组，均为 1-based 索引
        
    Raises:
        ValueError: 当单元格引用无效时
    """
    import re
    match = re.match(r"([A-Za-z]+)([0-9]+)", cell_ref)
    if not match:
        raise ValueError(f"无效的单元格引用: {cell_ref}")
    
    col_str, row_str = match.groups()
    row = int(row_str)
    col = column_index_from_string(col_str)
    
    return row, col


def cell_reference_from_indices(row: int, col: int) -> str:
    """从行列索引生成单元格引用（1-based）
    
    Args:
        row: 行索引，从1开始
        col: 列索引，从1开始
        
    Returns:
        单元格引用字符串，如 'A1'
    """
    col_str = column_string_from_index(col)
    return f"{col_str}{row}"


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """将十六进制颜色转换为 RGB 元组
    
    Args:
        hex_color: 十六进制颜色，如 '#FF0000' 或 'FF0000'
        
    Returns:
        (R, G, B) 元组
    """
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    raise ValueError(f"无效的十六进制颜色: {hex_color}")


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """将 RGB 值转换为十六进制颜色
    
    Args:
        r: 红色值 (0-255)
        g: 绿色值 (0-255)
        b: 蓝色值 (0-255)
        
    Returns:
        十六进制颜色字符串，如 '#FF0000'
    """
    return f"#{r:02X}{g:02X}{b:02X}"


def close_workbook(wb: xw.Book, save: bool = False) -> None:
    """关闭工作簿
    
    Args:
        wb: xlwings Book 对象
        save: 是否在关闭前保存
        
    Raises:
        WorkbookError: 当关闭失败时
    """
    try:
        if save:
            wb.save()
        wb.close()
    except Exception as e:
        raise WorkbookError(f"无法关闭工作簿: {e}")


class WorkbookContext:
    """工作簿上下文管理器 - 自动处理文件锁定问题
    
    智能关闭策略：
    - 如果文件原本已在 Excel 中打开，操作后保持打开状态
    - 如果文件是本次静默打开的，操作后自动关闭释放锁定
    
    Usage:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            # 操作...
            ctx.save()  # 可选，保存更改
    """
    
    def __init__(
        self, 
        filepath: str, 
        create_if_missing: bool = False,
        visible: bool = None
    ):
        """初始化上下文管理器
        
        Args:
            filepath: 工作簿文件路径
            create_if_missing: 如果文件不存在是否创建
            visible: 是否显示 Excel 窗口
        """
        self.filepath = filepath
        self.create_if_missing = create_if_missing
        self.visible = visible
        self.wb: Optional[xw.Book] = None
        self.was_already_open: bool = True
        self.app: Optional[xw.App] = None
    
    def __enter__(self) -> "WorkbookContext":
        """进入上下文，获取工作簿"""
        self.wb, self.was_already_open, self.app = get_workbook_ex(
            self.filepath,
            create_if_missing=self.create_if_missing,
            visible=self.visible
        )
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """退出上下文，智能关闭工作簿"""
        # 仅当文件是本次新打开的才关闭
        if not self.was_already_open and self.wb is not None:
            try:
                self.wb.close()
                if self.app is not None:
                    # 检查 app 是否还有其他工作簿
                    if len(self.app.books) == 0:
                        self.app.quit()
            except Exception:
                pass
    
    def get_sheet(self, sheet_name: str) -> xw.Sheet:
        """获取工作表"""
        return get_sheet(self.wb, sheet_name)
    
    def save(self, filepath: Optional[str] = None) -> None:
        """保存工作簿"""
        save_workbook(self.wb, filepath)


def cleanup_excel_app() -> None:
    """清理没有打开工作簿的 Excel 应用程序实例
    
    遍历所有 Excel 实例，关闭没有打开任何工作簿的实例。
    """
    try:
        for app in list(xw.apps):
            try:
                if len(app.books) == 0:
                    app.quit()
            except Exception:
                pass
    except Exception:
        pass
