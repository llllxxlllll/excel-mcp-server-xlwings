import logging
import os
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Import exceptions
from excel_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError,
    VBAExecutionError,
    VBASecurityError,
    VBATimeoutError,
    VBABusyError,
)

# Import from excel_mcp package with consistent _impl suffixes
from excel_mcp.validation import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
from excel_mcp.chart import chart_operation as chart_operation_impl
from excel_mcp.workbook import get_workbook_info
from excel_mcp.data import write_data
from excel_mcp.pivot import create_pivot_table as create_pivot_table_impl
from excel_mcp.tables import create_excel_table as create_table_impl
from excel_mcp.sheet import (
    worksheet_operation as worksheet_operation_impl,
    merge_cell_operation as merge_cell_operation_impl,
    row_column_operation as row_column_operation_impl,
    range_operation as range_operation_impl,
)
from excel_mcp.calculations import formula_operation as formula_operation_impl
from excel_mcp.vba_executor import VBAExecutor

# Get project root directory path for log file path.
# When using the stdio transmission method,
# relative paths may cause log files to fail to create
# due to the client's running location and permission issues,
# resulting in the program not being able to run.
# Thus using os.path.join(ROOT_DIR, "excel-mcp.log") instead.

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(ROOT_DIR, "excel-mcp.log")

# Initialize EXCEL_FILES_PATH variable without assigning a value
EXCEL_FILES_PATH = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        # Referring to https://github.com/modelcontextprotocol/python-sdk/issues/409#issuecomment-2816831318
        # The stdio mode server MUST NOT write anything to its stdout that is not a valid MCP message.
        logging.FileHandler(LOG_FILE)
    ],
)
logger = logging.getLogger("excel-mcp")
# Initialize FastMCP server
mcp = FastMCP(
    "excel-mcp",
    host=os.environ.get("FASTMCP_HOST", "0.0.0.0"),
    port=int(os.environ.get("FASTMCP_PORT", "8017")),
    instructions="Excel MCP Server for manipulating Excel files"
)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename

    # Check if in SSE mode (EXCEL_FILES_PATH is not None)
    if EXCEL_FILES_PATH is None:
        # Must use absolute path
        raise ValueError(f"Invalid filename: {filename}, must be an absolute path when not in SSE mode")

    # In SSE mode, if it's a relative path, resolve it based on EXCEL_FILES_PATH
    return os.path.join(EXCEL_FILES_PATH, filename)

@mcp.tool()
def formula_operation(
    filepath: str,
    sheet_name: str,
    cell: str,
    action: str,
    formula: str = None,
) -> str:
    """Unified formula operation tool.
    
    统一的公式操作工具。
    
    Actions / 操作类型:
    - apply: Apply formula to cell (requires formula)
             应用公式到单元格（需要 formula）
    - validate: Validate formula syntax (requires formula)
                验证公式语法（需要 formula）
    - get: Get formula from cell
           获取单元格公式
    """
    try:
        full_path = get_excel_path(filepath)
        result = formula_operation_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            cell=cell,
            action=action,
            formula=formula
        )
        return result.get("message", str(result))
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in formula operation: {e}")
        return f"Error: {str(e)}"

@mcp.tool()
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
) -> str:
    """Apply formatting to a range of cells.
    
    对单元格区域应用格式设置。
    支持字体、颜色、边框、对齐、数字格式、条件格式等。
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.formatting import format_range as format_range_func
        
        # Convert None values to appropriate defaults for the underlying function
        format_range_func(
            filepath=full_path,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,  # This can be None
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,  # This can be None
            font_color=font_color,  # This can be None
            bg_color=bg_color,  # This can be None
            border_style=border_style,  # This can be None
            border_color=border_color,  # This can be None
            number_format=number_format,  # This can be None
            alignment=alignment,  # This can be None
            wrap_text=wrap_text,
            merge_cells=merge_cells,
            protection=protection,  # This can be None
            conditional_format=conditional_format  # This can be None
        )
        return "Range formatted successfully"
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error formatting range: {e}")
        raise

@mcp.tool()
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> str:
    """
    Read data from Excel worksheet with cell metadata including validation rules.
    
    从Excel工作表读取数据，包含单元格元数据和验证规则。
    
    **智能压缩**: 当数据超过100行时，自动返回压缩的结构摘要，包含：
    - 表格基本信息（行数、列数、范围）
    - 列结构分析（列名、数据类型、唯一值统计）
    - 潜在索引列识别
    - 样例数据（前5行 + 后3行）
    
    Args:
        filepath: Path to Excel file / Excel文件路径
        sheet_name: Name of worksheet / 工作表名称
        start_cell: Starting cell (default A1) / 起始单元格（默认A1）
        end_cell: Ending cell (optional, auto-expands if not provided) / 结束单元格（可选，不提供则自动扩展）
        preview_only: Whether to return preview only / 是否仅返回预览
    
    Returns:  
    JSON string containing structured cell data with validation metadata.
    Each cell includes: address, value, row, column, and validation info (if any).
    For large datasets (>100 rows), returns compressed structure summary with sample data.
    
    返回包含结构化单元格数据的JSON字符串。
    每个单元格包含：地址、值、行、列和验证信息（如有）。
    对于大数据集（>100行），返回压缩的结构摘要和样例数据。
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.data import read_excel_range_with_metadata
        result = read_excel_range_with_metadata(
            full_path, 
            sheet_name, 
            start_cell, 
            end_cell
        )
        
        # 检查是否为压缩结果
        if result.get("compressed"):
            import json
            return json.dumps(result, indent=2, default=str, ensure_ascii=False)
        
        # 非压缩结果，检查是否有数据
        if not result or not result.get("cells"):
            return "No data found in specified range"
            
        # Return as formatted JSON string
        import json
        return json.dumps(result, indent=2, default=str)
        
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise

@mcp.tool()
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1",
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.
    
    向Excel工作表写入数据。
    公式直接写入，不进行验证。

    PARAMETERS / 参数:  
    filepath: Path to Excel file / Excel文件路径
    sheet_name: Name of worksheet to write to / 要写入的工作表名称
    data: List of lists containing data to write, sublists are rows / 要写入的数据（二维列表，子列表为行）
    start_cell: Cell to start writing to, default is "A1" / 起始单元格，默认"A1"
    """
    try:
        full_path = get_excel_path(filepath)
        result = write_data(full_path, sheet_name, data, start_cell)
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise

@mcp.tool()
def create_workbook(filepath: str) -> str:
    """Create new Excel workbook.
    
    创建新的Excel工作簿。
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_workbook as create_workbook_impl
        create_workbook_impl(full_path)
        return f"Created workbook at {full_path}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise

@mcp.tool()
def chart_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    data_range: str = None,
    chart_type: str = None,
    target_cell: str = None,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    chart_index: int = None,
    chart_name: str = None,
    font_name: str = None,
    font_size: int = None,
    title_font_size: int = None
) -> str:
    """Unified chart operation tool.
    
    统一的图表操作工具。
    
    Actions / 操作类型:
    - create: Create a chart (requires data_range, chart_type, target_cell)
              创建图表（需要 data_range, chart_type, target_cell）
    - list: List all charts in worksheet
            列出工作表中所有图表
    - delete: Delete a chart (requires chart_index or chart_name)
              删除图表（需要 chart_index 或 chart_name）
    - style: Update chart style (requires chart_index or chart_name, optional font settings)
             更新图表样式（需要 chart_index 或 chart_name，可选字体设置）
    """
    try:
        full_path = get_excel_path(filepath)
        result = chart_operation_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            action=action,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis,
            chart_index=chart_index,
            chart_name=chart_name,
            font_name=font_name,
            font_size=font_size,
            title_font_size=title_font_size
        )
        if action.lower() == "list":
            import json
            return json.dumps(result, ensure_ascii=False, indent=2)
        return result.get("message", str(result))
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in chart operation: {e}")
        return f"Error: {str(e)}"


@mcp.tool()
def create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "mean"
) -> str:
    """Create pivot table in worksheet.
    
    在工作表中创建数据透视表。
    """
    try:
        full_path = get_excel_path(filepath)
        result = create_pivot_table_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            data_range=data_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func
        )
        return result["message"]
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating pivot table: {e}")
        raise

@mcp.tool()
def create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> str:
    """Creates a native Excel table from a specified range of data.
    
    从指定数据范围创建Excel原生表格。
    """
    try:
        full_path = get_excel_path(filepath)
        result = create_table_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            data_range=data_range,
            table_name=table_name,
            table_style=table_style
        )
        return result["message"]
    except DataError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating table: {e}")
        raise

@mcp.tool()
def worksheet_operation(
    filepath: str,
    action: str,
    sheet_name: str = None,
    new_name: str = None,
    source_sheet: str = None
) -> str:
    """Unified worksheet operation tool.
    
    统一的工作表操作工具。
    
    Actions / 操作类型:
    - create: Create worksheet (requires new_name)
              创建工作表（需要 new_name）
    - copy: Copy worksheet (requires source_sheet, new_name)
            复制工作表（需要 source_sheet, new_name）
    - delete: Delete worksheet (requires sheet_name)
              删除工作表（需要 sheet_name）
    - rename: Rename worksheet (requires sheet_name, new_name)
              重命名工作表（需要 sheet_name, new_name）
    - list: List all worksheets
            列出所有工作表
    """
    try:
        full_path = get_excel_path(filepath)
        result = worksheet_operation_impl(
            filepath=full_path,
            action=action,
            sheet_name=sheet_name,
            new_name=new_name,
            source_sheet=source_sheet
        )
        if action.lower() == "list":
            import json
            return json.dumps(result, ensure_ascii=False, indent=2)
        return result.get("message", str(result))
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in worksheet operation: {e}")
        return f"Error: {str(e)}"

@mcp.tool()
def get_workbook_metadata(
    filepath: str,
    include_ranges: bool = False
) -> str:
    """Get metadata about workbook including sheets, ranges, etc.
    
    获取工作簿元数据，包括工作表列表、数据范围等。
    """
    try:
        full_path = get_excel_path(filepath)
        result = get_workbook_info(full_path, include_ranges=include_ranges)
        return str(result)
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise

@mcp.tool()
def merge_cell_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_cell: str = None,
    end_cell: str = None
) -> str:
    """Unified cell merge operation tool.
    
    统一的单元格合并操作工具。
    
    Actions / 操作类型:
    - merge: Merge cells (requires start_cell, end_cell)
             合并单元格（需要 start_cell, end_cell）
    - unmerge: Unmerge cells (requires start_cell, end_cell)
               取消合并（需要 start_cell, end_cell）
    - list: List all merged cells
            列出所有合并单元格
    """
    try:
        full_path = get_excel_path(filepath)
        result = merge_cell_operation_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            action=action,
            start_cell=start_cell,
            end_cell=end_cell
        )
        if action.lower() == "list":
            import json
            return json.dumps(result, ensure_ascii=False, indent=2)
        return result.get("message", str(result))
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in merge cell operation: {e}")
        return f"Error: {str(e)}"


@mcp.tool()
def range_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_cell: str = None,
    end_cell: str = None,
    target_cell: str = None,
    target_sheet: str = None,
    shift_direction: str = "up"
) -> str:
    """Unified range operation tool.
    
    统一的范围操作工具。
    
    Actions / 操作类型:
    - copy: Copy range (requires start_cell, end_cell, target_cell)
            复制范围（需要 start_cell, end_cell, target_cell）
    - delete: Delete range (requires start_cell, end_cell)
              删除范围（需要 start_cell, end_cell）
    - validate: Validate range (requires start_cell)
                验证范围（需要 start_cell）
    """
    try:
        full_path = get_excel_path(filepath)
        result = range_operation_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            action=action,
            start_cell=start_cell,
            end_cell=end_cell,
            target_cell=target_cell,
            target_sheet=target_sheet,
            shift_direction=shift_direction
        )
        return result.get("message", str(result))
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in range operation: {e}")
        return f"Error: {str(e)}"

@mcp.tool()
def get_data_validation_info(
    filepath: str,
    sheet_name: str
) -> str:
    """
    Get all data validation rules in a worksheet.
    
    This tool helps identify which cell ranges have validation rules
    and what types of validation are applied.
    
    获取工作表中的所有数据验证规则。
    帮助识别哪些单元格区域有验证规则及其类型。
    
    Args:
        filepath: Path to Excel file / Excel文件路径
        sheet_name: Name of worksheet / 工作表名称
        
    Returns:
        JSON string containing all validation rules in the worksheet
        返回包含所有验证规则的JSON字符串
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.cell_validation import get_data_validation_info as get_validation_impl
        
        result = get_validation_impl(full_path, sheet_name)
        
        if result.get("status") == "error":
            return f"Error: {result.get('message', 'Unknown error')}"
        
        if not result.get("validations"):
            return "No data validation rules found in this worksheet"
            
        import json
        return json.dumps({
            "sheet_name": sheet_name,
            "validation_rules": result.get("validations", [])
        }, indent=2, default=str)
        
    except Exception as e:
        logger.error(f"Error getting validation info: {e}")
        raise

@mcp.tool()
def execute_excel_vba(
    filepath: str,
    vba_code: str,
    entry_sub_name: str = "Main"
) -> str:
    """Execute dynamic VBA code on an Excel file.
    
    在 Excel 文件上执行动态 VBA 代码。
    
    SECURITY WARNING / 安全警告:
    - VBA code is scanned for sensitive keywords before execution
    - A backup file is created before any modifications
    - VBA 代码在执行前会扫描敏感关键词
    - 执行前会创建原文件的备份
    
    Args:
        filepath: Path to Excel file (absolute path required in stdio mode)
                  Excel 文件路径（stdio 模式需要绝对路径）
        vba_code: Complete VBA code string, must contain a Sub matching entry_sub_name
                  完整的 VBA 代码字符串，必须包含与 entry_sub_name 匹配的 Sub
        entry_sub_name: Name of the entry Sub procedure, default "Main"
                        入口 Sub 过程名称，默认为 "Main"
    
    Returns:
        JSON string containing execution result with status, message, logs, and backup_path
        包含执行结果的 JSON 字符串，包括 status, message, logs, backup_path
    
    Example VBA code / VBA 代码示例:
        Sub Main()
            Cells(1, 1).Value = "Hello"
        End Sub

    NOTE / 注意：
    - When calling this tool via MCP, avoid using MsgBox, InputBox or other interactive dialogs,
      as they will block the tool call waiting for user interaction.
    - 通过 MCP 调用本工具时，应避免使用 MsgBox、InputBox 等需要人工交互的对话框，
      否则会因为等待用户点击/输入而导致调用卡顿或挂起。推荐将调试信息写入单元格或返回结果中。
    """
    import json
    
    try:
        full_path = get_excel_path(filepath)
        executor = VBAExecutor()
        result = executor.execute_vba(full_path, vba_code, entry_sub_name)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except VBASecurityError as e:
        return json.dumps({
            "status": "error",
            "message": f"安全检查失败: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)
    except VBATimeoutError as e:
        return json.dumps({
            "status": "error",
            "message": f"执行超时: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)
    except VBABusyError as e:
        return json.dumps({
            "status": "error",
            "message": f"Excel 正忙: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)
    except VBAExecutionError as e:
        return json.dumps({
            "status": "error",
            "message": f"VBA 执行错误: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)
    except WorkbookError as e:
        return json.dumps({
            "status": "error",
            "message": f"工作簿错误: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Error executing VBA: {e}")
        return json.dumps({
            "status": "error",
            "message": f"未知错误: {str(e)}",
            "logs": []
        }, ensure_ascii=False, indent=2)


@mcp.tool()
def row_column_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    start_index: int = None,
    count: int = 1
) -> str:
    """Unified row/column operation tool.
    
    统一的行列操作工具。
    
    Actions / 操作类型:
    - insert_rows: Insert rows (requires start_index)
                   插入行（需要 start_index，行号从1开始）
    - insert_cols: Insert columns (requires start_index)
                   插入列（需要 start_index，列号从1开始）
    - delete_rows: Delete rows (requires start_index)
                   删除行（需要 start_index）
    - delete_cols: Delete columns (requires start_index)
                   删除列（需要 start_index）
    """
    try:
        full_path = get_excel_path(filepath)
        result = row_column_operation_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            action=action,
            start_index=start_index,
            count=count
        )
        return result.get("message", str(result))
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error in row/column operation: {e}")
        return f"Error: {str(e)}"

def run_sse():
    """Run Excel MCP server in SSE mode."""
    # Assign value to EXCEL_FILES_PATH in SSE mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with SSE transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="sse")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_streamable_http():
    """Run Excel MCP server in streamable HTTP mode."""
    # Assign value to EXCEL_FILES_PATH in streamable HTTP mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with streamable HTTP transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="streamable-http")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_stdio():
    """Run Excel MCP server in stdio mode."""
    # No need to assign EXCEL_FILES_PATH in stdio mode
    
    try:
        logger.info("Starting Excel MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")