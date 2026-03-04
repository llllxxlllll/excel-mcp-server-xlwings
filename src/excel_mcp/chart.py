# -*- coding: utf-8 -*-
"""图表模块 - 使用 xlwings 实现

本模块提供 Excel 图表创建功能。
"""

from typing import Any, Optional, Dict
import logging

from .xw_helper import get_workbook, get_sheet, save_workbook, WorkbookContext
from .exceptions import ValidationError, ChartError

logger = logging.getLogger(__name__)

# xlwings 图表类型常量映射
CHART_TYPE_MAP = {
    "line": -4120,      # xlLine
    "column": 51,       # xlColumnClustered (纵向柱状图)
    "bar": 57,          # xlBarClustered (横向条形图)
    "pie": 5,           # xlPie
    "scatter": -4169,   # xlXYScatter
    "area": 1,          # xlArea
}


def create_chart_in_sheet(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    style: Optional[Dict] = None
) -> Dict[str, Any]:
    """在工作表中创建图表
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        data_range: 数据范围，如 'A1:C10'
        chart_type: 图表类型 (line, bar, pie, scatter, area)
        target_cell: 图表放置位置
        title: 图表标题
        x_axis: X 轴标题
        y_axis: Y 轴标题
        style: 样式配置
        
    Returns:
        包含操作结果的字典
    """
    try:
        # 验证图表类型
        chart_type_lower = chart_type.lower()
        if chart_type_lower not in CHART_TYPE_MAP:
            raise ValidationError(
                f"Unsupported chart type: {chart_type}. "
                f"Supported types: {', '.join(CHART_TYPE_MAP.keys())}"
            )
        
        with WorkbookContext(filepath) as ctx:
            wb = ctx.wb
            sheet = ctx.get_sheet(sheet_name)
            
            # 解析数据范围
            if "!" in data_range:
                range_sheet_name, cell_range = data_range.split("!")
                data_sheet = ctx.get_sheet(range_sheet_name)
                source_range = data_sheet.range(cell_range)
            else:
                source_range = sheet.range(data_range)
            
            # 获取目标位置
            target_range = sheet.range(target_cell)
            
            # 创建图表
            chart = sheet.charts.add(
                left=target_range.left,
                top=target_range.top,
                width=400,
                height=250
            )
            
            # 设置图表类型
            xl_chart_type = CHART_TYPE_MAP[chart_type_lower]
            
            # 设置数据源
            chart.set_source_data(source_range)
            
            # 设置图表属性 - 使用 xlwings 原生属性
            try:
                chart.chart_type = chart_type_lower
            except Exception:
                # 如果 xlwings 原生方式失败，尝试通过 COM API
                pass
            
            # 通过 COM API 设置详细属性
            try:
                # 获取 COM 对象 - 使用正确的访问方式
                chart_obj = chart.api
                if hasattr(chart_obj, 'Chart'):
                    chart_api = chart_obj.Chart
                elif hasattr(chart_obj, '__getitem__'):
                    chart_api = chart_obj[1]
                else:
                    chart_api = chart_obj
                
                chart_api.ChartType = xl_chart_type
                
                if title:
                    chart_api.HasTitle = True
                    chart_api.ChartTitle.Text = title
                
                # 设置轴标题
                if x_axis and chart_type_lower not in ['pie']:
                    try:
                        chart_api.Axes(1).HasTitle = True  # xlCategory = 1
                        chart_api.Axes(1).AxisTitle.Text = x_axis
                    except Exception:
                        pass
                
                if y_axis and chart_type_lower not in ['pie']:
                    try:
                        chart_api.Axes(2).HasTitle = True  # xlValue = 2
                        chart_api.Axes(2).AxisTitle.Text = y_axis
                    except Exception:
                        pass
                        
            except Exception as e:
                logger.warning(f"Failed to set some chart properties: {e}")
            
            # 保存工作簿
            ctx.save()
            
            return {
                "message": f"{chart_type.capitalize()} chart created successfully",
                "details": {
                    "type": chart_type,
                    "location": target_cell,
                    "data_range": data_range
                }
            }
        
    except (ValidationError, ChartError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        import traceback
        logger.error(f"Unexpected error creating chart: {e}\n{traceback.format_exc()}")
        raise ChartError(f"Unexpected error creating chart: {str(e)}")


def list_charts_in_sheet(
    filepath: str,
    sheet_name: str
) -> Dict[str, Any]:
    """列出工作表中的所有图表
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        
    Returns:
        包含图表列表的字典
    """
    try:
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            charts_info = []
            for i, chart in enumerate(sheet.charts):
                chart_info = {
                    "index": i,
                    "name": chart.name,
                    "left": chart.left,
                    "top": chart.top,
                    "width": chart.width,
                    "height": chart.height,
                }
                # 尝试获取图表标题
                try:
                    chart_obj = chart.api
                    if hasattr(chart_obj, '__getitem__'):
                        chart_api = chart_obj[1]
                    else:
                        chart_api = chart_obj
                    if chart_api.HasTitle:
                        chart_info["title"] = chart_api.ChartTitle.Text
                except Exception:
                    pass
                charts_info.append(chart_info)
            
            return {
                "sheet_name": sheet_name,
                "chart_count": len(charts_info),
                "charts": charts_info
            }
        
    except Exception as e:
        logger.error(f"Error listing charts: {e}")
        raise ChartError(f"Error listing charts: {str(e)}")


def delete_chart_in_sheet(
    filepath: str,
    sheet_name: str,
    chart_index: int = None,
    chart_name: str = None
) -> Dict[str, Any]:
    """删除工作表中的图表
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        chart_index: 图表索引（从0开始）
        chart_name: 图表名称（与 chart_index 二选一）
        
    Returns:
        操作结果
    """
    try:
        if chart_index is None and chart_name is None:
            raise ValidationError("必须提供 chart_index 或 chart_name")
        
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            chart_to_delete = None
            
            if chart_name is not None:
                # 按名称查找
                for chart in sheet.charts:
                    if chart.name == chart_name:
                        chart_to_delete = chart
                        break
                if chart_to_delete is None:
                    raise ValidationError(f"未找到名为 '{chart_name}' 的图表")
            else:
                # 按索引查找
                charts_list = list(sheet.charts)
                if chart_index < 0 or chart_index >= len(charts_list):
                    raise ValidationError(
                        f"图表索引 {chart_index} 超出范围，"
                        f"当前共有 {len(charts_list)} 个图表"
                    )
                chart_to_delete = charts_list[chart_index]
            
            deleted_name = chart_to_delete.name
            chart_to_delete.delete()
            ctx.save()
            
            return {
                "message": f"图表 '{deleted_name}' 已删除",
                "deleted_chart": deleted_name
            }
        
    except (ValidationError, ChartError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Error deleting chart: {e}")
        raise ChartError(f"Error deleting chart: {str(e)}")



def update_chart_style(
    filepath: str,
    sheet_name: str,
    chart_index: int = None,
    chart_name: str = None,
    font_name: str = None,
    font_size: int = None,
    title_font_size: int = None
) -> Dict[str, Any]:
    """更新图表样式
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        chart_index: 图表索引
        chart_name: 图表名称
        font_name: 字体名称
        font_size: 字体大小
        title_font_size: 标题字体大小
        
    Returns:
        操作结果
    """
    try:
        if chart_index is None and chart_name is None:
            raise ValidationError("必须提供 chart_index 或 chart_name")
        
        with WorkbookContext(filepath) as ctx:
            sheet = ctx.get_sheet(sheet_name)
            
            chart_to_update = None
            
            if chart_name is not None:
                for chart in sheet.charts:
                    if chart.name == chart_name:
                        chart_to_update = chart
                        break
                if chart_to_update is None:
                    raise ValidationError(f"未找到名为 '{chart_name}' 的图表")
            else:
                charts_list = list(sheet.charts)
                if chart_index < 0 or chart_index >= len(charts_list):
                    raise ValidationError(f"图表索引 {chart_index} 超出范围")
                chart_to_update = charts_list[chart_index]
            
            # 获取 COM API
            chart_obj = chart_to_update.api
            if hasattr(chart_obj, '__getitem__'):
                chart_api = chart_obj[1]
            else:
                chart_api = chart_obj
            
            # 设置字体
            if font_name or font_size:
                # 设置图表区域字体
                try:
                    if font_name:
                        chart_api.ChartArea.Font.Name = font_name
                    if font_size:
                        chart_api.ChartArea.Font.Size = font_size
                except Exception as e:
                    logger.warning(f"设置图表区域字体失败: {e}")
                
                # 设置绘图区字体
                try:
                    if font_name:
                        chart_api.PlotArea.Font.Name = font_name
                except Exception:
                    pass
                
                # 设置图例字体
                try:
                    if chart_api.HasLegend:
                        if font_name:
                            chart_api.Legend.Font.Name = font_name
                        if font_size:
                            chart_api.Legend.Font.Size = font_size
                except Exception:
                    pass
                
                # 设置坐标轴字体
                try:
                    for axis_type in [1, 2]:  # xlCategory=1, xlValue=2
                        try:
                            axis = chart_api.Axes(axis_type)
                            if font_name:
                                axis.TickLabels.Font.Name = font_name
                            if font_size:
                                axis.TickLabels.Font.Size = font_size
                        except Exception:
                            pass
                except Exception:
                    pass
            
            # 设置标题字体
            if chart_api.HasTitle:
                try:
                    if font_name:
                        chart_api.ChartTitle.Font.Name = font_name
                    if title_font_size:
                        chart_api.ChartTitle.Font.Size = title_font_size
                    elif font_size:
                        chart_api.ChartTitle.Font.Size = font_size + 2
                except Exception as e:
                    logger.warning(f"设置标题字体失败: {e}")
            
            ctx.save()
            
            return {
                "message": f"图表 '{chart_to_update.name}' 样式已更新",
                "updated_chart": chart_to_update.name
            }
        
    except (ValidationError, ChartError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Error updating chart style: {e}")
        raise ChartError(f"Error updating chart style: {str(e)}")



def chart_operation(
    filepath: str,
    sheet_name: str,
    action: str,
    # create 参数
    data_range: str = None,
    chart_type: str = None,
    target_cell: str = None,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    # delete/style 参数
    chart_index: int = None,
    chart_name: str = None,
    # style 参数
    font_name: str = None,
    font_size: int = None,
    title_font_size: int = None
) -> Dict[str, Any]:
    """统一的图表操作接口
    
    Args:
        filepath: Excel 文件路径
        sheet_name: 工作表名称
        action: 操作类型 (create, list, delete, style)
        
        create 操作需要:
            data_range: 数据范围
            chart_type: 图表类型 (line, column, bar, pie, scatter, area)
            target_cell: 图表放置位置
            title: 图表标题（可选）
            x_axis: X轴标题（可选）
            y_axis: Y轴标题（可选）
            
        delete 操作需要:
            chart_index 或 chart_name
            
        style 操作需要:
            chart_index 或 chart_name
            font_name: 字体名称（可选）
            font_size: 字体大小（可选）
            title_font_size: 标题字体大小（可选）
            
    Returns:
        操作结果字典
    """
    action = action.lower()
    
    if action == "create":
        if not all([data_range, chart_type, target_cell]):
            raise ValidationError(
                "create 操作需要: data_range, chart_type, target_cell"
            )
        return create_chart_in_sheet(
            filepath=filepath,
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis
        )
    
    elif action == "list":
        return list_charts_in_sheet(
            filepath=filepath,
            sheet_name=sheet_name
        )
    
    elif action == "delete":
        return delete_chart_in_sheet(
            filepath=filepath,
            sheet_name=sheet_name,
            chart_index=chart_index,
            chart_name=chart_name
        )
    
    elif action == "style":
        return update_chart_style(
            filepath=filepath,
            sheet_name=sheet_name,
            chart_index=chart_index,
            chart_name=chart_name,
            font_name=font_name,
            font_size=font_size,
            title_font_size=title_font_size
        )
    
    else:
        raise ValidationError(
            f"不支持的操作: {action}。支持的操作: create, list, delete, style"
        )
