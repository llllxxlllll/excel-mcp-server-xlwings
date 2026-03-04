# Excel MCP 服务端工具说明

本文档介绍 Excel MCP 服务端中所有可用工具的详细说明。

## 工作簿操作

### create_workbook（创建工作簿）

创建新的 Excel 工作簿。

```python
create_workbook(filepath: str) -> str
```

- `filepath`：工作簿的创建路径
- 返回：成功信息及所创建文件的路径

### create_worksheet（创建工作表）

在已有工作簿中创建新工作表。

```python
create_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：新工作表名称
- 返回：成功信息

### get_workbook_metadata（获取工作簿元数据）

获取工作簿元数据，包括工作表列表、数据范围等。

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str
```

- `filepath`：Excel 文件路径
- `include_ranges`：是否包含范围信息
- 返回：工作簿元数据的字符串表示

## 数据操作

### write_data_to_excel（写入数据）

向 Excel 工作表写入数据。

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[Dict],
    start_cell: str = "A1"
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `data`：要写入的数据（二维列表，子列表为行）
- `start_cell`：起始单元格（默认 "A1"）
- 返回：成功信息

### read_data_from_excel（读取数据）

从 Excel 工作表读取数据，包含单元格元数据及验证规则。

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：源工作表名称
- `start_cell`：起始单元格（默认 "A1"）
- `end_cell`：可选，结束单元格
- `preview_only`：是否仅返回预览（超过 100 行时自动返回压缩摘要）
- 返回：数据的字符串表示（含验证等元数据）

## 格式操作

### format_range（设置区域格式）

对单元格区域应用格式（字体、颜色、边框、对齐、数字格式、条件格式等）。

```python
format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_cell`：区域起始单元格
- `end_cell`：可选，区域结束单元格
- 其他为各类格式选项（见参数列表）
- 返回：成功信息

### merge_cells（合并单元格）

合并指定区域的单元格。

```python
merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_cell`：区域起始单元格
- `end_cell`：区域结束单元格
- 返回：成功信息

### unmerge_cells（取消合并单元格）

取消先前已合并的单元格区域。

```python
unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_cell`：区域起始单元格
- `end_cell`：区域结束单元格
- 返回：成功信息

### get_merged_cells（获取合并单元格列表）

获取工作表中所有合并单元格区域。

```python
get_merged_cells(filepath: str, sheet_name: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- 返回：合并单元格的字符串表示

## 公式操作

### apply_formula（应用公式）

向单元格应用 Excel 公式。

```python
apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `cell`：目标单元格引用
- `formula`：要应用的 Excel 公式
- 返回：成功信息

### validate_formula_syntax（验证公式语法）

验证 Excel 公式语法，不实际写入单元格。

```python
validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `cell`：目标单元格引用
- `formula`：要验证的公式
- 返回：验证结果信息

## 图表操作

### create_chart（创建图表）

在工作表中创建图表。

```python
create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `data_range`：图表数据所在区域
- `chart_type`：图表类型（折线图、柱状图、饼图、散点图、面积图等）
- `target_cell`：图表放置的起始单元格
- `title`：可选，图表标题
- `x_axis`：可选，X 轴标签
- `y_axis`：可选，Y 轴标签
- 返回：成功信息

## 数据透视表操作

### create_pivot_table（创建数据透视表）

在工作表中创建数据透视表。

```python
create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    target_cell: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "mean"
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `data_range`：源数据区域
- `target_cell`：数据透视表放置的起始单元格
- `rows`：行标签字段
- `values`：值字段
- `columns`：可选，列标签字段
- `agg_func`：聚合方式（如 sum、count、average、max、min）
- 返回：成功信息

## 表格操作

### create_table（创建 Excel 表格）

从指定数据区域创建 Excel 原生表格。

```python
create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: str = None,
    table_style: str = "TableStyleMedium9"
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：工作表名称
- `data_range`：表格的单元格区域（如 "A1:D5"）
- `table_name`：可选，表格唯一名称
- `table_style`：可选，表格视觉样式
- 返回：成功信息

## 工作表操作

### copy_worksheet（复制工作表）

在工作簿内复制工作表。

```python
copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str
```

- `filepath`：Excel 文件路径
- `source_sheet`：要复制的工作表名称
- `target_sheet`：新工作表名称
- 返回：成功信息

### delete_worksheet（删除工作表）

从工作簿中删除工作表。

```python
delete_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：要删除的工作表名称
- 返回：成功信息

### rename_worksheet（重命名工作表）

重命名工作簿中的工作表。

```python
rename_worksheet(filepath: str, old_name: str, new_name: str) -> str
```

- `filepath`：Excel 文件路径
- `old_name`：当前工作表名称
- `new_name`：新工作表名称
- 返回：成功信息

## 范围操作

### copy_range（复制范围）

将单元格区域复制到另一位置。

```python
copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：源工作表名称
- `source_start`：源区域起始单元格
- `source_end`：源区域结束单元格
- `target_start`：粘贴目标起始单元格
- `target_sheet`：可选，目标工作表名称
- 返回：成功信息

### delete_range（删除范围）

删除单元格区域并移动其余单元格。

```python
delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_cell`：区域起始单元格
- `end_cell`：区域结束单元格
- `shift_direction`：移动方向（"up" 向上 或 "left" 向左）
- 返回：成功信息

### validate_excel_range（验证范围）

验证指定区域是否存在且格式正确。

```python
validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_cell`：区域起始单元格
- `end_cell`：可选，区域结束单元格
- 返回：验证结果信息

### get_data_validation_info（获取数据验证信息）

获取工作表中所有数据验证规则及元数据。

```python
get_data_validation_info(filepath: str, sheet_name: str) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- 返回：包含所有数据验证规则的 JSON 字符串，包括：
  - 验证类型（list、whole、decimal、date、time、textLength 等）
  - 运算符（between、notBetween、equal、greaterThan、lessThan 等）
  - 列表验证的允许值（从区域解析）
  - 数值/日期验证的公式约束
  - 应用验证的单元格区域
  - 提示信息与错误信息

**说明**：`read_data_from_excel` 在读取数据时会自动包含各单元格的验证元数据（如有）。

## 行列操作

### insert_rows（插入行）

从指定行开始插入一行或多行。

```python
insert_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_row`：插入起始行号（从 1 开始）
- `count`：插入行数（默认 1）
- 返回：成功信息

### insert_columns（插入列）

从指定列开始插入一列或多列。

```python
insert_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_col`：插入起始列号（从 1 开始）
- `count`：插入列数（默认 1）
- 返回：成功信息

### delete_sheet_rows（删除行）

从指定行开始删除一行或多行。

```python
delete_sheet_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_row`：删除起始行号（从 1 开始）
- `count`：删除行数（默认 1）
- 返回：成功信息

### delete_sheet_columns（删除列）

从指定列开始删除一列或多列。

```python
delete_sheet_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`：Excel 文件路径
- `sheet_name`：目标工作表名称
- `start_col`：删除起始列号（从 1 开始）
- `count`：删除列数（默认 1）
- 返回：成功信息

## VBA 执行操作

### execute_excel_vba（执行 Excel VBA）

在 Excel 文件上执行动态 VBA 代码。

```python
execute_excel_vba(
    filepath: str,
    vba_code: str,
    entry_sub_name: str = "Main"
) -> str
```

- `filepath`：Excel 文件路径（stdio 模式建议使用绝对路径）
- `vba_code`：完整的 VBA 代码字符串，必须包含与 `entry_sub_name` 对应的 Sub
- `entry_sub_name`：入口 Sub 过程名称（默认 "Main"）
- 返回：JSON 字符串，包含执行结果：
  - `status`：`"success"` 或 `"error"`
  - `message`：执行结果描述
  - `logs`：执行日志列表
  - `backup_path`：备份文件路径（成功时）

**安全机制：**

- 执行前会扫描 VBA 代码中的敏感关键词
- 禁止的关键词包括：Shell、Kill、CreateObject、FileSystemObject、SendKeys 等
- 执行前会自动创建备份文件
- 执行超时保护（默认 30 秒）
- 通过全局锁防止并发执行

**前置条件：**

使用本工具前，必须在 Excel 中启用对 VBA 工程对象模型的访问，否则会返回「安全检查失败：请在 Excel 信任中心勾选“信任对 VBA 工程对象模型的访问”」。

- 打开 Excel → **文件** → **选项** → **信任中心** → **信任中心设置**
- 进入 **宏设置**
- 勾选 **「信任对 VBA 工程对象模型的访问」**（Trust access to the VBA project object model）

**使用示例：**

```vba
' 简单单元格修改
Sub Main()
    Cells(1, 1).Value = "Hello World"
    Cells(1, 2).Value = Now()
End Sub
```

```vba
' 自定义入口
Sub ProcessData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1:A10").Value = 100
End Sub
```

**错误处理：**

工具会针对以下情况返回包含错误详情的 JSON：

- 安全检查失败（检测到敏感关键词）
- 文件未找到
- 未启用 VBA 信任
- VBA 运行时错误
- 执行超时
- Excel 忙碌（并发执行）

**禁止使用的关键词：**

出于安全考虑，以下关键词会被拦截：

- 系统命令：Shell、Kill、WScript、Scripting、SendKeys、AppActivate、Powershell、Cmd
- 文件操作：FileSystemObject、CreateObject、GetObject、DeleteFile、DeleteFolder、CopyFile、MoveFile、MkDir、RmDir、ChDir
- 网络相关：Adodb.stream、Shell.Application、WinHttp、XMLHTTP、UrlMon
- 环境访问：Environ、CallByName
- 自动运行宏：Workbook_Open、Auto_Open、Auto_Close
