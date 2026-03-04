# Excel MCP Server 项目架构解析

本文档详细解读 excel-mcp-server 项目的实现方式，帮助理解 MCP (Model Context Protocol) 服务器的构建方法。

## 1. 项目概述

这是一个基于 **xlwings** 的 Excel MCP 服务器，允许 AI 助手（如 Cursor、Claude Desktop、Kiro）通过 MCP 协议直接操作 Excel 文件。

### 核心特点
- 使用 **xlwings** 库实现实时 Excel 编辑（需要安装 Microsoft Excel）
- 基于 **FastMCP** 框架快速构建 MCP 服务器
- 支持三种传输模式：Stdio、Streamable HTTP、SSE

## 2. 技术栈

```
┌─────────────────────────────────────────────────────────┐
│                    MCP 客户端                            │
│              (Cursor / Claude Desktop / Kiro)           │
└─────────────────────┬───────────────────────────────────┘
                      │ MCP 协议
                      ▼
┌─────────────────────────────────────────────────────────┐
│                  FastMCP 框架                            │
│            (mcp[cli] + fastmcp)                         │
├─────────────────────────────────────────────────────────┤
│                  Typer CLI                              │
│           (命令行入口管理)                               │
├─────────────────────────────────────────────────────────┤
│                  xlwings                                │
│           (Excel COM 自动化)                            │
└─────────────────────┬───────────────────────────────────┘
                      │ COM 接口
                      ▼
┌─────────────────────────────────────────────────────────┐
│              Microsoft Excel                            │
└─────────────────────────────────────────────────────────┘
```

### 依赖说明

| 依赖包 | 版本要求 | 用途 |
|--------|----------|------|
| `mcp[cli]` | >=1.10.1 | MCP 协议核心库 |
| `fastmcp` | >=2.0.0,<3.0.0 | MCP 服务器快速开发框架 |
| `xlwings` | >=0.30.0 | Excel COM 自动化操作 |
| `typer` | >=0.16.0 | CLI 命令行框架 |

## 3. 项目结构

```
excel-mcp-server/
├── src/excel_mcp/           # 核心源码目录
│   ├── __main__.py          # CLI 入口点
│   ├── server.py            # MCP 服务器定义 & 工具注册
│   ├── xw_helper.py         # xlwings 辅助函数封装
│   ├── workbook.py          # 工作簿操作
│   ├── sheet.py             # 工作表操作
│   ├── data.py              # 数据读写
│   ├── formatting.py        # 格式化操作
│   ├── calculations.py      # 公式计算
│   ├── chart.py             # 图表创建
│   ├── pivot.py             # 数据透视表
│   ├── tables.py            # Excel 表格
│   ├── validation.py        # 范围/公式验证
│   ├── cell_validation.py   # 单元格数据验证
│   ├── cell_utils.py        # 单元格工具函数
│   └── exceptions.py        # 自定义异常
├── pyproject.toml           # 项目配置 & 依赖
├── tests/                   # 测试目录
└── excel_files/             # Excel 文件存储目录
```

## 4. 核心实现解析

### 4.1 入口点 (`__main__.py`)

使用 **Typer** 创建 CLI 应用，提供三种启动命令：

```python
import typer
from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(help="Excel MCP Server")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    run_stdio()

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    run_streamable_http()

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    run_sse()
```

### 4.2 MCP 服务器定义 (`server.py`)

#### 初始化 FastMCP 实例

```python
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "excel-mcp",
    host=os.environ.get("FASTMCP_HOST", "0.0.0.0"),
    port=int(os.environ.get("FASTMCP_PORT", "8017")),
    instructions="Excel MCP Server for manipulating Excel files"
)
```

#### 工具注册模式

使用 `@mcp.tool()` 装饰器注册 MCP 工具：

```python
@mcp.tool()
def create_workbook(filepath: str) -> str:
    """Create new Excel workbook."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_workbook as create_workbook_impl
        create_workbook_impl(full_path)
        return f"Created workbook at {full_path}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
```

**关键设计点：**
1. 函数签名定义工具参数
2. docstring 作为工具描述
3. 返回字符串作为工具输出
4. 异常处理确保稳定性

#### 路径处理策略

```python
def get_excel_path(filename: str) -> str:
    """根据运行模式处理文件路径"""
    # 绝对路径直接返回
    if os.path.isabs(filename):
        return filename
    
    # Stdio 模式：必须使用绝对路径
    if EXCEL_FILES_PATH is None:
        raise ValueError("必须使用绝对路径")
    
    # HTTP/SSE 模式：基于 EXCEL_FILES_PATH 解析相对路径
    return os.path.join(EXCEL_FILES_PATH, filename)
```

#### 三种传输模式

```python
def run_stdio():
    """Stdio 模式 - 本地开发推荐"""
    # 不设置 EXCEL_FILES_PATH，要求客户端提供绝对路径
    mcp.run(transport="stdio")

def run_streamable_http():
    """Streamable HTTP 模式 - 远程连接推荐"""
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    mcp.run(transport="streamable-http")

def run_sse():
    """SSE 模式 - 已弃用"""
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    mcp.run(transport="sse")
```

### 4.3 xlwings 辅助层 (`xw_helper.py`)

封装 xlwings 操作，提供统一接口：

```python
def get_app(visible: bool = None) -> xw.App:
    """获取或创建 Excel 应用程序实例"""
    apps = xw.apps
    if apps:
        app = apps.active
        if app is not None:
            return app
    return xw.App(visible=visible)

def get_workbook(filepath: str, create_if_missing: bool = False) -> xw.Book:
    """获取工作簿，支持连接已打开的文件"""
    # 检查文件是否已在 Excel 中打开
    for app in xw.apps:
        for book in app.books:
            if Path(book.fullname).resolve() == Path(filepath).resolve():
                return book  # 复用已打开的工作簿
    
    # 打开或创建工作簿
    if Path(filepath).exists():
        return get_app().books.open(filepath)
    elif create_if_missing:
        book = get_app().books.add()
        book.save(filepath)
        return book
```

**设计亮点：**
- 复用已打开的 Excel 实例，避免重复打开
- 统一的异常处理
- 支持自动创建工作簿

### 4.4 功能模块分层

每个功能模块遵循相同模式：

```python
# workbook.py 示例
def create_workbook(filepath: str, sheet_name: str = "Sheet1") -> dict:
    """创建新的 Excel 工作簿"""
    try:
        path = Path(filepath).resolve()
        if path.exists():
            raise WorkbookError(f"文件已存在: {filepath}")
        
        app = get_app(visible=True)
        wb = app.books.add()
        wb.sheets[0].name = sheet_name
        wb.save(str(path))
        
        return {
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
            "filepath": str(path)
        }
    except Exception as e:
        raise WorkbookError(f"创建工作簿失败: {e}")
```

## 5. 工具清单

服务器提供 **25+ 个工具**，分为以下类别：

### 工作簿操作
- `create_workbook` - 创建工作簿
- `get_workbook_metadata` - 获取工作簿元数据

### 工作表操作
- `create_worksheet` - 创建工作表
- `copy_worksheet` - 复制工作表
- `delete_worksheet` - 删除工作表
- `rename_worksheet` - 重命名工作表

### 数据操作
- `read_data_from_excel` - 读取数据（含元数据）
- `write_data_to_excel` - 写入数据

### 格式化操作
- `format_range` - 格式化单元格范围
- `merge_cells` / `unmerge_cells` - 合并/取消合并单元格
- `get_merged_cells` - 获取合并单元格信息

### 公式操作
- `apply_formula` - 应用公式
- `validate_formula_syntax` - 验证公式语法

### 图表 & 透视表
- `create_chart` - 创建图表
- `create_pivot_table` - 创建数据透视表
- `create_table` - 创建 Excel 表格

### 范围操作
- `copy_range` - 复制范围
- `delete_range` - 删除范围
- `validate_excel_range` - 验证范围
- `get_data_validation_info` - 获取数据验证规则

### 行列操作
- `insert_rows` / `insert_columns` - 插入行/列
- `delete_sheet_rows` / `delete_sheet_columns` - 删除行/列

## 6. 配置指南

### 6.1 Stdio 模式（推荐本地开发）

**Kiro/Cursor 配置：**

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"],
      "env": {
        "PYTHONUTF8": "1"
      }
    }
  }
}
```

**使用方式：** 工具调用时提供绝对路径

```
filepath: "E:\\data\\report.xlsx"
```

### 6.2 Streamable HTTP 模式（推荐远程连接）

**启动服务器：**

```powershell
$env:EXCEL_FILES_PATH = "E:\res_program\excel-mcp-server\excel_files"
$env:FASTMCP_PORT = "8017"
uvx excel-mcp-server streamable-http
```

**Kiro/Cursor 配置：**

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "type": "http",
      "url": "http://localhost:8017/mcp"
    }
  }
}
```

**使用方式：** 可使用相对路径（基于 EXCEL_FILES_PATH）

```
filepath: "report.xlsx"
```

### 6.3 环境变量

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| `EXCEL_FILES_PATH` | Excel 文件存储目录 | `./excel_files` |
| `FASTMCP_PORT` | HTTP 服务端口 | `8017` |
| `FASTMCP_HOST` | HTTP 服务主机 | `0.0.0.0` |
| `PYTHONUTF8` | Python UTF-8 模式 | - |

## 7. 开发新 MCP 服务器的关键步骤

基于本项目的实现，开发新 MCP 服务器的步骤：

### Step 1: 项目初始化

```toml
# pyproject.toml
[project]
name = "your-mcp-server"
dependencies = [
    "mcp[cli]>=1.10.1",
    "fastmcp>=2.0.0,<3.0.0",
    "typer>=0.16.0"
]

[project.scripts]
your-mcp-server = "your_package.__main__:app"
```

### Step 2: 创建 CLI 入口

```python
# __main__.py
import typer
from .server import run_stdio

app = typer.Typer()

@app.command()
def stdio():
    run_stdio()

if __name__ == "__main__":
    app()
```

### Step 3: 定义 MCP 服务器和工具

```python
# server.py
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("your-server-name")

@mcp.tool()
def your_tool(param1: str, param2: int = 10) -> str:
    """工具描述 - 会显示给 AI"""
    # 实现逻辑
    return "结果"

def run_stdio():
    mcp.run(transport="stdio")
```

### Step 4: 发布和使用

```bash
# 发布到 PyPI
uv build
uv publish

# 使用
uvx your-mcp-server stdio
```

## 8. 故障排除

### 问题 1: Excel 未安装

```
ExcelNotFoundError: 无法启动 Excel 应用程序
```

**解决方案：** 安装 Microsoft Excel（xlwings 需要 Excel COM 接口）

### 问题 2: uvx 命令未找到

```bash
pip install uv
# 或
pipx install uv
```

### 问题 3: Stdio 模式路径错误

```
ValueError: Invalid filename, must be an absolute path
```

**解决方案：** Stdio 模式必须使用绝对路径

### 问题 4: 日志查看

日志文件位置：`{项目根目录}/excel-mcp.log`

## 9. 相关链接

- [MCP 官方文档](https://modelcontextprotocol.io/)
- [FastMCP 文档](https://github.com/jlowin/fastmcp)
- [xlwings 文档](https://docs.xlwings.org/)
- [项目 GitHub](https://github.com/haris-musa/excel-mcp-server)
