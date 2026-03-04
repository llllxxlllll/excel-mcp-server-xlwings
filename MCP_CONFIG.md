# Excel MCP Server 配置指南

## 项目信息

- **项目路径**: `E:\res_program\excel-mcp-server-xlwings`
- **Python 路径**: `C:\anaconda3\python.exe` (通过 `--python` 参数明确指定)
- **Python 版本**: 3.12.7 (由 `.python-version` 文件指定)
- **服务器命令**: `uv run --python C:\anaconda3\python.exe excel-mcp-server`

## 配置方式

### 方式 1: 本地开发模式（推荐）

使用 `uv run --directory` 直接运行本地项目代码，适合开发调试。

#### Kiro 配置

在 `.kiro/settings/mcp.json` 中配置：

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "E:\\res_program\\excel-mcp-server-xlwings",
        "--python",
        "C:\\anaconda3\\python.exe",
        "excel-mcp-server",
        "stdio"
      ],
      "env": {
        "PYTHONUTF8": "1"
      },
      "disabled": false,
      "autoApprove": [
        "create_workbook",
        "write_data_to_excel",
        "read_data_from_excel",
        "apply_formula",
        "get_workbook_metadata",
        "create_worksheet",
        "rename_worksheet",
        "create_chart",
        "list_charts",
        "delete_chart",
        "update_chart_style",
        "chart_operation",
        "worksheet_operation",
        "formula_operation",
        "merge_cell_operation",
        "row_column_operation",
        "range_operation"
      ]
    }
  }
}
```

**说明：**
- `uv run --directory <path>` 会在指定目录下运行命令，自动使用该目录的虚拟环境和依赖
- `--python` 参数明确指定使用的 Python 解释器路径，确保环境一致性
- `autoApprove` 列表包含了常用的 Excel 操作工具，无需手动确认
- 修改本地代码后，重新连接 MCP 即可生效，无需重新发布

#### Cursor 配置

```json
{
  "name": "excel-mcp-server",
  "type": "stdio",
  "command": "uv",
  "args": [
    "run",
    "--directory",
    "E:\\res_program\\excel-mcp-server-xlwings",
    "--python",
    "C:\\anaconda3\\python.exe",
    "excel-mcp-server",
    "stdio"
  ],
  "env": {
    "PYTHONUTF8": "1"
  }
}
```

### 方式 2: PyPI 发布版本模式

使用 `uvx` 从 PyPI 下载并运行已发布的包，适合生产使用。

#### Kiro/Cursor 配置

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

**注意：** 此模式运行的是 PyPI 上的发布版本，不是本地代码

### 方式 3: Streamable HTTP 模式（推荐用于远程连接）

适合需要远程访问的场景。

#### 启动服务器

在 PowerShell 中执行：

```powershell
# 设置 Excel 文件存储路径（可选，默认为 ./excel_files）
$env:EXCEL_FILES_PATH="E:\res_program\excel-mcp-server-xlwings\excel_files"

# 设置端口（可选，默认为 8017）
$env:FASTMCP_PORT="8017"

# 启动服务器（本地开发版本）
uv run --directory E:\res_program\excel-mcp-server-xlwings excel-mcp-server streamable-http
```

#### Kiro/Cursor 配置

```json
{
  "name": "excel-mcp-server",
  "type": "http",
  "url": "http://localhost:8017/mcp",
  "env": {}
}
```

### 方式 4: SSE 模式（已弃用）

不推荐使用，但如需使用：

```powershell
$env:EXCEL_FILES_PATH="E:\res_program\excel-mcp-server-xlwings\excel_files"
$env:FASTMCP_PORT="8000"
uv run --directory E:\res_program\excel-mcp-server-xlwings excel-mcp-server sse
```

Kiro/Cursor 配置：

```json
{
  "name": "excel-mcp-server",
  "type": "sse",
  "url": "http://localhost:8000/sse",
  "env": {}
}
```

## 环境变量说明

### EXCEL_FILES_PATH
- **说明**: Excel 文件的存储路径（仅 SSE 和 Streamable HTTP 模式需要）
- **默认值**: `./excel_files`
- **示例**: `E:\res_program\excel-mcp-server-xlwings\excel_files`

### FASTMCP_PORT
- **说明**: 服务器监听端口（仅 SSE 和 Streamable HTTP 模式需要）
- **默认值**: `8017`（Streamable HTTP）或 `8000`（SSE）
- **示例**: `8017`

### PYTHONUTF8
- **说明**: 启用 Python UTF-8 模式
- **推荐值**: `1`

## 测试配置

### 测试本地开发版本（Stdio 模式）

```powershell
# 方式 1: 使用 uv run 指定 Python 路径
uv run --directory E:\res_program\excel-mcp-server-xlwings --python C:\anaconda3\python.exe excel-mcp-server stdio

# 方式 2: 进入项目目录后运行
cd E:\res_program\excel-mcp-server-xlwings
uv run --python C:\anaconda3\python.exe excel-mcp-server stdio

# 方式 3: 使用项目默认 Python（依赖 .python-version 文件）
cd E:\res_program\excel-mcp-server-xlwings
uv run excel-mcp-server stdio
```

### 测试 PyPI 发布版本

```powershell
uvx excel-mcp-server stdio
```

### 测试 Streamable HTTP 模式

```powershell
cd E:\res_program\excel-mcp-server-xlwings
$env:EXCEL_FILES_PATH="E:\res_program\excel-mcp-server-xlwings\excel_files"
uv run excel-mcp-server streamable-http
```

然后在浏览器中访问 `http://localhost:8017` 查看服务器状态。

## 可用工具

服务器提供完整的 Excel 操作工具，包括：

- 工作簿操作：创建、打开、保存
- 工作表操作：读取、写入、格式化
- 图表创建：支持多种图表类型
- 数据透视表：创建动态透视表
- 表格操作：创建和管理 Excel 表格
- 数据验证：范围验证、公式验证

详细工具列表请参考 [TOOLS.md](TOOLS.md)

## 故障排除

### 问题 1: uvx 命令未找到

**解决方案**: 确保已安装 `uv` 工具：

```powershell
pip install uv
```

### 问题 2: 端口被占用

**解决方案**: 更改 `FASTMCP_PORT` 环境变量为其他端口。

### 问题 3: 文件路径错误

**解决方案**: 
- Stdio 模式：确保在工具调用时提供正确的绝对路径
- HTTP/SSE 模式：确保 `EXCEL_FILES_PATH` 环境变量设置正确

## 相关链接

- [项目 GitHub](https://github.com/haris-musa/excel-mcp-server)
- [工具文档](TOOLS.md)
- [项目 README](README.md)


