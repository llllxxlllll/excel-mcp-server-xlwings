<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server" width="300"/>
</p>

[![PyPI 版本](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![总下载量](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![许可证: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery 徽章](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![安装 MCP 服务端](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)
基于 **Model Context Protocol (MCP)** 的 Excel 操作服务端，支持对 Excel 文件进行实时编辑。本服务端基于 **xlwings**，可直接连接本地已打开的 Excel 实例，在 Excel 窗口中即时查看修改结果。

## 环境要求

- **Microsoft Excel**（Windows 或 macOS）——xlwings 运行所必需
- **Python** >= 3.10
- **Windows**：需已安装 Excel，并可通过 COM 访问
- **macOS**：需已安装 Excel，并可通过 AppleScript 访问

## 功能特性

- 📊 **实时编辑**：在 Excel 文件打开状态下编辑，修改即时可见
- 📈 **数据操作**：公式、格式、图表、数据透视表、Excel 表格
- 🔍 **数据验证**：内置对区域、公式及数据完整性的验证
- 🎨 **格式设置**：字体、颜色、边框、对齐、条件格式等
- 📋 **表格操作**：创建与管理 Excel 表格，支持自定义样式
- 📊 **图表创建**：支持折线图、柱状图、饼图、散点图等多种图表
- 🔄 **数据透视表**：创建动态数据透视表进行数据分析
- 🔧 **工作表管理**：复制、重命名、删除工作表
- 🔌 **多种传输方式**：stdio、SSE（已弃用）、可流式 HTTP
- 🌐 **本地与远程**：支持本地运行或作为远程服务
- ⚡ **实时连接**：连接已打开的 Excel 文件，无需先关闭

## 使用方式

服务端支持三种传输方式：

### 1. Stdio 传输（本地使用）

```bash
uvx excel-mcp-server stdio
```

在 MCP 客户端中配置示例：

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

### 2. SSE 传输（Server-Sent Events，已弃用）

```bash
uvx excel-mcp-server sse
```

**SSE 连接配置示例**：

```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. 可流式 HTTP 传输（推荐用于远程连接）

```bash
uvx excel-mcp-server streamable-http
```

**可流式 HTTP 连接配置示例**：

```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## 环境变量与文件路径

### SSE 与可流式 HTTP 传输

使用 **SSE 或可流式 HTTP** 时，必须在**服务端**设置环境变量 **`EXCEL_FILES_PATH`**，用于指定 Excel 文件的读写目录。

- 未设置时，默认为 `./excel_files`。

可通过 **`FASTMCP_PORT`** 指定服务端监听端口（未设置时默认为 `8017`）。

- Windows PowerShell 示例：

  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server streamable-http
  ```

- Linux / macOS 示例：

  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

### Stdio 传输

使用 **stdio** 时，文件路径由每次工具调用时传入，**无需**在服务端设置 `EXCEL_FILES_PATH`，服务端会使用客户端提供的路径。

## 可用工具

服务端提供完整的 Excel 操作工具集，完整说明见 [TOOLS.md](TOOLS.md)。

主要能力包括：

- **工作簿**：创建、获取元数据
- **工作表**：创建、复制、重命名、删除、列表
- **数据**：读取、写入、公式、验证
- **格式**：字体、颜色、边框、对齐、条件格式、合并单元格
- **范围**：复制、删除、验证
- **行列**：插入/删除行与列
- **表格**：创建 Excel 表格
- **图表**：创建、列表、删除、样式
- **数据透视表**：创建
- **VBA**：在 Excel 文件上执行 VBA 代码（含安全与备份说明）

## 许可证

MIT License，详见 [LICENSE](LICENSE)。
