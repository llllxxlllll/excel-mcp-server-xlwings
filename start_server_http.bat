@echo off
REM Excel MCP Server - Streamable HTTP 模式启动脚本
REM 此脚本用于启动 HTTP 模式的 MCP 服务器

echo 正在启动 Excel MCP Server (Streamable HTTP 模式)...
echo.

REM 设置环境变量
set EXCEL_FILES_PATH=E:\res_program\excel-mcp-server\excel_files
set FASTMCP_PORT=8017
set PYTHONUTF8=1

REM 创建 Excel 文件目录（如果不存在）
if not exist "%EXCEL_FILES_PATH%" mkdir "%EXCEL_FILES_PATH%"

echo Excel 文件路径: %EXCEL_FILES_PATH%
echo 服务器端口: %FASTMCP_PORT%
echo.
echo 服务器启动后，可在 Cursor 中使用以下配置连接:
echo {
echo   "name": "excel-mcp-server",
echo   "type": "http",
echo   "url": "http://localhost:%FASTMCP_PORT%/mcp"
echo }
echo.

REM 启动服务器
uvx excel-mcp-server streamable-http

pause


