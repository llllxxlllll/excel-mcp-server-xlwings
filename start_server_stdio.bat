@echo off
REM Excel MCP Server - Stdio 模式启动脚本
REM 此脚本用于启动 stdio 模式的 MCP 服务器

echo 正在启动 Excel MCP Server (Stdio 模式)...
echo.

REM 设置 Python UTF-8 模式
set PYTHONUTF8=1

REM 启动服务器
uvx excel-mcp-server stdio

pause


