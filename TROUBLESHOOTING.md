# Excel MCP Server 故障排除指南

## 当前问题分析

根据日志文件，发现以下问题：

### 问题 1: dotnet 命令未找到
```
'dotnet' is not recognized as an internal or external command
```

**原因**: .NET 8 SDK 未安装或不在 PATH 环境变量中

### 问题 2: haris-musa/excel-mcp-server 的 I/O 错误
```
ValueError: I/O operation on closed file.
```

**状态**: 虽然有此错误，但服务器实际上可以工作（找到了 25 个工具）

## 解决方案

### 方案 A: 安装 .NET 8 使用 sbroenne/excel-mcp-server（推荐，如果已安装 Excel）

#### 步骤 1: 安装 .NET 8 SDK

**方法 1: 使用 winget（推荐）**
```powershell
winget install Microsoft.DotNet.SDK.8
```

**方法 2: 手动下载安装**
1. 访问：https://dotnet.microsoft.com/download/dotnet/8.0
2. 下载 .NET 8 SDK
3. 运行安装程序
4. 重启 PowerShell 或 Cursor

**方法 3: 使用 Chocolatey（如果已安装）**
```powershell
choco install dotnet-8.0-sdk
```

#### 步骤 2: 验证安装

```powershell
# 重新打开 PowerShell 后运行
dotnet --version
# 应该显示 8.0.x 或更高版本
```

#### 步骤 3: 安装 ExcelMcp MCP Server

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

#### 步骤 4: 验证工具安装

```powershell
dotnet tool list --global | Select-String "ExcelMcp"
```

#### 步骤 5: 重启 Cursor

安装完成后，重启 Cursor，配置应该可以正常工作。

---

### 方案 B: 继续使用 haris-musa/excel-mcp-server（无需 Excel，已可工作）

虽然日志显示有 I/O 错误，但服务器实际上已经可以工作（找到了 25 个工具）。这个错误可能是非致命的。

#### 更新配置使用 haris-musa 版本

将 `mcp.json` 中的 excel 配置改为：

```json
"excel": {
  "command": "uvx",
  "args": ["excel-mcp-server", "stdio"],
  "env": {
    "PYTHONUTF8": "1"
  }
}
```

#### 优点
- ✅ 无需安装 .NET
- ✅ 无需安装 Microsoft Excel
- ✅ 跨平台支持
- ✅ 已经可以工作（25 个工具可用）

#### 缺点
- ⚠️ 有一个 I/O 错误（但似乎不影响功能）
- ⚠️ 功能可能不如需要 Excel 的版本强大

---

## 推荐方案

### 如果您已安装 Microsoft Excel 2016+
→ **使用方案 A**（sbroenne/excel-mcp-server）
- 功能更强大
- 直接使用 Excel 的所有功能
- 需要安装 .NET 8

### 如果您没有安装 Excel 或想要跨平台
→ **使用方案 B**（haris-musa/excel-mcp-server）
- 无需额外安装
- 已经可以工作
- 跨平台支持

---

## 快速修复脚本

### 安装 .NET 8 并配置 ExcelMcp

创建并运行以下 PowerShell 脚本：

```powershell
# 检查 .NET 是否已安装
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Host "正在安装 .NET 8 SDK..." -ForegroundColor Yellow
    winget install Microsoft.DotNet.SDK.8 --accept-package-agreements --accept-source-agreements
    
    Write-Host "请重启 PowerShell 后继续..." -ForegroundColor Yellow
    exit
}

# 检查 ExcelMcp 是否已安装
$tools = dotnet tool list --global
if (-not ($tools -match "ExcelMcp")) {
    Write-Host "正在安装 ExcelMcp MCP Server..." -ForegroundColor Yellow
    dotnet tool install --global Sbroenne.ExcelMcp.McpServer
}

Write-Host "安装完成！请重启 Cursor。" -ForegroundColor Green
```

---

## 验证配置

重启 Cursor 后，检查日志文件：
- 如果看到 "Found X tools"，说明配置成功
- 如果看到 "dotnet is not recognized"，说明需要安装 .NET 8
- 如果看到 "I/O operation on closed file" 但工具可用，可以忽略（haris-musa 版本的已知问题）

---

## 需要帮助？

如果问题仍然存在，请：
1. 检查日志文件中的具体错误信息
2. 确认所有必需组件已正确安装
3. 尝试重启 Cursor 和 PowerShell
4. 查看项目文档获取更多信息

