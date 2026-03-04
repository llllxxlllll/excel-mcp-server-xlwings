# ExcelMcp MCP Server 安装验证脚本
# 此脚本检查所有必需的组件是否已正确安装

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ExcelMcp MCP Server 安装验证" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$allChecksPassed = $true

# 检查 1: .NET 8
Write-Host "[1/4] 检查 .NET 8..." -ForegroundColor Yellow
try {
    $dotnetVersion = dotnet --version 2>&1
    if ($LASTEXITCODE -eq 0 -and $dotnetVersion -match "^8\.") {
        Write-Host "  ✓ .NET 已安装: $dotnetVersion" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq 0) {
        Write-Host "  ✗ .NET 版本过低: $dotnetVersion (需要 8.0+)" -ForegroundColor Red
        $allChecksPassed = $false
    } else {
        Write-Host "  ✗ .NET 未安装或不在 PATH 中" -ForegroundColor Red
        Write-Host "    安装命令: winget install Microsoft.DotNet.SDK.8" -ForegroundColor Yellow
        $allChecksPassed = $false
    }
} catch {
    Write-Host "  ✗ .NET 未安装" -ForegroundColor Red
    Write-Host "    安装命令: winget install Microsoft.DotNet.SDK.8" -ForegroundColor Yellow
    $allChecksPassed = $false
}
Write-Host ""

# 检查 2: ExcelMcp MCP Server
Write-Host "[2/4] 检查 ExcelMcp MCP Server..." -ForegroundColor Yellow
try {
    $tools = dotnet tool list --global 2>&1
    if ($tools -match "ExcelMcp") {
        Write-Host "  ✓ ExcelMcp MCP Server 已安装" -ForegroundColor Green
    } else {
        Write-Host "  ✗ ExcelMcp MCP Server 未安装" -ForegroundColor Red
        Write-Host "    安装命令: dotnet tool install --global Sbroenne.ExcelMcp.McpServer" -ForegroundColor Yellow
        $allChecksPassed = $false
    }
} catch {
    Write-Host "  ✗ 无法检查 ExcelMcp 安装状态" -ForegroundColor Red
    $allChecksPassed = $false
}
Write-Host ""

# 检查 3: Microsoft Excel
Write-Host "[3/4] 检查 Microsoft Excel..." -ForegroundColor Yellow
$excelFound = $false

# 方法 1: 检查注册表
try {
    $excelPath = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
    if ($excelPath) {
        Write-Host "  ✓ Excel 已安装: $($excelPath.'(default)')" -ForegroundColor Green
        $excelFound = $true
    }
} catch {
    # 继续其他检查方法
}

# 方法 2: 检查常见安装路径
if (-not $excelFound) {
    $commonPaths = @(
        "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
        "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            Write-Host "  ✓ Excel 已安装: $path" -ForegroundColor Green
            $excelFound = $true
            break
        }
    }
}

# 方法 3: 尝试运行 Excel
if (-not $excelFound) {
    try {
        $excelProcess = Get-Process excel -ErrorAction SilentlyContinue
        if ($excelProcess) {
            Write-Host "  ✓ Excel 正在运行" -ForegroundColor Green
            $excelFound = $true
        }
    } catch {
        # Excel 未运行，但可能已安装
    }
}

if (-not $excelFound) {
    Write-Host "  ✗ Excel 未找到" -ForegroundColor Red
    Write-Host "    请安装 Microsoft Excel 2016 或更高版本" -ForegroundColor Yellow
    $allChecksPassed = $false
}
Write-Host ""

# 检查 4: Cursor MCP 配置
Write-Host "[4/4] 检查 Cursor MCP 配置..." -ForegroundColor Yellow
$mcpConfigPath = "$env:USERPROFILE\.cursor\mcp.json"
if (Test-Path $mcpConfigPath) {
    try {
        $config = Get-Content $mcpConfigPath -Raw | ConvertFrom-Json
        if ($config.mcpServers.excel) {
            $excelConfig = $config.mcpServers.excel
            if ($excelConfig.command -eq "dotnet" -and $excelConfig.args -contains "mcp-excel") {
                Write-Host "  ✓ Cursor MCP 配置正确" -ForegroundColor Green
                Write-Host "    命令: $($excelConfig.command)" -ForegroundColor Gray
                Write-Host "    参数: $($excelConfig.args -join ' ')" -ForegroundColor Gray
            } else {
                Write-Host "  ⚠ Cursor MCP 配置存在，但可能不正确" -ForegroundColor Yellow
                Write-Host "    当前配置: command=$($excelConfig.command), args=$($excelConfig.args -join ',')" -ForegroundColor Gray
            }
        } else {
            Write-Host "  ✗ Cursor MCP 配置中未找到 excel 服务器" -ForegroundColor Red
            $allChecksPassed = $false
        }
    } catch {
        Write-Host "  ✗ 无法读取 Cursor MCP 配置文件" -ForegroundColor Red
        $allChecksPassed = $false
    }
} else {
    Write-Host "  ✗ Cursor MCP 配置文件不存在: $mcpConfigPath" -ForegroundColor Red
    $allChecksPassed = $false
}
Write-Host ""

# 总结
Write-Host "========================================" -ForegroundColor Cyan
if ($allChecksPassed) {
    Write-Host "✓ 所有检查通过！" -ForegroundColor Green
    Write-Host ""
    Write-Host "下一步:" -ForegroundColor Cyan
    Write-Host "1. 重启 Cursor" -ForegroundColor White
    Write-Host "2. 测试 Excel 功能，例如：创建一个名为 'test.xlsx' 的空 Excel 文件" -ForegroundColor White
} else {
    Write-Host "✗ 部分检查未通过，请根据上述提示修复问题" -ForegroundColor Red
    Write-Host ""
    Write-Host "详细安装说明请查看: excel-mcp-server/EXCELMCP_INSTALLATION.md" -ForegroundColor Yellow
}
Write-Host "========================================" -ForegroundColor Cyan

