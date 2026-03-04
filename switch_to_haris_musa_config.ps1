# 切换到 haris-musa/excel-mcp-server 配置
# 此配置不需要 .NET 和 Excel，已经可以工作（25 个工具可用）

$mcpConfigPath = "$env:USERPROFILE\.cursor\mcp.json"

Write-Host "正在切换到 haris-musa/excel-mcp-server 配置..." -ForegroundColor Yellow

if (Test-Path $mcpConfigPath) {
    try {
        $config = Get-Content $mcpConfigPath -Raw | ConvertFrom-Json
        
        # 更新 excel 配置
        $config.mcpServers.excel = @{
            command = "uvx"
            args = @("excel-mcp-server", "stdio")
            env = @{
                PYTHONUTF8 = "1"
            }
        }
        
        # 保存配置
        $config | ConvertTo-Json -Depth 10 | Set-Content $mcpConfigPath -Encoding UTF8
        
        Write-Host "✓ 配置已更新为 haris-musa/excel-mcp-server" -ForegroundColor Green
        Write-Host ""
        Write-Host "下一步:" -ForegroundColor Cyan
        Write-Host "1. 重启 Cursor" -ForegroundColor White
        Write-Host "2. 服务器应该可以正常工作（25 个工具可用）" -ForegroundColor White
        Write-Host ""
        Write-Host "注意: 虽然日志中可能有 I/O 错误，但不影响功能使用" -ForegroundColor Yellow
    } catch {
        Write-Host "✗ 更新配置失败: $_" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "✗ 配置文件不存在: $mcpConfigPath" -ForegroundColor Red
    exit 1
}

