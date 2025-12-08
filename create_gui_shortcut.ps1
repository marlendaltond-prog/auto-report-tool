# 自动化报表工具图形界面快捷方式创建脚本

# 定义路径
$scriptPath = "$PSScriptRoot\启动报表工具.bat"
$desktopPath = [Environment]::GetFolderPath('Desktop')

# 检查启动脚本是否存在
if (-not (Test-Path $scriptPath)) {
    Write-Host "错误：启动脚本不存在！" -ForegroundColor Red
    Read-Host "按任意键退出..."
    exit 1
}

# 创建快捷方式
$WScriptShell = New-Object -ComObject WScript.Shell

# 桌面快捷方式
$desktopShortcut = $WScriptShell.CreateShortcut("$desktopPath\自动化报表工具.lnk")
$desktopShortcut.TargetPath = $scriptPath
$desktopShortcut.WorkingDirectory = $PSScriptRoot
$desktopShortcut.Save()
Write-Host "已在桌面创建快捷方式: 自动化报表工具.lnk" -ForegroundColor Green

# 开始菜单快捷方式
$startMenuPath = [Environment]::GetFolderPath('StartMenu')
$startMenuShortcut = $WScriptShell.CreateShortcut("$startMenuPath\自动化报表工具.lnk")
$startMenuShortcut.TargetPath = $scriptPath
$startMenuShortcut.WorkingDirectory = $PSScriptRoot
$startMenuShortcut.Save()
Write-Host "已在开始菜单创建快捷方式: 自动化报表工具.lnk" -ForegroundColor Green

Write-Host "快捷方式创建完成！" -ForegroundColor Green