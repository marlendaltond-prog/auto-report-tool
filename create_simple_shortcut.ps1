# 创建简单的快捷方式

$scriptPath = "C:\Users\25331\Desktop\新建文件夹\启动报表工具.bat"
$shortcutPath = "C:\Users\25331\Desktop\AutoReport.lnk"

# 检查启动脚本是否存在
if (-not (Test-Path $scriptPath)) {
    Write-Host "Error: Startup script not found!"
    exit 1
}

# 创建快捷方式
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $scriptPath
$shortcut.WorkingDirectory = "C:\Users\25331\Desktop\新建文件夹"
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"