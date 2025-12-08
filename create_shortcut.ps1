# 创建自动化报表工具的快捷方式
$WScriptShell = New-Object -ComObject WScript.Shell

# 定义快捷方式信息
$ShortcutPath = "$env:USERPROFILE\Desktop\自动化报表工具.lnk"
$TargetPath = "$PSScriptRoot\启动报表工具.bat"
$WorkingDirectory = "$PSScriptRoot"
$IconLocation = "$env:SystemRoot\System32\shell32.dll,4"

# 创建快捷方式
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = $WorkingDirectory
$Shortcut.IconLocation = $IconLocation
$Shortcut.Description = "自动化报表工具 - 快速生成各类数据报表"
$Shortcut.Save()

# 提示用户
Write-Host "快捷方式已创建到桌面：自动化报表工具.lnk"
Write-Host "可以将此快捷方式复制到开始菜单或任务栏以方便使用"

# 可选：创建开始菜单快捷方式
$StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\自动化报表工具.lnk"
$Shortcut2 = $WScriptShell.CreateShortcut($StartMenuPath)
$Shortcut2.TargetPath = $TargetPath
$Shortcut2.WorkingDirectory = $WorkingDirectory
$Shortcut2.IconLocation = $IconLocation
$Shortcut2.Description = "自动化报表工具 - 快速生成各类数据报表"
$Shortcut2.Save()

Write-Host "快捷方式已创建到开始菜单"
