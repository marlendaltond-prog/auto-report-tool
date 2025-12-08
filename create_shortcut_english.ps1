# Create shortcut for Auto Report Tool
$WScriptShell = New-Object -ComObject WScript.Shell

# Define shortcut information
$ShortcutPath = "$env:USERPROFILE\Desktop\AutoReportTool.lnk"
$TargetPath = "$PSScriptRoot\启动报表工具.bat"
$WorkingDirectory = "$PSScriptRoot"
$IconLocation = "$env:SystemRoot\System32\shell32.dll,4"

# Create shortcut
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = $WorkingDirectory
$Shortcut.IconLocation = $IconLocation
$Shortcut.Description = "Auto Report Tool - Generate reports quickly"
$Shortcut.Save()

Write-Host "Shortcut created on desktop: AutoReportTool.lnk"
