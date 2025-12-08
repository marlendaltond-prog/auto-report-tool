# Create shortcut for Auto Report Tool in Scripts directory
$WScriptShell = New-Object -ComObject WScript.Shell

# Define Scripts directory path
$ScriptsPath = "C:\Users\25331\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\LocalCache\local-packages\Python313\Scripts"

# Define shortcut information
$ShortcutPath = "$env:USERPROFILE\Desktop\Auto Report Tool.lnk"
$TargetPath = "$ScriptsPath\auto-report.exe"
$WorkingDirectory = "$ScriptsPath"
$IconLocation = "$env:SystemRoot\System32\shell32.dll,4"  # Use default icon

# Create shortcut
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = $WorkingDirectory
$Shortcut.IconLocation = $IconLocation
$Shortcut.Description = "Auto Report Tool - Generate reports quickly"
$Shortcut.Save()

# Notify user
Write-Host "Shortcut created on desktop: Auto Report Tool.lnk"
Write-Host "You can copy this shortcut to Start Menu or Taskbar for easy access"

# Optional: Create Start Menu shortcut
$StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Auto Report Tool.lnk"
$Shortcut2 = $WScriptShell.CreateShortcut($StartMenuPath)
$Shortcut2.TargetPath = $TargetPath
$Shortcut2.WorkingDirectory = $WorkingDirectory
$Shortcut2.IconLocation = $IconLocation
$Shortcut2.Description = "Auto Report Tool - Generate reports quickly"
$Shortcut2.Save()

Write-Host "Shortcut created in Start Menu"