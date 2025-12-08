# 自动化报表工具卸载脚本
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "            自动化报表工具卸载程序              " -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

# 获取脚本所在目录
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 确认卸载
$Confirm = Read-Host "确定要卸载自动化报表工具吗？(Y/N)"
if ($Confirm -notmatch "^[Yy]$") {
    Write-Host "卸载已取消。" -ForegroundColor Yellow
    pause
    exit 0
}

Write-Host ""
Write-Host "正在执行卸载操作..." -ForegroundColor Green

# 删除桌面快捷方式
Write-Host "删除桌面快捷方式..." -NoNewline
$DesktopShortcut = "$env:USERPROFILE\Desktop\AutoReportTool.lnk"
if (Test-Path $DesktopShortcut) {
    Remove-Item -Path $DesktopShortcut -Force
    Write-Host " √" -ForegroundColor Green
} else {
    Write-Host " √ (未找到)" -ForegroundColor Yellow
}

# 检查开始菜单快捷方式
Write-Host "检查开始菜单快捷方式..." -NoNewline
$StartMenuShortcut = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\AutoReportTool.lnk"
if (Test-Path $StartMenuShortcut) {
    Remove-Item -Path $StartMenuShortcut -Force
    Write-Host " √ (已删除)" -ForegroundColor Green
} else {
    Write-Host " √ (未找到)" -ForegroundColor Yellow
}

# 检查Python依赖
Write-Host "检查Python依赖包..."
$Dependencies = @("pandas", "openpyxl", "sqlalchemy", "jinja2", "reportlab", "requests")
$InstalledDeps = @()

foreach ($Dep in $Dependencies) {
    try {
        python -c "import $Dep" | Out-Null
        $InstalledDeps += $Dep
    } catch {
        # 依赖未安装
    }
}

if ($InstalledDeps.Count -gt 0) {
    Write-Host "已安装的依赖包: $($InstalledDeps -join ", ")" -ForegroundColor Cyan
    $UninstallPython = Read-Host "是否要卸载这些依赖包？(Y/N)"
    if ($UninstallPython -match "^[Yy]$") {
        Write-Host "正在卸载Python依赖包..."
        pip uninstall -y $InstalledDeps | Out-Null
        Write-Host "依赖包卸载完成！" -ForegroundColor Green
    } else {
        Write-Host "保留Python依赖包。" -ForegroundColor Yellow
    }
}

# 检查生成的报告文件
$ReportsDir = "$ScriptDir\reports"
if (Test-Path $ReportsDir) {
    Write-Host ""
    $ReportFiles = Get-ChildItem -Path $ReportsDir -Recurse | Measure-Object | Select-Object -ExpandProperty Count
    Write-Host "发现 $ReportFiles 个报告文件保存在: $ReportsDir" -ForegroundColor Cyan
    $DeleteReports = Read-Host "是否要删除这些报告文件？(Y/N)"
    if ($DeleteReports -match "^[Yy]$") {
        Write-Host "正在删除报告文件..." -NoNewline
        Remove-Item -Path $ReportsDir -Recurse -Force
        Write-Host " √" -ForegroundColor Green
    } else {
        Write-Host "保留报告文件。" -ForegroundColor Yellow
    }
}

# 检查虚拟环境
$VenvDir = "$ScriptDir\.venv"
if (Test-Path $VenvDir) {
    Write-Host ""
    Write-Host "发现虚拟环境: $VenvDir" -ForegroundColor Cyan
    $DeleteVenv = Read-Host "是否要删除虚拟环境？(Y/N)"
    if ($DeleteVenv -match "^[Yy]$") {
        Write-Host "正在删除虚拟环境..." -NoNewline
        Remove-Item -Path $VenvDir -Recurse -Force
        Write-Host " √" -ForegroundColor Green
    } else {
        Write-Host "保留虚拟环境。" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "            卸载完成！                          " -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "您可以手动删除以下文件夹来完全移除程序：" -ForegroundColor Yellow
Write-Host $ScriptDir -ForegroundColor White
Write-Host ""
Write-Host "感谢您使用自动化报表工具！" -ForegroundColor Green
Write-Host ""
pause