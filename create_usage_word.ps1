# 创建自动化报表工具使用说明书的PowerShell脚本

# 创建Word对象
$Word = New-Object -ComObject Word.Application
$Word.Visible = $false

# 创建新文档
$Document = $Word.Documents.Add()

# 添加标题
$TitleRange = $Document.Content
$TitleRange.Text = "自动化报表工具 - AutoReport Pro 使用说明书"
$TitleRange.Font.Size = 24
$TitleRange.Font.Name = "微软雅黑"
$TitleRange.Font.Bold = $true
$TitleRange.ParagraphFormat.Alignment = 1  # 居中对齐

# 插入分页符
$Document.Content.InsertParagraphAfter()
$Document.Content.InsertBreak(7)  # wdPageBreak

# 1. 工具介绍
$IntroRange = $Document.Content
$IntroRange.Collapse(0)  # 移动到文档末尾
$IntroRange.Text = "1. 工具介绍"
$IntroRange.Font.Size = 18
$IntroRange.Font.Name = "微软雅黑"
$IntroRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$IntroTextRange = $Document.Content
$IntroTextRange.Collapse(0)
$IntroTextRange.Text = "AutoReport Pro 是一款功能强大的自动化报表生成工具，支持多种数据源、多种输出格式和灵活的数据处理能力。"
$IntroTextRange.Font.Size = 12
$IntroTextRange.Font.Name = "宋体"
$Document.Content.InsertParagraphAfter()

$FeaturesRange = $Document.Content
$FeaturesRange.Collapse(0)
$FeaturesRange.Text = "主要功能特点："
$FeaturesRange.Font.Size = 12
$FeaturesRange.Font.Name = "宋体"
$FeaturesRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$FeatureList = @(
    "• 📊 支持多种数据源：Excel、CSV、SQL数据库、API",
    "• 📄 支持多种输出格式：Excel、PDF、HTML、邮件",
    "• 🔧 可配置的数据处理：过滤、计算、图表生成",
    "• 🎨 可配置的报表样式和模板",
    "• ⏰ 支持定时执行和邮件发送",
    "• 📧 支持报表自动发送到指定邮箱"
)

foreach ($Feature in $FeatureList) {
    $FeatureRange = $Document.Content
    $FeatureRange.Collapse(0)
    $FeatureRange.Text = $Feature
    $FeatureRange.Font.Size = 12
    $FeatureRange.Font.Name = "宋体"
    $Document.Content.InsertParagraphAfter()
}

# 2. 安装要求
$InstallRange = $Document.Content
$InstallRange.Collapse(0)
$InstallRange.Text = "2. 安装要求"
$InstallRange.Font.Size = 18
$InstallRange.Font.Name = "微软雅黑"
$InstallRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$SysReqRange = $Document.Content
$SysReqRange.Collapse(0)
$SysReqRange.Text = "系统要求："
$SysReqRange.Font.Size = 12
$SysReqRange.Font.Name = "宋体"
$SysReqRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$SysReqList = @(
    "• Windows/macOS/Linux",
    "• Python 3.7 或更高版本"
)

foreach ($Req in $SysReqList) {
    $ReqRange = $Document.Content
    $ReqRange.Collapse(0)
    $ReqRange.Text = $Req
    $ReqRange.Font.Size = 12
    $ReqRange.Font.Name = "宋体"
    $Document.Content.InsertParagraphAfter()
}

$DepReqRange = $Document.Content
$DepReqRange.Collapse(0)
$DepReqRange.Text = "依赖包安装："
$DepReqRange.Font.Size = 12
$DepReqRange.Font.Name = "宋体"
$DepReqRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$DepTextRange = $Document.Content
$DepTextRange.Collapse(0)
$DepTextRange.Text = "工具需要以下第三方依赖包："
$DepTextRange.Font.Size = 12
$DepTextRange.Font.Name = "宋体"
$Document.Content.InsertParagraphAfter()

$CodeRange = $Document.Content
$CodeRange.Collapse(0)
$CodeRange.Text = "pip install pandas openpyxl sqlalchemy jinja2 reportlab requests"
$CodeRange.Font.Size = 10
$CodeRange.Font.Name = "Consolas"
$Document.Content.InsertParagraphAfter()

$OrTextRange = $Document.Content
$OrTextRange.Collapse(0)
$OrTextRange.Text = "或者使用提供的 requirements.txt 文件："
$OrTextRange.Font.Size = 12
$OrTextRange.Font.Name = "宋体"
$Document.Content.InsertParagraphAfter()

$Code2Range = $Document.Content
$Code2Range.Collapse(0)
$Code2Range.Text = "pip install -r requirements.txt"
$Code2Range.Font.Size = 10
$Code2Range.Font.Name = "Consolas"
$Document.Content.InsertParagraphAfter()

# 3. 快速开始
$QuickStartRange = $Document.Content
$QuickStartRange.Collapse(0)
$QuickStartRange.Text = "3. 快速开始"
$QuickStartRange.Font.Size = 18
$QuickStartRange.Font.Name = "微软雅黑"
$QuickStartRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$Example1Range = $Document.Content
$Example1Range.Collapse(0)
$Example1Range.Text = "1. 使用命令行参数生成报表："
$Example1Range.Font.Size = 12
$Example1Range.Font.Name = "宋体"
$Example1Range.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$Code3Range = $Document.Content
$Code3Range.Collapse(0)
$Code3Range.Text = "python auto_report.py --data data.xlsx --output reports --format excel,pdf"
$Code3Range.Font.Size = 10
$Code3Range.Font.Name = "Consolas"
$Document.Content.InsertParagraphAfter()

$Example2Range = $Document.Content
$Example2Range.Collapse(0)
$Example2Range.Text = "2. 使用配置文件生成报表："
$Example2Range.Font.Size = 12
$Example2Range.Font.Name = "宋体"
$Example2Range.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$Code4Range = $Document.Content
$Code4Range.Collapse(0)
$Code4Range.Text = "python auto_report.py --config report_config.json"
$Code4Range.Font.Size = 10
$Code4Range.Font.Name = "Consolas"
$Document.Content.InsertParagraphAfter()

# 4. 使用流程
$FlowRange = $Document.Content
$FlowRange.Collapse(0)
$FlowRange.Text = "4. 使用流程"
$FlowRange.Font.Size = 18
$FlowRange.Font.Name = "微软雅黑"
$FlowRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$FlowSteps = @(
    "1. 准备数据源",
    "   • 确保数据源文件（Excel/CSV）格式正确",
    "   • 检查数据完整性和格式一致性",
    "   • 如果使用SQL或API数据源，确保连接信息正确",
    "2. 配置报表参数",
    "   • 选择输出格式（Excel/PDF/HTML/邮件）",
    "   • 设置输出目录",
    "   • 配置数据处理规则（可选）",
    "   • 配置报表样式和模板（可选）",
    "3. 运行工具",
    "   • 使用命令行参数直接运行",
    "   • 或使用配置文件运行",
    "   • 检查运行日志和错误提示",
    "4. 查看和使用报表",
    "   • 在输出目录查看生成的报表文件",
    "   • 如果配置了邮件发送，检查收件邮箱",
    "   • 验证报表数据准确性",
    "5. 高级配置（可选）",
    "   • 配置定时执行",
    "   • 设置自定义数据处理逻辑",
    "   • 使用自定义报表模板"
)

foreach ($Step in $FlowSteps) {
    $StepRange = $Document.Content
    $StepRange.Collapse(0)
    $StepRange.Text = $Step
    $StepRange.Font.Size = 12
    $StepRange.Font.Name = "宋体"
    if ($Step -match "^\d+") {
        $StepRange.Font.Bold = $true
    }
    $Document.Content.InsertParagraphAfter()
}

# 5. 命令行参数说明
$ParamsRange = $Document.Content
$ParamsRange.Collapse(0)
$ParamsRange.Text = "5. 命令行参数说明"
$ParamsRange.Font.Size = 18
$ParamsRange.Font.Name = "微软雅黑"
$ParamsRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

# 6. 实际使用示例
$ExampleRange = $Document.Content
$ExampleRange.Collapse(0)
$ExampleRange.Text = "6. 实际使用示例"
$ExampleRange.Font.Size = 18
$ExampleRange.Font.Name = "微软雅黑"
$ExampleRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$RealExampleRange = $Document.Content
$RealExampleRange.Collapse(0)
$RealExampleRange.Text = "示例：使用Excel数据源生成Excel和PDF格式报表"
$RealExampleRange.Font.Size = 12
$RealExampleRange.Font.Name = "宋体"
$RealExampleRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$RealCodeRange = $Document.Content
$RealCodeRange.Collapse(0)
$RealCodeRange.Text = 'python auto_report.py --data "302594156_按序号_大学生对新能源汽车购买意向调查研究_254_246.xlsx" --output reports --format excel,pdf'
$RealCodeRange.Font.Size = 10
$RealCodeRange.Font.Name = "Consolas"
$Document.Content.InsertParagraphAfter()

$ExplanationRange = $Document.Content
$ExplanationRange.Collapse(0)
$ExplanationRange.Text = "参数说明："
$ExplanationRange.Font.Size = 12
$ExplanationRange.Font.Name = "宋体"
$ExplanationRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$ExplanationList = @(
    "   • --data：指定数据源文件路径，这里使用了完整的文件名",
    "   • --output：设置输出目录为当前目录下的reports文件夹",
    "   • --format：指定输出格式为Excel和PDF，用逗号分隔"
)

foreach ($Item in $ExplanationList) {
    $ItemRange = $Document.Content
    $ItemRange.Collapse(0)
    $ItemRange.Text = $Item
    $ItemRange.Font.Size = 12
    $ItemRange.Font.Name = "宋体"
    $Document.Content.InsertParagraphAfter()
}

# 7. 注意事项
$NotesRange = $Document.Content
$NotesRange.Collapse(0)
$NotesRange.Text = "7. 注意事项"
$NotesRange.Font.Size = 18
$NotesRange.Font.Name = "微软雅黑"
$NotesRange.Font.Bold = $true
$Document.Content.InsertParagraphAfter()

$NotesList = @(
    "• 确保数据源文件路径正确，文件名包含空格时需要用引号括起来",
    "• 输出目录如果不存在，工具会自动创建",
    "• 确保有足够的磁盘空间存储生成的报表文件",
    "• 对于大型数据集，可能需要较长时间生成报表",
    "• 使用API数据源时，确保网络连接正常且有访问权限"
)

foreach ($Note in $NotesList) {
    $NoteRange = $Document.Content
    $NoteRange.Collapse(0)
    $NoteRange.Text = $Note
    $NoteRange.Font.Size = 12
    $NoteRange.Font.Name = "宋体"
    $Document.Content.InsertParagraphAfter()
}

# 保存文档
$SavePath = "$PSScriptRoot\自动化报表工具使用说明书.docx"
$Document.SaveAs([ref]$SavePath)

# 关闭文档和Word
$Document.Close()
$Word.Quit()

# 释放COM对象
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Word文档已成功创建：$SavePath" -ForegroundColor Green