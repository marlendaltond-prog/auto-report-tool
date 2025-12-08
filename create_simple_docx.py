#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建简化版的Word使用说明书
"""

# 这里我们使用直接写入XML的方式创建Word文档
# 因为python-docx可能安装有问题

word_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps2="http://schemas.microsoft.com/office/word/2013/wordprocessingShape" xmlns:wpg2="http://schemas.microsoft.com/office/word/2013/wordprocessingGroup" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg14="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:t xml:space="preserve">自动化报表工具 - AutoReport Pro 使用说明书</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:spacing w:after="240"/>
      </w:pPr>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>1. 工具介绍</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>AutoReport Pro 是一款功能强大的自动化报表生成工具，支持多种数据源、多种输出格式和灵活的数据处理能力。</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>主要功能特点：</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 支持多种数据源：Excel、CSV、SQL数据库、API</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 支持多种输出格式：Excel、PDF、HTML、邮件</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 可配置的数据处理：过滤、计算、图表生成</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 可配置的报表样式和模板</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 支持定时执行和邮件发送</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 支持报表自动发送到指定邮箱</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>2. 安装要求</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>系统要求：Windows/macOS/Linux，Python 3.7 或更高版本</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>依赖包安装：</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>pip install pandas openpyxl sqlalchemy jinja2 reportlab requests</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>3. 使用流程</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>步骤1：准备数据源</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 确保数据源文件（Excel/CSV）格式正确</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 检查数据完整性和格式一致性</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>步骤2：配置报表参数</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 选择输出格式（Excel/PDF/HTML/邮件）</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 设置输出目录</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>步骤3：运行工具</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 使用命令行参数直接运行</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 或使用配置文件运行</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>步骤4：查看和使用报表</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 在输出目录查看生成的报表文件</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>4. 命令行参数说明</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --help：显示基本帮助信息</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --help-all：显示详细帮助信息</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --config：配置文件路径</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --data：数据源路径</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --output：输出目录</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --format：输出格式，多个用逗号分隔</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>5. 实际使用示例</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>使用Excel数据源生成Excel和PDF格式报表：</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>python auto_report.py --data "302594156_按序号_大学生对新能源汽车购买意向调查研究_254_246.xlsx" --output reports --format excel,pdf</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>参数说明：</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --data：指定数据源文件路径</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --output：设置输出目录</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• --format：指定输出格式</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>6. 注意事项</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 确保数据源文件路径正确</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 输出目录如果不存在，工具会自动创建</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>• 对于大型数据集，可能需要较长时间生成报表</w:t>
      </w:r>
    </w:p>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>'''

# 创建一个简化的Word文档
# 这里我们创建一个基本的.docx文件结构
import zipfile
import os

# 定义.docx文件路径
docx_path = "c:\\Users\\25331\\Desktop\\新建文件夹\\自动化报表工具使用说明书.docx"

# 创建临时目录用于构建.docx文件结构
temp_dir = "temp_docx"
os.makedirs(os.path.join(temp_dir, "word"), exist_ok=True)
os.makedirs(os.path.join(temp_dir, "_rels"), exist_ok=True)

# 创建[Content_Types].xml文件
content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

with open(os.path.join(temp_dir, "[Content_Types].xml"), "w", encoding="utf-8") as f:
    f.write(content_types)

# 创建_rels/.rels文件
rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

with open(os.path.join(temp_dir, "_rels", ".rels"), "w", encoding="utf-8") as f:
    f.write(rels)

# 创建word/document.xml文件
with open(os.path.join(temp_dir, "word", "document.xml"), "w", encoding="utf-8") as f:
    f.write(word_xml)

# 创建word/_rels/document.xml.rels文件
word_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>'''

os.makedirs(os.path.join(temp_dir, "word", "_rels"), exist_ok=True)
with open(os.path.join(temp_dir, "word", "_rels", "document.xml.rels"), "w", encoding="utf-8") as f:
    f.write(word_rels)

# 创建word/settings.xml文件
settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps2="http://schemas.microsoft.com/office/word/2013/wordprocessingShape" xmlns:wpg2="http://schemas.microsoft.com/office/word/2013/wordprocessingGroup" xmlns:wpg14="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
</w:settings>'''

with open(os.path.join(temp_dir, "word", "settings.xml"), "w", encoding="utf-8") as f:
    f.write(settings)

# 创建word/styles.xml文件
styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps2="http://schemas.microsoft.com/office/word/2013/wordprocessingShape" xmlns:wpg2="http://schemas.microsoft.com/office/word/2013/wordprocessingGroup" xmlns:wpg14="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:lang w:val="zh-CN"/>
    </w:rPr>
  </w:style>
</w:styles>'''

with open(os.path.join(temp_dir, "word", "styles.xml"), "w", encoding="utf-8") as f:
    f.write(styles)

# 创建word/theme/theme1.xml文件
os.makedirs(os.path.join(temp_dir, "word", "theme"), exist_ok=True)
theme = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="333333"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>'''

with open(os.path.join(temp_dir, "word", "theme", "theme1.xml"), "w", encoding="utf-8") as f:
    f.write(theme)

# 打包成.docx文件
with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = file_path.replace(temp_dir, "", 1).lstrip(os.sep)
            docx.write(file_path, arcname)

# 清理临时文件
import shutil
shutil.rmtree(temp_dir)

print(f"Word文档已生成：{docx_path}")
