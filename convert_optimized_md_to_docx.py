#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将优化后的Markdown使用说明书转换为Word文档
"""

import zipfile
import os
import shutil

# 读取Markdown文件内容
def read_markdown(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

# 将Markdown转换为Word XML格式
def markdown_to_word_xml(markdown_content):
    # 解析Markdown标题和段落
    lines = markdown_content.split('\n')
    word_xml = ['<w:body>']
    current_level = 0
    in_code_block = False
    code_language = ''
    code_content = ''
    in_list = False
    list_type = ''
    list_counter = 0
    
    for line in lines:
        line = line.strip()
        
        # 处理空行
        if not line:
            if in_code_block:
                code_content += '\n'
            else:
                word_xml.append('<w:p/>')
            continue
        
        # 处理代码块
        if line.startswith('```'):
            if not in_code_block:
                in_code_block = True
                code_language = line[3:].strip()
                code_content = ''
            else:
                in_code_block = False
                # 输出代码块
                word_xml.append('<w:p>')
                word_xml.append('<w:r>')
                word_xml.append('<w:t xml:space="preserve">' + code_content.replace('<', '&lt;').replace('>', '&gt;') + '</w:t>')
                word_xml.append('</w:r>')
                word_xml.append('</w:p>')
            continue
            
        if in_code_block:
            code_content += line + '\n'
            continue
        
        # 处理标题
        if line.startswith('#'):
            # 结束列表
            if in_list:
                word_xml.append('</w:tbl>')
                in_list = False
            
            level = line.count('#')
            text = line[level:].strip()
            word_xml.append('<w:p>')
            word_xml.append('<w:pPr>')
            word_xml.append('<w:jc w:val="center"/>')
            word_xml.append('</w:pPr>')
            word_xml.append('<w:r>')
            word_xml.append('<w:rPr>')
            word_xml.append('<w:b/>')
            word_xml.append('<w:sz w:val="' + str(24 - level * 2) + '"/>')
            word_xml.append('</w:rPr>')
            word_xml.append('<w:t xml:space="preserve">' + text + '</w:t>')
            word_xml.append('</w:r>')
            word_xml.append('</w:p>')
            continue
        
        # 处理列表
        if line.startswith('* '):
            if not in_list:
                in_list = True
                list_type = 'unordered'
                word_xml.append('<w:p>')
            else:
                word_xml.append('</w:p><w:p>')
            
            text = line[2:].strip()
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">• ' + text + '</w:t>')
            word_xml.append('</w:r>')
            continue
            
        if line.startswith('- '):
            if not in_list:
                in_list = True
                list_type = 'unordered'
                word_xml.append('<w:p>')
            else:
                word_xml.append('</w:p><w:p>')
            
            text = line[2:].strip()
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">• ' + text + '</w:t>')
            word_xml.append('</w:r>')
            continue
            
        if line.startswith('1. '):
            if not in_list:
                in_list = True
                list_type = 'ordered'
                list_counter = 1
                word_xml.append('<w:p>')
            else:
                word_xml.append('</w:p><w:p>')
                list_counter += 1
            
            text = line[3:].strip()
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">' + str(list_counter) + '. ' + text + '</w:t>')
            word_xml.append('</w:r>')
            continue
            
        # 处理表格
        if line.startswith('|') and '|' in line[1:]:
            # 简单处理表格，只保留文本
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            word_xml.append('<w:p>')
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">' + ' | '.join(cells) + '</w:t>')
            word_xml.append('</w:r>')
            word_xml.append('</w:p>')
            continue
            
        # 处理粗体
        if '**' in line:
            parts = line.split('**')
            word_xml.append('<w:p>')
            for i, part in enumerate(parts):
                if i % 2 == 0:
                    word_xml.append('<w:r><w:t xml:space="preserve">' + part + '</w:t></w:r>')
                else:
                    word_xml.append('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + part + '</w:t></w:r>')
            word_xml.append('</w:p>')
            continue
            
        # 处理普通段落
        if in_list:
            word_xml.append('</w:p><w:p>')
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">' + line + '</w:t>')
            word_xml.append('</w:r>')
        else:
            word_xml.append('<w:p>')
            word_xml.append('<w:r>')
            word_xml.append('<w:t xml:space="preserve">' + line + '</w:t>')
            word_xml.append('</w:r>')
            word_xml.append('</w:p>')
    
    # 关闭未结束的标签
    if in_code_block:
        word_xml.append('<w:p>')
        word_xml.append('<w:r>')
        word_xml.append('<w:t xml:space="preserve">' + code_content.replace('<', '&lt;').replace('>', '&gt;') + '</w:t>')
        word_xml.append('</w:r>')
        word_xml.append('</w:p>')
    
    if in_list:
        word_xml.append('</w:p>')
    
    word_xml.append('</w:body>')
    return ''.join(word_xml)

# 创建Word文档
def create_word_document(content, output_path):
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
    word_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps2="http://schemas.microsoft.com/office/word/2013/wordprocessingShape" xmlns:wpg2="http://schemas.microsoft.com/office/word/2013/wordprocessingGroup" xmlns:wpg14="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  {content}
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
  </w:sectPr>
</w:document>'''
    
    with open(os.path.join(temp_dir, "word", "document.xml"), "w", encoding="utf-8") as f:
        f.write(word_xml)
    
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
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = file_path.replace(temp_dir, "", 1).lstrip(os.sep)
                docx.write(file_path, arcname)
    
    # 清理临时文件
    shutil.rmtree(temp_dir)

# 主函数
def main():
    # 读取Markdown文件
    markdown_path = "c:\\Users\\25331\\Desktop\\新建文件夹\\自动化报表工具使用说明书(优化版).md"
    markdown_content = read_markdown(markdown_path)
    
    # 转换为Word XML
    word_xml_content = markdown_to_word_xml(markdown_content)
    
    # 创建Word文档
    output_path = "c:\\Users\\25331\\Desktop\\新建文件夹\\自动化报表工具使用说明书(优化版).docx"
    create_word_document(word_xml_content, output_path)
    
    print(f"优化版Word文档已生成：{output_path}")

if __name__ == "__main__":
    main()
