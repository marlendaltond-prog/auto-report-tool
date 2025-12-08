#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试所有模板类型的生成功能
"""

import sys
import os
import pandas as pd
from datetime import datetime
from auto_report import HTMLReportGenerator

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_html_report_generator():
    """测试HTMLReportGenerator的所有功能"""
    print("=== 测试HTMLReportGenerator所有功能 ===")
    
    # 创建测试数据
    data = {
        '产品名称': ['产品A', '产品B', '产品C', '产品D', '产品E'],
        '销售数量': [100, 200, 150, 300, 250],
        '销售金额': [10000, 20000, 15000, 30000, 25000],
        '销售区域': ['华东', '华南', '华北', '西南', '西北'],
        '销售日期': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04', '2024-01-05']
    }
    df = pd.DataFrame(data)
    
    # 准备测试数据
    metrics = {
        'total_records': len(df),
        'total_columns': len(df.columns),
        'avg_sales': df['销售金额'].mean(),
        'total_sales': df['销售金额'].sum()
    }
    
    charts = [
        {
            'title': '产品销售金额分布',
            'type': '柱状图',
            'data_range': '所有产品'
        },
        {
            'title': '销售区域分布',
            'type': '饼图',
            'data_range': '所有区域'
        }
    ]
    
    report_name = '销售数据分析报表'
    
    # 测试所有模板类型
    template_types = ['default', 'simple', 'detailed', 'business']
    
    for template_type in template_types:
        print(f"\n--- 测试模板类型: {template_type} ---")
        try:
            # 创建报表生成器
            html_generator = HTMLReportGenerator(template_type=template_type)
            
            # 生成报表
            output_path = f"reports/test_{template_type}_report.html"
            result = html_generator.generate(df, metrics, output_path, charts, report_name)
            
            # 验证生成结果
            if os.path.exists(result) and os.path.getsize(result) > 0:
                print(f"✓ 成功生成报表: {result}")
                print(f"  文件大小: {os.path.getsize(result)} 字节")
            else:
                print(f"✗ 报表生成失败: {result}")
                print(f"  文件不存在或为空")
                
        except Exception as e:
            print(f"✗ 测试失败: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n=== 所有模板类型测试完成 ===")

def test_large_dataset():
    """测试大数据集的处理能力"""
    print("\n=== 测试大数据集处理能力 ===")
    
    # 创建大数据集（10000行）
    data = {
        '产品名称': [f'产品{i}' for i in range(10000)],
        '销售数量': [i * 2 for i in range(10000)],
        '销售金额': [i * 100 for i in range(10000)],
        '销售区域': ['华东', '华南', '华北', '西南', '西北'] * 2000,
        '销售日期': [datetime.now().strftime('%Y-%m-%d')] * 10000
    }
    df = pd.DataFrame(data)
    
    metrics = {
        'total_records': len(df),
        'total_columns': len(df.columns),
        'avg_sales': df['销售金额'].mean(),
        'total_sales': df['销售金额'].sum()
    }
    
    try:
        html_generator = HTMLReportGenerator(template_type='detailed')
        output_path = "reports/test_large_dataset.html"
        result = html_generator.generate(df, metrics, output_path)
        
        if os.path.exists(result) and os.path.getsize(result) > 0:
            print(f"✓ 大数据集报表生成成功: {result}")
            print(f"  文件大小: {os.path.getsize(result)} 字节")
        else:
            print(f"✗ 大数据集报表生成失败")
            
    except Exception as e:
        print(f"✗ 大数据集测试失败: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n=== 大数据集测试完成 ===")

def main():
    """主测试函数"""
    # 创建输出目录
    if not os.path.exists('reports'):
        os.makedirs('reports')
    
    # 运行所有测试
    test_html_report_generator()
    test_large_dataset()
    
    print("\n=== 所有测试完成 ===")
    print("您可以在 'reports' 目录下查看生成的报表文件")

if __name__ == "__main__":
    main()
