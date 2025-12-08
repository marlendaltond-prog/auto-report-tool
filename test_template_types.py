#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试多类型模板功能
"""

import os
import sys
import pandas as pd
from auto_report import HTMLReportGenerator, ReportConfig

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_template_types():
    """测试不同模板类型"""
    print("测试多类型模板功能...")
    
    # 创建测试数据
    data = {
        '日期': pd.date_range(start='2023-01-01', periods=12, freq='M'),
        '销售额': [15000, 22000, 18000, 25000, 28000, 32000, 35000, 38000, 42000, 45000, 50000, 55000],
        '利润': [3000, 4400, 3600, 5000, 5600, 6400, 7000, 7600, 8400, 9000, 10000, 11000],
        '产品类别': ['A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B']
    }
    df = pd.DataFrame(data)
    
    # 计算关键指标
    metrics = {
        'total_records': len(df),
        'total_columns': len(df.columns),
        '总销售额': df['销售额'].sum(),
        '平均销售额': df['销售额'].mean(),
        '总利润': df['利润'].sum(),
        '平均利润': df['利润'].mean()
    }
    
    # 定义图表配置
    charts = [
        {
            'type': 'bar',
            'title': '月度销售额',
            'x_field': '日期',
            'y_field': '销售额'
        },
        {
            'type': 'line',
            'title': '利润趋势',
            'x_field': '日期',
            'y_field': '利润'
        }
    ]
    
    # 测试不同模板类型
    template_types = ['default', 'simple', 'detailed', 'business']
    
    for template_type in template_types:
        print(f"\n测试模板类型: {template_type}")
        
        try:
            # 创建HTMLReportGenerator实例
            report_generator = HTMLReportGenerator(template_type=template_type)
            
            # 生成报表
            output_path = os.path.join('reports', f'test_report_{template_type}.html')
            
            # 创建输出目录
            os.makedirs('reports', exist_ok=True)
            
            # 生成报表
            report_generator.generate(
                df=df,
                metrics=metrics,
                report_name='月度销售报表',
                output_path=output_path,
                charts=charts
            )
            
            print(f"✓ 成功生成报表: {output_path}")
        except Exception as e:
            print(f"✗ 生成报表失败: {e}")
    
    print("\n多类型模板功能测试完成！")

if __name__ == "__main__":
    test_template_types()
