import pandas as pd
import sys
import os

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from auto_report import AutoReportEngine, ReportConfig, DataSourceConfig

# 创建单数据源配置
single_config = ReportConfig(
    report_name="单数据源测试",
    data_source_type="excel",
    data_source_path="test_data.xlsx",
    output_format=["excel", "html"]
)

# 创建多数据源配置
multi_config = ReportConfig(
    report_name="多数据源测试",
    output_format=["excel", "html"],
    data_sources=[
        DataSourceConfig(
            name="销售数据",
            type="excel",
            path="test_data.xlsx",
            parameters={"sheet_name": "Sheet1"}
        ),
        DataSourceConfig(
            name="成本数据",
            type="csv",
            path="data/cost_data.csv",
            parameters={"delimiter": ",", "parse_dates": ["日期"]}
        )
    ]
)

# 测试单数据源
print("=== 单数据源测试 ===")
try:
    engine = AutoReportEngine(single_config)
    generated_files = engine.run()
    print("单数据源报表生成成功！")
    for fmt, path in generated_files.items():
        print(f"{fmt}: {path}")
except Exception as e:
    print(f"单数据源测试失败: {e}")

# 测试多数据源（仅两个数据源，避免客户数据的复杂性）
print("\n=== 多数据源测试（销售+成本）===")
try:
    engine = AutoReportEngine(multi_config)
    generated_files = engine.run()
    print("多数据源报表生成成功！")
    for fmt, path in generated_files.items():
        print(f"{fmt}: {path}")
except Exception as e:
    print(f"多数据源测试失败: {e}")
    import traceback
    traceback.print_exc()