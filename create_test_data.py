import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# 设置随机种子以确保结果可重复
np.random.seed(42)

# 生成日期范围
start_date = datetime(2023, 1, 1)
dates = [start_date + timedelta(days=i) for i in range(30)]

# 生成随机销售数据
data = {
    '日期': dates,
    '金额': np.random.randint(1000, 10000, 30),  # 金额在1000到10000之间
    '类别': np.random.choice(['电子产品', '服装', '食品', '家居', '图书'], 30),  # 随机类别
    '地区': np.random.choice(['北京', '上海', '广州', '深圳', '杭州'], 30),  # 随机地区
    '数量': np.random.randint(1, 20, 30)  # 数量在1到20之间
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存到Excel文件
df.to_excel('test_data.xlsx', index=False, sheet_name='Sheet1')

print("测试数据已生成：test_data.xlsx")
print("数据结构：")
print(df.head())
print("\n列名：", df.columns.tolist())