import pandas as pd
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_data_merge():
    # 创建测试数据
    df1 = pd.DataFrame({
        '日期': pd.date_range('2023-01-01', periods=3),
        '金额': [100, 200, 300],
        '类别': ['A', 'B', 'A']
    })
    
    df2 = pd.DataFrame({
        '日期': ['2023-01-01', '2023-01-02', '2023-01-03'],
        '成本': [50, 100, 150],
        '部门': ['销售部', '市场部', '销售部']
    })
    
    df3 = pd.DataFrame({
        '类别': ['A', 'B'],
        '客户名称': ['张三', '李四'],
        '地区': ['华东', '华南']
    })
    
    print("原始数据:")
    print("df1:")
    print(df1)
    print("\ndf2:")
    print(df2)
    print("\ndf3:")
    print(df3)
    
    print("\n数据类型:")
    print("df1['日期']:", df1['日期'].dtype)
    print("df2['日期']:", df2['日期'].dtype)
    
    # 转换df2的日期列类型
    df2['日期'] = pd.to_datetime(df2['日期'])
    print("\n转换后的数据类型:")
    print("df2['日期']:", df2['日期'].dtype)
    
    # 测试合并
    print("\n开始合并数据:")
    
    # 合并df1和df2
    result1 = pd.merge(df1, df2, on='日期', how='inner')
    print("合并df1和df2:")
    print(result1)
    
    # 合并df3
    result2 = pd.merge(result1, df3, on='类别', how='left')
    print("\n合并df3:")
    print(result2)
    
    return result2

if __name__ == "__main__":
    test_data_merge()