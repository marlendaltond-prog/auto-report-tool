import sqlite3

# 创建SQLite数据库连接
conn = sqlite3.connect('data/customer.db')
cursor = conn.cursor()

# 创建customers表
cursor.execute('''
CREATE TABLE IF NOT EXISTS customers (
    customer_id INTEGER PRIMARY KEY,
    customer_name TEXT,
    region TEXT
)
''')

# 插入示例数据
customers = [
    (1, '张三', '华东'),
    (2, '李四', '华南'),
    (3, '王五', '华北'),
    (4, '赵六', '华东'),
    (5, '钱七', '西南')
]

cursor.executemany('INSERT INTO customers VALUES (?, ?, ?)', customers)

# 提交并关闭连接
conn.commit()
conn.close()

print("客户数据库创建完成！")