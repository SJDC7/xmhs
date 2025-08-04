import os
import sqlite3
from datetime import datetime

# 定义数据库文件路径
# 获取当前脚本的绝对目录路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_PATH = os.path.join(BASE_DIR,'data', 'employee_data.db')


def initialize_database():
    # 连接到SQLite数据库（如果没有数据库文件将自动创建）
    with sqlite3.connect(DATABASE_PATH) as conn:
        cursor = conn.cursor()

        # 创建表结构（如果表不存在）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS EmployeeCosts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT,
            cost REAL,
            timestamp TEXT)
            ''')

        # 创建表结构（如果表不存在）
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS EmployeeEngagement (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT,
            project_name TEXT,
            product_name TEXT,
            hours float,
            content TEXT,
            timestamp TEXT)''')

        conn.commit()


def insert_record(id, employee_name, cost, timestamp=None):
    if timestamp is None:
        timestamp = datetime.now().strftime("%Y-%m")
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute('''INSERT INTO EmployeeCosts (id,employee_name, cost, timestamp)VALUES (?, ?, ?, ?)''',
                   (id, employee_name, cost, timestamp))
    # 提交事务并关闭数据库连接
    conn.commit()
    conn.close()


def insert_record_gs(employee_name, project_name, product_name,hours, content, timestamp=None):
    if timestamp is None:
        timestamp = datetime.now().strftime("%Y-%m")
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute(
        '''INSERT INTO EmployeeEngagement (employee_name, project_name, product_name,hours,content,timestamp)VALUES (?, ?,?, ?,?,?)''',
        (employee_name, project_name, product_name,hours, content, timestamp))
    # 提交事务并关闭数据库连接
    conn.commit()
    conn.close()


def select_employee_project_costs(employee_name, timestamp):

    db_path = os.path.join('data', 'employee_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 构建 WHERE 条件
    conditions = []
    params = []  # 初始化为列表而不是元组

    if employee_name != "全部":
        conditions.append("a.employee_name LIKE ?")
        params.append("%" + employee_name + "%")

    if timestamp:
        conditions.append("a.timestamp = ?")  # 明确使用 a 表的 timestamp
        params.append(timestamp)

    where_clause = "WHERE " + " AND ".join(conditions) if conditions else ""

    query = f'''
    SELECT
      a.employee_name,
      a.project_name,
      a.product_name,
      SUM(a.hours) AS SUMHOURS,
      SUM(a.hours) / 8 AS RR,
      b.cost / c.SUMH  * SUM(a.hours) AS CB,
      a.timestamp
    FROM
      EmployeeEngagement a
      INNER JOIN EmployeeCosts b ON a.employee_name = b.employee_name AND a.timestamp = b.timestamp
      INNER JOIN (
        SELECT employee_name, SUM(hours) AS SUMH FROM EmployeeEngagement GROUP BY employee_name
      ) c ON b.employee_name = c.employee_name
    {where_clause}
    GROUP BY
      a.employee_name,
      a.project_name,
      a.product_name
    ORDER BY
      a.employee_name,
      SUMHOURS DESC
    '''

    # 执行查询
    cursor.execute(query, params)
    results = cursor.fetchall()

    conn.close()
    return results
if __name__ == "__main__":
    initialize_database()