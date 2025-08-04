import pandas as pd
from datetime import datetime
from database import insert_record, insert_record_gs


def import_data_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, usecols=["序号", "姓名", "总计成本"])
    except ValueError:
        print("Excel 文件中缺少必要数据的列：员工名称，成本！")
        return
    if df.empty:
        print("Excel文件为空，未找到数据！")
    timestamp = datetime.now().strftime("%Y—%m")

    for _, row in df.iterrows():
        employee_name = row["姓名"],
        cost = row["总计成本"],
        timestamp = row.get('时间戳', datetime.now().strftime("%Y-%m"))

        insert_record(row["序号"], row["姓名"], row["总计成本"], timestamp)
    print("成本数据成功导入")


def import_data_from_excel_gs(file_path):
    try:
        df = pd.read_excel(file_path, usecols=["记录人", "项目", "产品","耗时", "工作内容"])
    except ValueError:
        print("Excel 文件中缺少必要数据的列：记录人，成本，耗时，项目，产品，工作内容！")
        return
    if df.empty:
        print("Excel文件为空，未找到数据！")
    timestamp = datetime.now().strftime("%Y—%m")

    for _, row in df.iterrows():
        employee_name = row["记录人"],
        project_name = row["项目"],
        product_name = row["产品"],
        hours = row["耗时"],
        content = row["工作内容"]
        timestamp = row.get('时间戳', datetime.now().strftime("%Y-%m"))

        insert_record_gs(row["记录人"], row["项目"], row["产品"],row["耗时"], row["工作内容"], timestamp)
    print("员工数据成功导入")
