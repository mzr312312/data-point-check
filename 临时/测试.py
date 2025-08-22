import psycopg2
import pandas as pd
from openpyxl import Workbook

# 数据库连接配置 - 请替换为您的实际凭据
db_config = {
    'host': '172.25.4.87',
    'port': '5432',
    'database': 'ems_test',
    'user': 'mazhuoran',  # 替换为您的用户名
    'password': 'mzrEMS1+1'  # 替换为您的密码
}


def export_tables_to_single_excel():
    try:
        # 创建数据库连接
        conn = psycopg2.connect(**db_config)

        # 创建Excel写入对象
        excel_file = "dim_tables.xlsx"
        writer = pd.ExcelWriter(excel_file, engine='openpyxl')

        # 要导出的表列表
        tables = ['dim_area', 'dim_base']

        for table in tables:
            print(f"正在导出表: {table}")

            # 读取表数据到DataFrame
            query = f"SELECT * FROM {table}"
            df = pd.read_sql(query, conn)

            # 导出到Excel的不同sheet
            df.to_excel(writer, sheet_name=table, index=False)
            print(f"表 {table} 已成功导出到 {excel_file} 的 {table} sheet")

        # 保存Excel文件
        writer.close()
        print(f"所有表已成功导出到 {excel_file}")

    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        # 确保连接关闭
        if 'conn' in locals():
            conn.close()
        print("导出过程完成")


# 执行导出
export_tables_to_single_excel()