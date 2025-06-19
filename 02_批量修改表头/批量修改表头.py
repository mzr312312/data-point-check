import os
import pandas as pd
from tkinter import Tk, Button, filedialog
from datetime import datetime  # 导入 datetime 模块

def load_and_transform_excel():
    # Step 1: 打开文件选择对话框，让用户选择目标 .xlsx 文件
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择要处理的 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx")]
    )
    if not file_path:
        print("未选择文件，程序退出。")
        return

    # Step 2: 加载目标 Excel 文件
    try:
        df = pd.read_excel(file_path)
        print("成功加载文件：", file_path)
    except Exception as e:
        print("加载文件失败：", e)
        return

    # Step 3: 加载列名字典文件
    dict_file = "列名字典.xlsx"
    if not os.path.exists(dict_file):
        print(f"未找到文件 {dict_file}，请确保它与脚本在同一目录下。")
        return

    try:
        # 读取“列名映射关系”工作表
        mapping_df = pd.read_excel(dict_file, sheet_name="列名映射关系")
        # 读取“列排序规则”工作表
        order_df = pd.read_excel(dict_file, sheet_name="列排序规则")

        # 构建旧列名到新列名的映射字典
        name_mapping = dict(zip(mapping_df['旧列名'], mapping_df['新列名']))
        # 构建排序规则字典
        order_mapping = dict(zip(order_df['列名'], order_df['排序序号']))

        print("成功加载列名字典文件：", dict_file)
    except Exception as e:
        print("加载列名字典文件失败：", e)
        return

    # Step 4: 修改列名并排序
    try:
        # 替换列名
        df.rename(columns=name_mapping, inplace=True)

        # 按照排序规则重新排列列
        sorted_columns = sorted(df.columns, key=lambda x: order_mapping.get(x, float('inf')))
        df = df[sorted_columns]

        # 清理空值，避免生成无效文件
        df.fillna("", inplace=True)

        print("列名修改和排序完成。")
    except Exception as e:
        print("列名修改或排序失败：", e)
        return

    # Step 5: 保存结果到新文件
    try:
        # 获取当前时间戳并格式化为字符串
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # 格式：YYYYMMDD_HHMMSS
        # 构造输出文件名，包含时间戳
        output_file = os.path.splitext(file_path)[0] + f"d_{timestamp}.xlsx"

        # 显式指定使用 openpyxl 引擎
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"处理完成，结果已保存到：{output_file}")
    except Exception as e:
        print("保存结果文件失败：", e)

# 创建 GUI 界面
if __name__ == "__main__":
    root = Tk()
    root.title("Excel 文件处理工具")
    root.geometry("300x100")

    btn = Button(root, text="选择 Excel 文件并处理", command=load_and_transform_excel)
    btn.pack(pady=20)

    root.mainloop()