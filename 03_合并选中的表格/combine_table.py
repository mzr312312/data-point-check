import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from datetime import datetime
import os
import threading


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel表格合并工具")
        root.geometry("700x600")

        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 11, "bold"))

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(self.main_frame, text="Excel表格合并工具", style="Header.TLabel")
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))

        # 列名行设置
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=5)

        ttk.Label(header_frame, text="列名所在行:").pack(side=tk.LEFT)
        self.header_row = tk.IntVar(value=1)
        header_spinbox = ttk.Spinbox(header_frame, from_=1, to=10, width=5,
                                     textvariable=self.header_row)
        header_spinbox.pack(side=tk.LEFT, padx=5)

        # 文件选择区域
        file_frame = ttk.Frame(self.main_frame)
        file_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=10)

        ttk.Button(file_frame, text="选择Excel文件",
                   command=self.select_files).pack(side=tk.LEFT, padx=5)

        self.file_list = tk.Listbox(file_frame, width=80, height=5, selectmode=tk.EXTENDED)
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.file_list.config(yscrollcommand=scrollbar.set)

        # 日志区域
        ttk.Label(self.main_frame, text="操作日志:").grid(row=3, column=0, sticky="w", pady=(15, 5))
        self.log_area = scrolledtext.ScrolledText(self.main_frame, width=85, height=15)
        self.log_area.grid(row=4, column=0, columnspan=3, sticky="we")

        # 按钮区域
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=15)

        ttk.Button(button_frame, text="开始合并",
                   command=self.start_merge_thread).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame, text="清除日志",
                   command=self.clear_log).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame, text="退出",
                   command=root.destroy).pack(side=tk.LEFT, padx=10)

        # 进度条
        self.progress = ttk.Progressbar(self.main_frame, orient="horizontal",
                                        length=500, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=15)

        # 状态标签
        self.status_label = ttk.Label(self.main_frame, text="就绪", foreground="green")
        self.status_label.grid(row=7, column=0, columnspan=3)

        # 初始化
        self.selected_files = []
        self.log("就绪: 请选择Excel文件进行合并")

    def log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """清除日志"""
        self.log_area.delete(1.0, tk.END)
        self.log("日志已清除")

    def select_files(self):
        """选择要合并的Excel文件"""
        files = filedialog.askopenfilenames(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )

        if files:
            self.selected_files = list(files)
            self.file_list.delete(0, tk.END)
            for file in self.selected_files:
                # 只显示文件名，不显示完整路径
                filename = os.path.basename(file)
                self.file_list.insert(tk.END, filename)

            file_count = len(self.selected_files)
            self.log(f"已选择 {file_count} 个Excel文件")
            self.update_status(f"已选择 {file_count} 个文件", "green")

    def update_status(self, message, color="black"):
        """更新状态标签"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()

    def start_merge_thread(self):
        """启动合并线程以防止GUI冻结"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要合并的Excel文件")
            return

        # 禁用按钮
        self.update_status("正在合并...", "blue")
        self.progress["value"] = 0

        # 在新线程中运行合并
        thread = threading.Thread(target=self.merge_excel_files)
        thread.daemon = True
        thread.start()

    def merge_excel_files(self):
        """执行Excel文件合并"""
        try:
            header_row = self.header_row.get() - 1  # pandas使用0-indexed

            # 收集所有文件中所有sheet的信息
            sheets_data = {}
            total_files = len(self.selected_files)

            for file_index, file_path in enumerate(self.selected_files):
                self.log(f"处理文件: {os.path.basename(file_path)}")
                self.progress["value"] = (file_index / total_files) * 50
                self.update_status(f"正在处理文件 {file_index + 1}/{total_files}", "blue")

                # 读取Excel文件
                try:
                    with pd.ExcelFile(file_path) as xls:
                        sheet_names = xls.sheet_names
                        if not sheet_names:
                            self.log(f"  - 警告: 文件不包含任何工作表")
                            continue

                        for sheet_name in sheet_names:
                            try:
                                # 读取工作表数据
                                df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

                                # 忽略空工作表
                                if df.empty:
                                    self.log(f"  - 跳过空工作表: {sheet_name}")
                                    continue

                                # 添加源文件信息
                                df["来源文件"] = os.path.basename(file_path)

                                # 修复: 正确分组同名sheet
                                # 使用原始sheet_name作为分组键
                                if sheet_name not in sheets_data:
                                    # 新sheet名称分组
                                    sheets_data[sheet_name] = {
                                        "data": [df],
                                        "columns": set(df.columns),
                                        "files": [file_path]
                                    }
                                else:
                                    # 同名sheet存在，合并到同一分组
                                    sheets_data[sheet_name]["data"].append(df)
                                    sheets_data[sheet_name]["columns"].update(df.columns)
                                    sheets_data[sheet_name]["files"].append(file_path)

                            except Exception as e:
                                self.log(f"  - 错误处理工作表 '{sheet_name}': {str(e)}")
                except Exception as e:
                    self.log(f"  - 错误处理文件 '{file_path}': {str(e)}")

            # 如果没有任何sheet数据
            if not sheets_data:
                self.log("警告: 未找到任何有效数据")
                self.update_status("未找到有效数据", "orange")
                messagebox.showwarning("警告", "未在所选文件中找到任何有效数据")
                return

            # 合并数据
            merged_data = {}
            total_sheets = len(sheets_data)

            for sheet_index, (sheet_name, data) in enumerate(sheets_data.items()):
                self.progress["value"] = 50 + (sheet_index / total_sheets) * 50
                self.update_status(f"正在合并工作表 {sheet_index + 1}/{total_sheets}", "blue")

                # 检查该分组中所有DataFrame是否有相同的列
                same_columns = True
                for df in data["data"]:
                    if set(df.columns) != data["columns"]:
                        same_columns = False
                        break

                if same_columns:
                    # 列名相同的情况下，直接连接
                    merged_df = pd.concat(data["data"], ignore_index=True)
                    self.log(f"合并工作表: {sheet_name} (相同列, 来自 {len(data['files'])} 个文件)")
                else:
                    # 列名不同，需要特殊处理
                    # 创建包含所有列的空DataFrame
                    all_columns = sorted(data["columns"])  # 排序以确保顺序一致
                    merged_df = pd.DataFrame(columns=all_columns)

                    # 依次添加每个df，填充缺失列
                    for df in data["data"]:
                        # 确保所有列都存在
                        for col in all_columns:
                            if col not in df.columns:
                                df[col] = np.nan
                        # 确保列顺序一致
                        df = df[all_columns]
                        merged_df = pd.concat([merged_df, df], ignore_index=True)

                    self.log(f"合并工作表: {sheet_name} (不同列, 来自 {len(data['files'])} 个文件)")

                # 确保列的顺序一致
                all_columns = list(data["columns"])
                if "来源文件" in all_columns:
                    all_columns.remove("来源文件")
                    all_columns.append("来源文件")
                merged_df = merged_df[all_columns]

                merged_data[sheet_name] = merged_df

            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"合并表格_{timestamp}.xlsx"

            # 确保输出目录存在
            if not os.path.exists("output"):
                os.makedirs("output")

            output_path = os.path.join("output", output_file)

            # 保存合并后的Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in merged_data.items():
                    # 限制sheet名称长度（Excel要求31个字符以内）
                    safe_sheet_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                    df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

            self.log(f"合并完成! 文件保存至: {output_path}")
            self.progress["value"] = 100
            self.update_status(f"合并完成: {output_file}", "green")
            messagebox.showinfo("完成", f"合并完成!\n文件已保存至:\n{output_path}")

        except Exception as e:
            self.log(f"合并过程中发生错误: {str(e)}")
            self.update_status(f"错误: {str(e)}", "red")
            messagebox.showerror("错误", f"合并过程中发生错误:\n{str(e)}")
            self.progress["value"] = 0


def main():
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()