# python -m venv ./env
# pyinstaller -F ./process.py -c --paths=./env/Lib/site-packages

import openpyxl
import xlrd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import time
import os
import threading
import pandas as pd

POLLING_DELAY = 250  # ms
FILEPATH_PLACEHOLDER = "点击此处选择文件"


class ExcelProcessorGUI:
    def __init__(self, master):
        self.master = master
        master.title("Excel 数据处理程序")

        # 创建文件选择框
        self.file_path_var = tk.StringVar()
        self.file_path_var.set(FILEPATH_PLACEHOLDER)
        self.file_select_label = tk.Label(
            master, textvariable=self.file_path_var, width=40, height=2, relief=tk.SUNKEN)
        self.file_select_label.grid(row=0, column=0, padx=5, pady=5)
        self.file_select_label.bind("<Button-1>", self.choose_file)

        # 创建处理按钮
        self.process_button = tk.Button(
            master, text="处理", command=self.process_excel)
        self.process_button.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # 创建状态栏和进度条
        self.status_label = tk.Label(master, text="请选择要处理的 Excel 文件")
        self.status_label.grid(row=2, column=0, padx=5, pady=5)

        self.lock = threading.Lock()
        self.finished = True
        self.df_out = None
        self.process_iter_max = 0
        self.process_iter = 0

        self.progress_bar = ttk.Progressbar(
            master, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=3, column=0, padx=5, pady=5)

    # 文件选择框绑定的事件回调，选择要处理的 Excel 文件
    def choose_file(self, event):
        self.process_button['text'] = "处理"
        file_path = filedialog.askopenfilename(
            title="请选择要处理的 Excel 文件", filetypes=[("Excel 文件", ["*.xlsx", "*.xls", "*.csv"])])
        if file_path:
            self.file_path_var.set(file_path)
            self.status_label.config(text="已选择文件：" + file_path)
            self.process_button['state'] = tk.NORMAL
        else:
            self.file_path_var.set(FILEPATH_PLACEHOLDER)
            self.status_label.config(text="请选择要处理的 Excel 文件")

    # 按钮上绑定的事件回调，处理 Excel 文件
    def process_excel(self):
        self.file_path = self.file_path_var.get()
        if self.file_path == FILEPATH_PLACEHOLDER:
            self.status_label.config(text="请选择要处理的 Excel 文件")
            return
        try:
            with self.lock:
                self.process_button['state'] = tk.DISABLED
                self.finished = False
                self.df_out = None

            t = threading.Thread(target=self.process_thread)
            t.daemon = True
            self.master.after(POLLING_DELAY, self.check_status)
            t.start()
        except Exception as e:
            # 处理结束后重置进度条和状态栏
            self.status_label.config(text=f"处理出错：{e}, 请重新选择要处理的 Excel 文件")


    # 轮询检查任务是否已经完成的回调，tkinter 只有单线程
    def check_status(self):
        with self.lock:
            if self.finished:
                self.progress_bar["value"] = self.process_iter_max
                self.process_button['state'] = tk.NORMAL
                # 保存 Excel 文件
                new_file_path = filedialog.asksaveasfilename(
                    title="请选择保存文件路径", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
                if new_file_path:
                    self.df_out.to_excel(new_file_path, index=False)
                    self.status_label.config(text="处理完成！已保存到：" + new_file_path)
                    self.file_path_var.set(FILEPATH_PLACEHOLDER)
                else:
                    self.status_label.config(text="处理完成！但未保存文件")
            else:
                # 继续轮询检查
                self.progress_bar["maximum"] = self.process_iter_max
                self.progress_bar["value"] = self.process_iter
                self.master.after(POLLING_DELAY, self.check_status)

    def process_thread(self):
        # 加载 Excel 表格
        df = self.load_excel()
        self.df_out = self.process_inner(df)
        with self.lock:
            self.finished = True

    def load_excel(self):
        print(f"加载文件中：{self.file_path}")
        if self.file_path.endswith(".xlsx") or self.file_path.endswith("xls"):
            df = pd.read_excel(self.file_path)
            return df
        elif self.file_path.endswith("csv"):
            df = pd.read_csv(self.file_path, on_bad_lines='warn')
            return df
        else:
            raise Exception

    def process_inner(self, df):
        df = df.fillna('')
        # 定义字典变量，用于记录每个人的所有检查项目和指标值
        data_dict = {}

        last_name = None
        last_project = None

        # 遍历 Excel 表格的每一行数据，
        # 因为我们用 self.finished 监听了后台线程是否结束，
        # 所以事实上只会有一个后台线程，这些数据不需要加锁，不会有并发问题
        self.process_iter_max = len(df)
        print(f"共计 {self.process_iter_max} 行数据")
        self.process_iter = 0
        for i, row in df.iterrows():
            self.process_iter = i
            row = list(row.to_dict().values())
            name, gender, age, phone, company, project, value, interval = row[:8]
            # 按照人名进行聚合，将每个人的所有检查项目和指标值放入一个列表中
            if project == '':
                # 心电图项目的数据可能会到下一行
                if last_name != '' and name == '':
                    data_dict[last_name][last_project] += name
                continue
            last_name = name
            last_project = project
            project_ref = project+"范围"
            if name in data_dict:
                data_dict[name][project] = value
                data_dict[name][project_ref] = interval
            else:
                data_dict[name] = {"姓名": name, "性别": gender, "年龄": age,
                                   "电话": phone, "单位": company, project: value, project_ref: interval}

        print(
            f"逐行加载完毕，开始进行分析")
        # 找出体检项目最多的一个人，
        # 或者搞个集合求并集
        project_cnt = 0
        projects = []
        for data in data_dict.values():
            if len(data.keys()) > project_cnt:
                project_cnt = len(data.keys())
                projects = list(data.keys())
        projects = projects[5:]

        headers = ["姓名", "性别", "年龄", "电话", "单位"] + \
            projects
        print(
            f"按姓名解析完毕，共解析到 {len(data_dict)} 人，每人 {len(headers)} 项数据，开始整理新表格")

        # 将字典转换为 Pandas 数据帧
        df_out = pd.DataFrame.from_dict(data_dict, orient='index')
        # 选择需要的列
        df_out = df_out[['姓名', '性别', '年龄', '电话', '单位'] + projects]
        '''
        df_out = pd.DataFrame(columns=headers)
        # 写入数据
        for i, data in enumerate(data_dict.values()):
            self.process_iter = i + 1
            row_data = [data["姓名"], data["性别"],
                        data["年龄"], data["电话"], data["单位"]]
            for header in projects:
                row_data.append(data.get(header, ""))
            df_out.append(row_data)
        '''
        return df_out


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()
