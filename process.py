# python -m venv ./env
# pyinstaller -F ./process.py -c --paths=./env/Lib/site-packages

import openpyxl
import xlrd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


class ExcelProcessorGUI:
    def __init__(self, master):
        self.master = master
        master.title("Excel 数据处理程序")

        # 创建文件选择框
        self.file_path = tk.StringVar()
        self.file_path.set("点击此处选择文件")
        self.file_select_label = tk.Label(
            master, textvariable=self.file_path, width=40, height=2, relief=tk.SUNKEN)
        self.file_select_label.grid(row=0, column=0, padx=5, pady=5)
        self.file_select_label.bind("<Button-1>", self.choose_file)

        # 创建处理按钮
        self.process_button = tk.Button(
            master, text="处理", command=self.process_excel)
        self.process_button.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # 创建状态栏和进度条
        self.status_label = tk.Label(master, text="请选择要处理的 Excel 文件")
        self.status_label.grid(row=2, column=0, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(
            master, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=3, column=0, padx=5, pady=5)

    def load_excel(self):
        try:
            file_path = self.file_path.get()
            if file_path.endswith(".xlsx"):
                workbook = openpyxl.load_workbook(file_path)
                sheet = workbook["Sheet1"]
                return sheet
            else:
                # 读取 xls 文件
                workbook = xlrd.open_workbook(file_path)
                worksheet = workbook.sheet_by_index(0)

                # 创建新的 xlsx 文件
                workbook_new = openpyxl.Workbook()
                worksheet_new = workbook_new.active

                # 复制数据到新文件
                for row in range(worksheet.nrows):
                    for col in range(worksheet.ncols):
                        value = worksheet.cell(row, col).value
                        worksheet_new.cell(
                            row=row+1, column=col+1, value=value)
                return worksheet_new
        except Exception as e:
            raise e

    # 选择要处理的 Excel 文件
    def choose_file(self, event):
        file_path = filedialog.askopenfilename(
            title="请选择要处理的 Excel 文件", filetypes=[("Excel 文件", ["*.xlsx", "*.xls"])])
        if file_path:
            self.file_path.set(file_path)
            self.status_label.config(text="已选择文件：" + file_path)
        else:
            self.file_path.set("点击此处选择文件")
            self.status_label.config(text="请选择要处理的 Excel 文件")

    # 处理 Excel 文件
    def process_excel(self):
        file_path = self.file_path.get()
        if not file_path:
            self.status_label.config(text="请选择要处理的 Excel 文件")
            return
        try:
            # 打开 Excel 表格
            sheet = self.load_excel()
            new_workbook = self.process_inner(sheet)
            # 保存 Excel 文件
            new_file_path = filedialog.asksaveasfilename(
                title="请选择保存文件路径", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
            if new_file_path:
                new_workbook.save(new_file_path)
                self.status_label.config(text="处理完成！已保存到：" + new_file_path)
            else:
                self.status_label.config(text="处理完成！但未保存文件")
        except Exception as e:
            # 处理结束后重置进度条和状态栏
            self.status_label.config(text=f"处理出错：{e}, 请重新选择要处理的 Excel 文件")
            self.progress_bar["value"] = 0

    def process_inner(self, sheet):
        try:
            # 定义字典变量，用于记录每个人的所有检查项目和指标值
            data_dict = {}

            last_name = None
            last_project = None
            print("开始解析原始数据")

            # 遍历 Excel 表格的每一行数据
            rows = sheet.iter_rows(min_row=2, values_only=True)
            self.progress_bar["maximum"] = sheet.max_row - 1
            self.progress_bar["value"] = 0
            for i, row in enumerate(rows):
                self.progress_bar["value"] = i + 1
                name, gender, age, phone, company, project, value, interval = row[:8]
                # 按照人名进行聚合，将每个人的所有检查项目和指标值放入一个列表中
                if project is None:
                    # 心电图项目的数据可能会到下一行
                    if last_name is not None and name is not None:
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

            # 找出体检项目最多的一个人，
            # 或者搞个集合求并集
            print("before project_cnt")
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
                f"解析完毕，共解析到 {len(data_dict)} 人，每人 {len(headers)} 项数据，开始整理新表格")

            # 创建新的 Excel 表格，并写入处理后的数据
            new_workbook = openpyxl.Workbook()
            new_sheet = new_workbook.active
            # 写入表头
            new_sheet.append(headers)

            # 写入数据
            self.progress_bar["maximum"] = len(data_dict)
            self.progress_bar["value"] = 0
            for i, data in enumerate(data_dict.values()):
                self.progress_bar["value"] = i + 1
                row_data = [data["姓名"], data["性别"],
                            data["年龄"], data["电话"], data["单位"]]
                for header in projects:
                    row_data.append(data.get(header, ""))
                new_sheet.append(row_data)
            return new_workbook
        except Exception as e:
            raise e


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()
