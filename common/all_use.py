# -*- coding=utf-8-*-
# 通用类和模块
import threading
from functools import wraps
import tkinter.filedialog

import openpyxl
import python_calamine
import polars as pl
import tkinter as tk

# 全局变量，控制组件的销毁
tk_list = []


class Use:
    def __init__(self, root):
        self.check_open_file1 = root.register(self.check_open_file)
        self.tk_list = tk_list

    ##提前检查文件是否可以正常打开
    def check_open_file(self, content, Entry_name):

        Entry_name = Entry_name[1:]
        for Component in self.tk_list:
            if Component.widgetName == 'entry' and Component.winfo_name() == Entry_name:
                try:
                    # print(Component)
                    if content[-5:] != '.xlsx':
                        Component.configure(fg='red')
                        return 1
                    python_calamine.CalamineWorkbook.from_path(content)
                    # wb1 = openpyxl.load_workbook(content)
                    Component.configure(fg='black')
                except:
                    Component.configure(fg='red')
                return 1


def div_forget(component):
    print('1')
    # print(component)
    if component.widgetName == 'entry':
        print(component.get())
        component.delete("0", 'end')
    component.forget()


# 选择文件
def choose_flie(entry):
    file_name = tkinter.filedialog.askopenfilename()
    if file_name == '':
        return
    file_name = file_name.replace("/", "\\")
    entry.delete("0", 'end')
    entry.insert(tk.INSERT, file_name)
    ##显示末尾
    entry.xview_moveto(1)


# 装饰器：创建线程
def creat_thread(func):
    @wraps(func)
    def wrapTheFunction(*args, **kwargs):
        # print('线程')
        new_thread2 = threading.Thread(target=lambda: func(*args, **kwargs), name='l1')
        new_thread2.daemon = 1
        new_thread2.start()

    return wrapTheFunction


# 读取excel文件到polars dataframe
def read_excel_to_dataframe(entry_list: list, entry_name_list: list, title_began: int, title_need_list: list,
                            jg_entry: object, rows_end=0, save_first_rows=False):
    # entry_list：组件列表，entry_name_list：组件列表各对应的名字，title_began：表头开始的行数，title_need_list：需要使用到的列名（用于检测是否存在某列）
    # jg_entry：结果提示组件，rows_end：需要删除的表最后多少行，save_first_rows：保留并返回表头开始的前几行
    if type(entry_list) != list:
        entry_list = [entry_list]
    if type(entry_name_list) != list:
        entry_name_list = [entry_name_list]
    if type(title_need_list) != list:
        title_need_list = [title_need_list]

    if len(entry_list) != len(entry_name_list):
        print('传入的entry_list和entry_name_list长度不一样!!!')
        return (None, None) if save_first_rows else None

    df_list = []
    df = None
    rows = ''
    for index, entry in enumerate(entry_list):
        print(entry)
        print(entry.get())
        if entry.get() != '':
            try:
                wb = python_calamine.CalamineWorkbook.from_path(entry.get())
                rows = list(wb.get_sheet_by_index(0).to_python())
                # 解决python_calamine读取大额数时转为科学计数法的问题，暂用
                # for index1, data1 in enumerate(rows):
                #     try:
                #         rows[index1][0] = int(rows[index1][0])
                #     except:
                #         pass
                ##
                if title_began > 0:
                    for index1, data in enumerate(rows[title_began - 1]):
                        if data in rows[title_began - 1][index1 + 1:]:
                            rows[title_began - 1][index1] = str(rows[title_began - 1][index1]) + str(index1)

                if rows_end == 0 and title_began != 0:
                    df = pl.DataFrame(data=rows[title_began:], schema=rows[title_began - 1], orient='row',
                                      infer_schema_length=None)
                elif rows_end != 0 and title_began != 0:
                    df = pl.DataFrame(data=rows[title_began:-rows_end], schema=rows[title_began - 1], orient='row',
                                      infer_schema_length=None)
                else:
                    df = pl.DataFrame(data=rows, orient='row', infer_schema_length=None)

            except:
                entry.configure(fg='darkorange')
                jg_entry.configure(fg='red')
                jg_entry.delete("0", 'end')
                jg_entry.insert(tk.INSERT, entry_name_list[index] + '出错，请重新选择！！！')
                return (None, None) if save_first_rows else None
            for data in title_need_list:
                if data not in df.columns:
                    entry.configure(fg='darkorange')
                    jg_entry.configure(fg='red')
                    jg_entry.delete("0", 'end')
                    jg_entry.insert(tk.INSERT, entry_name_list[index] + '中不存在' + data + '，请重新选择！！！')
                    return (None, None) if save_first_rows else None
            df_list.append(df)
        elif index == 0:
            jg_entry.configure(fg='red')
            jg_entry.delete("0", 'end')
            jg_entry.insert(tk.INSERT, entry_name_list[index] + '为空，请重新选择！！！')
            return (None, None) if save_first_rows else None
    if len(df_list) > 1:
        df = pl.concat(df_list, how="vertical_relaxed")
    # print(df)
    # print(save_first_rows)
    return (rows[:title_began], df) if save_first_rows else df


# 通过openpyxl保存polars dataframe到excel
def polarsdf_openpyxl_to_excel(df: pl.DataFrame, filename: str, need_title=True, first_row=None):
    # df:需要保存的dataframe，filename：保存的文件名，need_title：是否需要保留表头（True表示保留，Flase表示不保留）
    salary_list = [list(salary) for salary in df.rows()]
    # 添加表头
    if need_title:
        salary_list.insert(0, list(df.columns))
    # 当first_row不为None时，向表前加入数据
    if first_row is not None:
        salary_list = [item for sublist in [first_row, salary_list] for item in sublist]

    # 使用openpyxl保存数据到excel
    wb = openpyxl.Workbook()
    sheet1 = wb.active
    for data in salary_list:
        sheet1.append(data)
    try:
        wb.save(filename)
    except:
        return False
    return True
