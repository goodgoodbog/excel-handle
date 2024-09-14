# -*- coding=utf-8-*-
# 功能4：对同一运单号的商品进行合并
import os
import time
from tkinter import filedialog, ttk
import tkinter as tk
import polars as pl
import openpyxl
import python_calamine
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use


class merge:
    def __init__(self, root):
        use = all_use.Use(root)
        self.f_list = 0
        self.f1 = tk.Label(root, text='功能4：合并信息：')
        self.f_duizhao_label = tk.Label(root, text='产品中英文对照表:')
        self.f_dingdan1_label = tk.Label(root, text='订单列表1:')
        self.f_dingdan2_label = tk.Label(root, text='订单列表2:')
        self.f_dingdan3_label = tk.Label(root, text='订单列表3:')
        self.f_dingdan4_label = tk.Label(root, text='订单列表4:')
        self.f_dingdan5_label = tk.Label(root, text='订单列表5:')
        self.f_dingdan6_label = tk.Label(root, text='订单列表6:')
        self.f_dingdan7_label = tk.Label(root, text='订单列表7:')
        self.f_dingdan8_label = tk.Label(root, text='订单列表8:')
        self.f_dingdan9_label = tk.Label(root, text='订单列表9:')
        self.f_save_label = tk.Label(root, text='保存到:')
        self.f_jg_label = tk.Label(root, text='结果提示:')

        self.f_duizhao_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                        name='f_duizhao_entry', validate="key", relief='sunken',
                                        validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan1_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan1_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_save_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.f_jg_entry = tk.Entry(root, width=56)
        self.f_dingdan2_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan2_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan3_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan3_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan4_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan4_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan5_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan5_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan6_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan6_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan7_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan7_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan8_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan8_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan9_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='f_dingdan9_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        # 选择文件按钮
        self.f_choose_duizhao = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_duizhao_entry))
        self.f_Button_choose_files = ttk.Button(root, text='选择多个表', command=self.f_fun_choose_files)
        self.f_choose_add = ttk.Button(root, text='订单列表+1', command=self.f_add)
        self.f_choose_dingdan1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan1_entry))
        self.f_choose_dingdan2 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan2_entry))
        self.f_choose_dingdan3 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan3_entry))
        self.f_choose_dingdan4 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan4_entry))
        self.f_choose_dingdan5 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan5_entry))
        self.f_choose_dingdan6 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan6_entry))
        self.f_choose_dingdan7 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan7_entry))
        self.f_choose_dingdan8 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan8_entry))
        self.f_choose_dingdan9 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan9_entry))
        self.f_choose_word = ttk.Button(root, text='开始对照', command=self.work)

    # 布局函数
    # @grid_progressbar
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.f1.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.f1)
        self.f_duizhao_label.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.f_duizhao_label)
        self.f_duizhao_entry.delete("0", 'end')
        self.f_duizhao_entry.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.f_duizhao_entry)
        self.f_choose_duizhao.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.f_choose_duizhao)
        self.f_dingdan1_label.grid(row=4, column=0, sticky='e')
        all_use.tk_list.append(self.f_dingdan1_label)
        self.f_dingdan1_entry.delete("0", 'end')
        self.f_dingdan1_entry.grid(row=4, column=1, sticky='w')
        all_use.tk_list.append(self.f_dingdan1_entry)
        self.f_choose_dingdan1.grid(row=4, column=2, sticky='w')
        all_use.tk_list.append(self.f_choose_dingdan1)

        self.f_Button_choose_files.grid(row=4, column=3, sticky='w')
        all_use.tk_list.append(self.f_Button_choose_files)
        self.f_choose_add.grid(row=4, column=4, sticky='w')
        all_use.tk_list.append(self.f_choose_add)
        self.f_save_label.grid(row=14, column=0, sticky='e')
        all_use.tk_list.append(self.f_save_label)
        self.f_save_entry.delete("0", 'end')
        self.f_save_entry.grid(row=14, column=1, sticky='w')
        all_use.tk_list.append(self.f_save_entry)
        self.f_choose_word.grid(row=15, column=1)
        all_use.tk_list.append(self.f_choose_word)
        self.f_jg_label.grid(row=16, column=0, sticky='e')
        all_use.tk_list.append(self.f_jg_label)
        self.f_jg_entry.delete("0", 'end')
        self.f_jg_entry.grid(row=16, column=1, sticky='w')
        all_use.tk_list.append(self.f_jg_entry)

        self.f_list = 8
        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.f_save_entry.delete("0", 'end')
        self.f_save_entry.insert(tk.INSERT, '对照-' + now_time)

    # 一次性选择多个文件
    def f_fun_choose_files(self):
        file_paths = filedialog.askopenfilenames()
        if file_paths == '':
            return
        print(file_paths)
        self.f_list = 8
        more_list = [[self.f_dingdan2_label, self.f_dingdan2_entry, self.f_choose_dingdan2],
                     [self.f_dingdan3_label, self.f_dingdan3_entry, self.f_choose_dingdan3],
                     [self.f_dingdan4_label, self.f_dingdan4_entry, self.f_choose_dingdan4],
                     [self.f_dingdan5_label, self.f_dingdan5_entry, self.f_choose_dingdan5],
                     [self.f_dingdan6_label, self.f_dingdan6_entry, self.f_choose_dingdan6],
                     [self.f_dingdan7_label, self.f_dingdan7_entry, self.f_choose_dingdan7],
                     [self.f_dingdan8_label, self.f_dingdan8_entry, self.f_choose_dingdan8],
                     [self.f_dingdan9_label, self.f_dingdan9_entry, self.f_choose_dingdan9], ]
        more_list = iter(more_list)
        self.f_dingdan1_entry.delete("0", 'end')
        self.f_dingdan1_entry.insert(tk.INSERT, file_paths[0])
        now_list = next(more_list)
        list2 = [self.f_dingdan2_entry,
                 self.f_dingdan3_entry,
                 self.f_dingdan4_entry,
                 self.f_dingdan5_entry,
                 self.f_dingdan6_entry,
                 self.f_dingdan7_entry,
                 self.f_dingdan8_entry,
                 self.f_dingdan9_entry,
                 ]
        for file_path, entry in zip(file_paths[1:], list2):
            entry.delete("0", 'end')
            self.f_add()
            entry.insert(tk.INSERT, file_path)
            now_list = next(more_list)
        while now_list is not None:
            for data in now_list:
                if data.widgetName == 'entry':
                    data.delete("0", 'end')
                data.grid_forget()
            try:
                now_list = next(more_list)
            except StopIteration:
                break

    # 文件数加1
    def f_add(self):
        i = self.f_list
        if i == 8:
            self.f_dingdan2_label.grid(row=5, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan2_label)
            self.f_dingdan2_entry.delete("0", 'end')
            self.f_dingdan2_entry.grid(row=5, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan2_entry)
            self.f_choose_dingdan2.grid(row=5, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan2)
            self.f_list = self.f_list - 1
        elif i == 7:
            self.f_dingdan3_label.grid(row=6, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan3_label)
            self.f_dingdan3_entry.delete("0", 'end')
            self.f_dingdan3_entry.grid(row=6, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan3_entry)
            self.f_choose_dingdan3.grid(row=6, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan3)
            self.f_list = self.f_list - 1
        elif i == 6:
            self.f_dingdan4_label.grid(row=7, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan4_label)
            self.f_dingdan4_entry.delete("0", 'end')
            self.f_dingdan4_entry.grid(row=7, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan4_entry)
            self.f_choose_dingdan4.grid(row=7, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan4)
            self.f_list = self.f_list - 1
        elif i == 5:
            self.f_dingdan5_label.grid(row=8, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan5_label)
            self.f_dingdan5_entry.delete("0", 'end')
            self.f_dingdan5_entry.grid(row=8, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan5_entry)
            self.f_choose_dingdan5.grid(row=8, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan5)
            self.f_list = self.f_list - 1
        elif i == 4:
            self.f_dingdan6_label.grid(row=9, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan6_label)
            self.f_dingdan6_entry.delete("0", 'end')
            self.f_dingdan6_entry.grid(row=9, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan6_entry)
            self.f_choose_dingdan6.grid(row=9, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan6)
            self.f_list = self.f_list - 1
        elif i == 3:
            self.f_dingdan7_label.grid(row=11, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan7_label)
            self.f_dingdan7_entry.delete("0", 'end')
            self.f_dingdan7_entry.grid(row=11, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan7_entry)
            self.f_choose_dingdan7.grid(row=11, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan7)
            self.f_list = self.f_list - 1
        elif i == 2:
            self.f_dingdan8_label.grid(row=12, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan8_label)
            self.f_dingdan8_entry.delete("0", 'end')
            self.f_dingdan8_entry.grid(row=12, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan8_entry)
            self.f_choose_dingdan8.grid(row=12, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan8)
            self.f_list = self.f_list - 1
        elif i == 1:
            self.f_dingdan9_label.grid(row=13, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan9_label)
            self.f_dingdan9_entry.delete("0", 'end')
            self.f_dingdan9_entry.grid(row=13, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan9_entry)
            self.f_choose_dingdan9.grid(row=13, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan9)
            self.f_list = self.f_list - 1

    # 数据处理函数
    @creat_thread
    def work(self):
        try:
            wb1 = python_calamine.CalamineWorkbook.from_path(self.f_duizhao_entry.get())
            rows1 = list(wb1.get_sheet_by_index(0).to_python())
        except:
            self.f_jg_entry.delete("0", 'end')
            self.f_jg_entry.insert(tk.INSERT, '对照文件选择出错,请重新选择')
            self.f_jg_entry.configure(fg='red')
            return
        dict1 = {}
        for row in rows1:
            dict1[row[0]] = list(filter(None, row[1:]))

        df = all_use.read_excel_to_dataframe([self.f_dingdan1_entry, self.f_dingdan2_entry, self.f_dingdan3_entry,
                                              self.f_dingdan4_entry, self.f_dingdan5_entry,
                                              self.f_dingdan6_entry, self.f_dingdan7_entry,
                                              self.f_dingdan8_entry, self.f_dingdan9_entry],
                                             ['订单列表1', '订单列表2', '订单列表3', '订单列表4', '订单列表5',
                                              '订单列表6',
                                              '订单列表7', '订单列表8', '订单列表9'], 2,
                                             ['订单编号', '订单创建时间', '物流费用', '商品名称', '款式1', '款式2',
                                              '款式3'],
                                             self.f_jg_entry)
        if df is None:
            return
        df = df.cast(
            {'物流费用': pl.Float64, "商品名称": pl.String, '款式1': pl.String, '款式2': pl.String,
             '款式3': pl.String, '售出数量': pl.Int32},
            strict=False)

        # 统计各个订单编号的商品
        df_end2 = df.with_columns((pl.col('商品名称')
                                   + pl.when(pl.col('款式1').is_not_null()).then(' ' + pl.col('款式1')).otherwise(
                    pl.col('款式1'))
                                   + pl.when(pl.col('款式2').is_not_null()).then(' ' + pl.col('款式2')).otherwise(
                    pl.col('款式2'))
                                   + pl.when(pl.col('款式3').is_not_null()).then(' ' + pl.col('款式3')).otherwise(
                    pl.col('款式3'))).alias('商品名称加款式'))
        df_end2 = df_end2.select(['订单编号', '商品名称加款式', '售出数量'])
        df_end2 = df_end2.with_columns(pl.col('商品名称加款式').str.strip_chars_end(' '))
        salary_list = [list(salary) for salary in df_end2.rows()]

        dict_end = {}
        for list1 in salary_list:
            while list1[2] != 0:
                if list1[0] in dict_end:
                    if list1[1] in dict1:
                        dict_end[list1[0]].extend(dict1[list1[1]])
                    else:
                        dict_end[list1[0]].append(list1[1])
                else:
                    dict_end[list1[0]] = [list1[1]]
                    if list1[1] in dict1:
                        dict_end[list1[0]].extend(dict1[list1[1]])
                    else:
                        dict_end[list1[0]].append(list1[1])
                list1[2] = list1[2] - 1

        list_end = []
        for key, value in dict_end.items():
            value.insert(0, key)
            list_end.append(value)

        df1 = df.unique('订单编号', keep='first')
        df1 = df1.select(['订单编号', '订单创建时间', '物流费用'])
        dict_time = {}
        dict_cost = {}
        for list1 in df1.rows():
            dict_time[list1[0]] = list1[1]
            dict_cost[list1[0]] = list1[2]

        # 添加商品数量
        for index, list1 in enumerate(list_end):
            list_end[index].insert(1, len(list1) - 1)
            list_end[index].insert(1, dict_cost[list1[0]])
            list_end[index].insert(1, dict_time[list1[0]][5:10])

        full_path = os.path.join(os.path.dirname(self.f_duizhao_entry.get()), self.f_save_entry.get() + '.xlsx')

        wb = openpyxl.Workbook()
        sheet1 = wb.active
        for data in list_end:
            sheet1.append(data)
        try:
            wb.save(full_path)
        except:
            self.e10.delete("0", 'end')
            self.e10.insert(tk.INSERT, '对照表保存失败')
            self.e10.configure(fg='red')
            return

        self.f_jg_entry.delete("0", 'end')
        self.f_jg_entry.insert(tk.INSERT,
                               '已完成，已保存到：' + self.f_save_entry.get() + '.xlsx')
        self.f_jg_entry.configure(fg='green')
        return