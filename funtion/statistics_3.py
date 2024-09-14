# -*- coding=utf-8-*-
# 功能11：根据订单表中的订单编号，统计对应信息生成统计表
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import polars as pl
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use


class Statistics_3:
    def __init__(self, root):
        use = all_use.Use(root)
        self.Statistical_tables_Labeltip = tk.Label(root, text='功能11：根据订单表中的订单编号，统计对应信息生成统计表')
        # 订单列表
        self.Statistical_tables_Label1 = tk.Label(root, text='订单列表1：')
        self.Statistical_tables_Entry1 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                  name='statistical_tables_Entry1', validate="key",
                                                  relief='sunken',
                                                  validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1 = ttk.Button(root, text='选择文件',
                                                     command=lambda: choose_flie(self.Statistical_tables_Entry1))

        self.Statistical_tables_Button_add1 = ttk.Button(root, text='订单列表+1', command=self.Statistical_tables_add1)
        self.Statistical_tables_Button_choose_files1 = ttk.Button(root, text='选择多个订单列表',
                                                                  command=self.Statistical_tables_choose_files1)

        self.Statistical_tables_Label1_2 = tk.Label(root, text='订单列表2：')
        self.Statistical_tables_Entry1_2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_2', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_2 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_2))

        self.Statistical_tables_Label1_3 = tk.Label(root, text='订单列表3：')
        self.Statistical_tables_Entry1_3 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_3', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_3 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_3))

        self.Statistical_tables_Label1_4 = tk.Label(root, text='订单列表4：')
        self.Statistical_tables_Entry1_4 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_4', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_4 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_4))

        self.Statistical_tables_Label1_5 = tk.Label(root, text='订单列表5：')
        self.Statistical_tables_Entry1_5 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_5', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_5 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_5))

        self.Statistical_tables_Label1_6 = tk.Label(root, text='订单列表6：')
        self.Statistical_tables_Entry1_6 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_6', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_6 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_6))

        self.Statistical_tables_Label1_7 = tk.Label(root, text='订单列表7：')
        self.Statistical_tables_Entry1_7 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_7', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_7 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_7))

        self.Statistical_tables_Label1_8 = tk.Label(root, text='订单列表8：')
        self.Statistical_tables_Entry1_8 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_8', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_8 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_8))

        self.Statistical_tables_Label1_9 = tk.Label(root, text='订单列表9：')
        self.Statistical_tables_Entry1_9 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry1_9', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose1_9 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry1_9))
        # 运单明细表
        self.Statistical_tables_Label2 = tk.Label(root, text='运单明细表：')
        self.Statistical_tables_Entry2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                  name='statistical_tables_Entry2', validate="key",
                                                  relief='sunken',
                                                  validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose2 = ttk.Button(root, text='选择文件',
                                                     command=lambda: choose_flie(self.Statistical_tables_Entry2))
        # 商品成本表
        self.Statistical_tables_Label_shangpinchengben = tk.Label(root, text='商品成本表：')
        self.Statistical_tables_Entry_shangpinchengben = tk.Entry(root, width=56, highlightthickness=1.5,
                                                                  highlightcolor='blue',
                                                                  name='statistical_tables_Entry_shangpinchengben',
                                                                  validate="key",
                                                                  relief='sunken',
                                                                  validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_shangpinchengben = ttk.Button(root, text='选择文件', command=lambda: choose_flie(
            self.Statistical_tables_Entry_shangpinchengben))
        # pp订单表
        self.Statistical_tables_Label3 = tk.Label(root, text='pp订单表1：')
        self.Statistical_tables_Entry3 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                  name='statistical_tables_Entry3', validate="key",
                                                  relief='sunken',
                                                  validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose3 = ttk.Button(root, text='选择文件',
                                                     command=lambda: choose_flie(self.Statistical_tables_Entry3))

        self.Statistical_tables_Button_add = ttk.Button(root, text='pp订单表+1', command=self.Statistical_tables_add)
        self.Statistical_tables_Button_choose_files = ttk.Button(root, text='选择多个pp订单表',
                                                                 command=self.Statistical_tables_choose_files)
        self.Statistical_tables_Label_p2 = tk.Label(root, text='pp订单表2：')
        self.Statistical_tables_Entry_p2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p2', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p2 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p2))
        self.Statistical_tables_Label_p3 = tk.Label(root, text='pp订单表3：')
        self.Statistical_tables_Entry_p3 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p3', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p3 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p3))
        self.Statistical_tables_Label_p4 = tk.Label(root, text='pp订单表4：')
        self.Statistical_tables_Entry_p4 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p4', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p4 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p4))
        self.Statistical_tables_Label_p5 = tk.Label(root, text='pp订单表5：')
        self.Statistical_tables_Entry_p5 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p5', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p5 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p5))
        self.Statistical_tables_Label_p6 = tk.Label(root, text='pp订单表6：')
        self.Statistical_tables_Entry_p6 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p6', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p6 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p6))
        self.Statistical_tables_Label_p7 = tk.Label(root, text='pp订单表7：')
        self.Statistical_tables_Entry_p7 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p7', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p7 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p7))
        self.Statistical_tables_Label_p8 = tk.Label(root, text='pp订单表8：')
        self.Statistical_tables_Entry_p8 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p8', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p8 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p8))
        self.Statistical_tables_Label_p9 = tk.Label(root, text='pp订单表9：')
        self.Statistical_tables_Entry_p9 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                                    name='statistical_tables_Entry_p9', validate="key",
                                                    relief='sunken',
                                                    validatecommand=(use.check_open_file1, '%P', '%W'))
        self.Statistical_tables_Choose_p9 = ttk.Button(root, text='选择文件',
                                                       command=lambda: choose_flie(self.Statistical_tables_Entry_p9))

        self.Statistical_tables_Label_save = tk.Label(root, text='保存到：')
        self.Statistical_tables_Entry_save = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.Statistical_tables_Choose_work = ttk.Button(root, text='开始统计', command=self.work)
        self.Statistical_tables_Label_jg = tk.Label(root, text='结果：')
        self.Statistical_tables_Entry_jg = tk.Entry(root, width=56)

    # 布局函数
    # @grid_progressbar
    def show(self):
        global Statistical_tables_add_list1, Statistical_tables_add_list
        Statistical_tables_add_list1 = 8
        Statistical_tables_add_list = 8
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.Statistical_tables_Labeltip.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.Statistical_tables_Labeltip)
        # 订单列表
        self.Statistical_tables_Label1.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label1)
        self.Statistical_tables_Entry1.delete("0", 'end')
        self.Statistical_tables_Entry1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry1)
        self.Statistical_tables_Choose1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Choose1)

        self.Statistical_tables_Button_choose_files1.grid(row=3, column=3, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Button_choose_files1)
        self.Statistical_tables_Button_add1.grid(row=3, column=4, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Button_add1)

        # 运单明细表
        self.Statistical_tables_Label2.grid(row=12, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label2)
        self.Statistical_tables_Entry2.delete("0", 'end')
        self.Statistical_tables_Entry2.grid(row=12, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry2)
        self.Statistical_tables_Choose2.grid(row=12, column=2, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Choose2)
        # 商品成本表
        self.Statistical_tables_Label_shangpinchengben.grid(row=13, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label_shangpinchengben)
        self.Statistical_tables_Entry_shangpinchengben.delete("0", 'end')
        self.Statistical_tables_Entry_shangpinchengben.grid(row=13, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry_shangpinchengben)
        self.Statistical_tables_Choose_shangpinchengben.grid(row=13, column=2, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Choose_shangpinchengben)

        self.Statistical_tables_Label3.grid(row=14, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label3)
        self.Statistical_tables_Entry3.delete("0", 'end')
        self.Statistical_tables_Entry3.grid(row=14, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry3)
        self.Statistical_tables_Choose3.grid(row=14, column=2, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Choose3)

        self.Statistical_tables_Button_choose_files.grid(row=14, column=3, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Button_choose_files)
        self.Statistical_tables_Button_add.grid(row=14, column=4, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Button_add)

        self.Statistical_tables_Label_save.grid(row=24, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label_save)
        self.Statistical_tables_Entry_save.delete("0", 'end')
        self.Statistical_tables_Entry_save.grid(row=24, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry_save)
        self.Statistical_tables_Choose_work.grid(row=25, column=1)
        all_use.tk_list.append(self.Statistical_tables_Choose_work)
        self.Statistical_tables_Label_jg.grid(row=26, column=0, sticky='e')
        all_use.tk_list.append(self.Statistical_tables_Label_jg)
        self.Statistical_tables_Entry_jg.delete("0", 'end')
        self.Statistical_tables_Entry_jg.grid(row=26, column=1, sticky='w')
        all_use.tk_list.append(self.Statistical_tables_Entry_jg)

        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.Statistical_tables_Entry_save.delete("0", 'end')
        self.Statistical_tables_Entry_save.insert(tk.INSERT, '统计-' + now_time)

    # 一次性选择多个订单列表
    def Statistical_tables_choose_files1(self):
        global Statistical_tables_add_list1
        file_paths = filedialog.askopenfilenames()
        if file_paths == '':
            return
        Statistical_tables_add_list1 = 8
        more_list = [
            [self.Statistical_tables_Label1_2, self.Statistical_tables_Entry1_2, self.Statistical_tables_Choose1_2],
            [self.Statistical_tables_Label1_3, self.Statistical_tables_Entry1_3, self.Statistical_tables_Choose1_3],
            [self.Statistical_tables_Label1_4, self.Statistical_tables_Entry1_4, self.Statistical_tables_Choose1_4],
            [self.Statistical_tables_Label1_5, self.Statistical_tables_Entry1_5, self.Statistical_tables_Choose1_5],
            [self.Statistical_tables_Label1_6, self.Statistical_tables_Entry1_6, self.Statistical_tables_Choose1_6],
            [self.Statistical_tables_Label1_7, self.Statistical_tables_Entry1_7, self.Statistical_tables_Choose1_7],
            [self.Statistical_tables_Label1_8, self.Statistical_tables_Entry1_8, self.Statistical_tables_Choose1_8],
            [self.Statistical_tables_Label1_9, self.Statistical_tables_Entry1_9, self.Statistical_tables_Choose1_9], ]
        more_list = iter(more_list)
        self.Statistical_tables_Entry1.delete("0", 'end')
        self.Statistical_tables_Entry1.insert(tk.INSERT, file_paths[0])
        now_list = next(more_list)
        list2 = [self.Statistical_tables_Entry1_2,
                 self.Statistical_tables_Entry1_3,
                 self.Statistical_tables_Entry1_4,
                 self.Statistical_tables_Entry1_5,
                 self.Statistical_tables_Entry1_6,
                 self.Statistical_tables_Entry1_7,
                 self.Statistical_tables_Entry1_8,
                 self.Statistical_tables_Entry1_9,
                 ]
        for file_path, entry in zip(file_paths[1:], list2):
            entry.delete("0", 'end')
            self.Statistical_tables_add1()
            entry.insert(tk.INSERT, file_path)
            now_list = next(more_list)
        while now_list is not None:
            for data in now_list:
                data.grid_forget()
            try:
                now_list = next(more_list)
            except StopIteration:
                break

    # 增加订单列表
    def Statistical_tables_add1(self):
        global Statistical_tables_add_list1
        i = Statistical_tables_add_list1
        if i == 8:
            self.Statistical_tables_Label1_2.grid(row=4, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_2)
            self.Statistical_tables_Entry1_2.delete("0", 'end')
            self.Statistical_tables_Entry1_2.grid(row=4, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_2)
            self.Statistical_tables_Choose1_2.grid(row=4, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_2)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 7:
            self.Statistical_tables_Label1_3.grid(row=5, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_3)
            self.Statistical_tables_Entry1_3.delete("0", 'end')
            self.Statistical_tables_Entry1_3.grid(row=5, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_3)
            self.Statistical_tables_Choose1_3.grid(row=5, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_3)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 6:
            self.Statistical_tables_Label1_4.grid(row=6, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_4)
            self.Statistical_tables_Entry1_4.delete("0", 'end')
            self.Statistical_tables_Entry1_4.grid(row=6, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_4)
            self.Statistical_tables_Choose1_4.grid(row=6, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_4)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 5:
            self.Statistical_tables_Label1_5.grid(row=7, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_5)
            self.Statistical_tables_Entry1_5.delete("0", 'end')
            self.Statistical_tables_Entry1_5.grid(row=7, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_5)
            self.Statistical_tables_Choose1_5.grid(row=7, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_5)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 4:
            self.Statistical_tables_Label1_6.grid(row=8, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_6)
            self.Statistical_tables_Entry1_6.delete("0", 'end')
            self.Statistical_tables_Entry1_6.grid(row=8, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_6)
            self.Statistical_tables_Choose1_6.grid(row=8, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_6)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 3:
            self.Statistical_tables_Label1_7.grid(row=9, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_7)
            self.Statistical_tables_Entry1_7.delete("0", 'end')
            self.Statistical_tables_Entry1_7.grid(row=9, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_7)
            self.Statistical_tables_Choose1_7.grid(row=9, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_7)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 2:
            self.Statistical_tables_Label1_8.grid(row=10, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_8)
            self.Statistical_tables_Entry1_8.delete("0", 'end')
            self.Statistical_tables_Entry1_8.grid(row=10, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_8)
            self.Statistical_tables_Choose1_8.grid(row=10, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_8)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1
        elif i == 1:
            self.Statistical_tables_Label1_9.grid(row=11, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label1_9)
            self.Statistical_tables_Entry1_9.delete("0", 'end')
            self.Statistical_tables_Entry1_9.grid(row=11, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry1_9)
            self.Statistical_tables_Choose1_9.grid(row=11, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose1_9)
            Statistical_tables_add_list1 = Statistical_tables_add_list1 - 1

    # 一次性选择多个pp订单表
    def Statistical_tables_choose_files(self):
        global Statistical_tables_add_list
        file_paths = filedialog.askopenfilenames()
        if file_paths == '':
            return
        Statistical_tables_add_list = 8
        more_list = [
            [self.Statistical_tables_Label_p2, self.Statistical_tables_Entry_p2, self.Statistical_tables_Choose_p2],
            [self.Statistical_tables_Label_p3, self.Statistical_tables_Entry_p3, self.Statistical_tables_Choose_p3],
            [self.Statistical_tables_Label_p4, self.Statistical_tables_Entry_p4, self.Statistical_tables_Choose_p4],
            [self.Statistical_tables_Label_p5, self.Statistical_tables_Entry_p5, self.Statistical_tables_Choose_p5],
            [self.Statistical_tables_Label_p6, self.Statistical_tables_Entry_p6, self.Statistical_tables_Choose_p6],
            [self.Statistical_tables_Label_p7, self.Statistical_tables_Entry_p7, self.Statistical_tables_Choose_p7],
            [self.Statistical_tables_Label_p8, self.Statistical_tables_Entry_p8, self.Statistical_tables_Choose_p8],
            [self.Statistical_tables_Label_p9, self.Statistical_tables_Entry_p9, self.Statistical_tables_Choose_p9], ]
        more_list = iter(more_list)
        self.Statistical_tables_Entry3.delete("0", 'end')
        self.Statistical_tables_Entry3.insert(tk.INSERT, file_paths[0])
        now_list = next(more_list)
        list2 = [self.Statistical_tables_Entry_p2,
                 self.Statistical_tables_Entry_p3,
                 self.Statistical_tables_Entry_p4,
                 self.Statistical_tables_Entry_p5,
                 self.Statistical_tables_Entry_p6,
                 self.Statistical_tables_Entry_p7,
                 self.Statistical_tables_Entry_p8,
                 self.Statistical_tables_Entry_p9,
                 ]
        for file_path, entry in zip(file_paths[1:], list2):
            entry.delete("0", 'end')
            self.Statistical_tables_add()
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

    # 增加pp订单表
    def Statistical_tables_add(self):
        global Statistical_tables_add_list
        i = Statistical_tables_add_list
        if i == 8:
            self.Statistical_tables_Label_p2.grid(row=15, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p2)
            self.Statistical_tables_Entry_p2.delete("0", 'end')
            self.Statistical_tables_Entry_p2.grid(row=15, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p2)
            self.Statistical_tables_Choose_p2.grid(row=15, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p2)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 7:
            self.Statistical_tables_Label_p3.grid(row=16, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p3)
            self.Statistical_tables_Entry_p3.delete("0", 'end')
            self.Statistical_tables_Entry_p3.grid(row=16, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p3)
            self.Statistical_tables_Choose_p3.grid(row=16, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p3)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 6:
            self.Statistical_tables_Label_p4.grid(row=17, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p4)
            self.Statistical_tables_Entry_p4.delete("0", 'end')
            self.Statistical_tables_Entry_p4.grid(row=17, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p4)
            self.Statistical_tables_Choose_p4.grid(row=17, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p4)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 5:
            self.Statistical_tables_Label_p5.grid(row=18, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p5)
            self.Statistical_tables_Entry_p5.delete("0", 'end')
            self.Statistical_tables_Entry_p5.grid(row=18, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p5)
            self.Statistical_tables_Choose_p5.grid(row=18, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p5)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 4:
            self.Statistical_tables_Label_p6.grid(row=19, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p6)
            self.Statistical_tables_Entry_p6.delete("0", 'end')
            self.Statistical_tables_Entry_p6.grid(row=19, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p6)
            self.Statistical_tables_Choose_p6.grid(row=19, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p6)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 3:
            self.Statistical_tables_Label_p7.grid(row=20, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p7)
            self.Statistical_tables_Entry_p7.delete("0", 'end')
            self.Statistical_tables_Entry_p7.grid(row=20, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p7)
            self.Statistical_tables_Choose_p7.grid(row=20, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p7)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 2:
            self.Statistical_tables_Label_p8.grid(row=21, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p8)
            self.Statistical_tables_Entry_p8.delete("0", 'end')
            self.Statistical_tables_Entry_p8.grid(row=21, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p8)
            self.Statistical_tables_Choose_p8.grid(row=21, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p8)
            Statistical_tables_add_list = Statistical_tables_add_list - 1
        elif i == 1:
            self.Statistical_tables_Label_p9.grid(row=22, column=0, sticky='e')
            all_use.tk_list.append(self.Statistical_tables_Label_p9)
            self.Statistical_tables_Entry_p9.delete("0", 'end')
            self.Statistical_tables_Entry_p9.grid(row=22, column=1, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Entry_p9)
            self.Statistical_tables_Choose_p9.grid(row=22, column=2, sticky='w')
            all_use.tk_list.append(self.Statistical_tables_Choose_p9)
            Statistical_tables_add_list = Statistical_tables_add_list - 1

    # 数据处理函数
    @creat_thread
    def work(self):
        df1 = all_use.read_excel_to_dataframe(
            [self.Statistical_tables_Entry1, self.Statistical_tables_Entry1_2, self.Statistical_tables_Entry1_3,
             self.Statistical_tables_Entry1_4, self.Statistical_tables_Entry1_5,
             self.Statistical_tables_Entry1_6, self.Statistical_tables_Entry1_7,
             self.Statistical_tables_Entry1_8, self.Statistical_tables_Entry1_9],
            ['订单列表1', '订单列表2', '订单列表3', '订单列表4', '订单列表5', '订单列表6',
             '订单列表7', '订单列表8', '订单列表9'],
            2,
            ['订单编号', '订单总价', '支付订单号（仅支持paypal 、stripe、worldpay、ingenico收款渠道）',
             '商品名称', '售出数量'],
            self.Statistical_tables_Entry_jg)
        if df1 is None:
            return
        df2 = all_use.read_excel_to_dataframe(self.Statistical_tables_Entry2, '运单明细表', 1,
                                              ['订单号', '计费重量(g)', '长(cm)', '宽(cm)', '高(cm)', '计费时间',
                                               '账单日期'],
                                              self.Statistical_tables_Entry_jg, rows_end=1)
        if df2 is None:
            return
        df3 = all_use.read_excel_to_dataframe(self.Statistical_tables_Entry_shangpinchengben, '商品成本表', 1,
                                              ['商品名称', '成本'],
                                              self.Statistical_tables_Entry_jg)
        if df3 is None:
            return
        df4 = all_use.read_excel_to_dataframe(
            [self.Statistical_tables_Entry3, self.Statistical_tables_Entry_p2, self.Statistical_tables_Entry_p3,
             self.Statistical_tables_Entry_p4, self.Statistical_tables_Entry_p5,
             self.Statistical_tables_Entry_p6, self.Statistical_tables_Entry_p7,
             self.Statistical_tables_Entry_p8, self.Statistical_tables_Entry_p9],
            ['pp订单号1', 'pp订单号2', 'pp订单号3', 'pp订单号4', 'pp订单号5', 'pp订单号6',
             'pp订单号7', 'pp订单号8', 'pp订单号9'], 1,
            ['描述', '净额', '参考交易号'],
            self.Statistical_tables_Entry_jg)
        if df4 is None:
            return

        # 统计订单列表的售出总数
        df1 = df1.cast({"售出数量": pl.Int32, '订单总价': pl.Float64}, strict=False)
        sum = df1.group_by('订单编号').agg(pl.col('售出数量').sum()).rename({"售出数量": "总售出数量"})
        df_end = df1.join(sum, on='订单编号', how='left')

        # 连接运单明细表
        df2 = df2.filter(pl.col('费用类型') != '应收赔偿')
        df2 = df2.with_columns(pl.when(True).then(pl.lit('是')).alias('是否发货'))
        df_end = df_end.join(df2, left_on='订单编号', right_on='订单号', how='left')
        df_end = df_end.with_columns(pl.col('是否发货').fill_null('否'))
        # 连接商品成本表
        # 拼接成本表的商品名称和款式一
        df3 = df3.with_columns((pl.col('商品名称')
                                + pl.when(pl.col('款式1').is_not_null()).then(' ' + pl.col('款式1')).otherwise(
                    pl.col('款式1'))).alias('商品名称加款式'))
        df3 = df3.with_columns(pl.col('商品名称加款式').str.strip_chars_end(' '))
        # 拼接处理表的商品名称和款式一
        df_end = df_end.with_columns((pl.col('商品名称')
                                      + pl.when(pl.col('款式1').is_not_null()).then(' ' + pl.col('款式1')).otherwise(
                    pl.col('款式1'))
                                      + pl.when(pl.col('款式2').is_not_null()).then(' ' + pl.col('款式2')).otherwise(
                    pl.col('款式2'))
                                      + pl.when(pl.col('款式3').is_not_null()).then(' ' + pl.col('款式3')).otherwise(
                    pl.col('款式3'))).alias('商品名称加款式'))
        df_end = df_end.with_columns(pl.col('商品名称加款式').str.strip_chars_end(' '))

        df_end = df_end.join(df3.unique(subset='商品名称加款式', keep='first'), left_on='商品名称',
                             right_on='商品名称加款式',
                             how='left')
        mor_chenben = df3.filter(pl.col('商品名称') == '默认成本')['成本'][0]
        df_end = df_end.with_columns(pl.col('成本').fill_null(mor_chenben))
        df_end = df_end.cast({"售出数量": pl.Int32, "成本": pl.Float64}, strict=False)
        df_end = df_end.with_columns((pl.col('售出数量') * pl.col('成本')).alias('各个商品成本'))

        df_end = df_end.join(
            df_end.select(['订单编号', '各个商品成本']).group_by("订单编号", maintain_order=True).sum().rename(
                {'各个商品成本': '商品成本'}).select(
                ['订单编号', '商品成本']), on='订单编号', how='left')

        # 连接pp订单表
        df4 = df4.filter(pl.col('参考交易号') != '')
        df4 = df4.with_columns(
            pl.when(pl.col('描述').is_in(['付款退款', '退单', '争议费'])).then(pl.col('总额')).otherwise(0).alias(
                '退款金额'))
        df_end1 = df_end.join(df4, left_on='支付订单号（仅支持paypal 、stripe、worldpay、ingenico收款渠道）',
                              right_on='参考交易号', how='left')
        df_end = df_end.join(df_end1.group_by("订单编号").sum().select(['订单编号', '退款金额']), on='订单编号',
                             how='left')
        df_end = df_end.unique(subset='订单编号', keep='first')

        df_end = df_end.select(
            ['订单编号', '订单创建时间', '总售出数量', '计费重量(g)', '长(cm)', '宽(cm)', '高(cm)', '计费时间',
             '账单日期',
             '订单总价', '商品成本', '账单金额(元)', '退款金额', '是否发货'])

        df_end = df_end.rename({'订单编号': '订单号', '总售出数量': '商品数量', '账单金额(元)': '￥运费'})
        # 先排未发货的有退款的，再到未发货无退款的。接着到发货有退款的，后面其他正常的
        df_end = df_end.with_columns(pl.when(pl.col('是否发货') == '否').then(1).otherwise(0).alias('是否发货顺序'),
                                     pl.when(pl.col('退款金额') != 0).then(1).otherwise(0).alias('退款金额顺序'))
        df_end = df_end.sort(by=['是否发货', '退款金额', '订单号'], nulls_last=[True, True, True])  # 默认降排序
        df_end = df_end.select(
            ['订单号', '订单创建时间', '商品数量', '计费重量(g)', '长(cm)', '宽(cm)', '高(cm)', '计费时间', '账单日期',
             '订单总价', '商品成本', '￥运费', '退款金额', '是否发货'])
        full_path = os.path.join(os.path.dirname(self.Statistical_tables_Entry1.get()),
                                 self.Statistical_tables_Entry_save.get() + '.xlsx')

        if not all_use.polarsdf_openpyxl_to_excel(df_end, full_path):
            self.Statistical_tables_Entry_jg.delete("0", 'end')
            self.Statistical_tables_Entry_jg.insert(tk.INSERT, '文件保存失败')
            self.Statistical_tables_Entry_jg.configure(fg='red')
            return

        q = str(df_end.shape[0])  # 获取行数
        self.Statistical_tables_Entry_jg.delete("0", 'end')
        self.Statistical_tables_Entry_jg.insert(tk.INSERT,
                                                '共统计了' + q + '个信息' + '，已保存到文件：' + full_path)
        self.Statistical_tables_Entry_jg.configure(fg='green')
