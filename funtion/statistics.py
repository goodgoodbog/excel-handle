# -*- coding=utf-8-*-
# @Time : 2024/9/10 11:39
# @Author : BQ
# @File : statistics.py
# @Software: PyCharm
#功能6：统计商品数量
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
import python_calamine
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use
import polars as pl
class statistics:
    def __init__(self,root):
        self.use=all_use.Use(root)
        self.seven_Label1 = tk.Label(root, text='采购信息表:')
        self.seven_Label2 = tk.Label(root, text='发货信息表:')
        self.seven_Label_save = tk.Label(root, text='保存到:')
        self.seven_Label_jg = tk.Label(root, text='结果提示:')
        self.seven_Labeltip = tk.Label(root, text='功能6：统计商品数量')
        self.seven_Choose1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.seven_Entry1))
        self.seven_Choose2 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.seven_Entry2))
        self.seven_Choose_work = ttk.Button(root, text='开始统计', command=self.work)
        self.seven_Entry1 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='seven_Entry1',
                                validate="key", relief='sunken',
                                validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.seven_Entry2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='seven_Entry2',
                                validate="key", relief='sunken',
                                validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.seven_Entry_save = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.seven_Entry_jg = tk.Entry(root, width=56)

    # 布局函数
    # @grid_progressbar
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.seven_Labeltip.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.seven_Labeltip)
        self.seven_Label1.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.seven_Label1)
        self.seven_Entry1.delete("0", 'end')
        self.seven_Entry1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.seven_Entry1)
        self.seven_Choose1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.seven_Choose1)
        self.seven_Label2.grid(row=4, column=0, sticky='e')
        all_use.tk_list.append(self.seven_Label2)
        self.seven_Entry2.delete("0", 'end')
        self.seven_Entry2.grid(row=4, column=1, sticky='w')
        all_use.tk_list.append(self.seven_Entry2)
        self.seven_Choose2.grid(row=4, column=2, sticky='w')
        all_use.tk_list.append(self.seven_Choose2)

        self.seven_Label_save.grid(row=5, column=0, sticky='e')
        all_use.tk_list.append(self.seven_Label_save)
        self.seven_Entry_save.delete("0", 'end')
        self.seven_Entry_save.grid(row=5, column=1, sticky='w')
        all_use.tk_list.append(self.seven_Entry_save)
        self.seven_Choose_work.grid(row=6, column=1)
        all_use.tk_list.append(self.seven_Choose_work)
        self.seven_Label_jg.grid(row=7, column=0, sticky='e')
        all_use.tk_list.append(self.seven_Label_jg)
        self.seven_Entry_jg.delete("0", 'end')
        self.seven_Entry_jg.grid(row=7, column=1, sticky='w')
        all_use.tk_list.append(self.seven_Entry_jg)
        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.seven_Entry_save.delete("0", 'end')
        self.seven_Entry_save.insert(tk.INSERT, '统计商品数量表-' + now_time)

    # 数据处理函数
    @creat_thread
    def work(self):  # polars
        dict1 = {}
        dict2 = {}
        try:
            rows1 = list(
                python_calamine.CalamineWorkbook.from_path(self.seven_Entry1.get()).get_sheet_by_index(0).to_python())
        except:
            self.seven_Entry1.configure(fg='red')
            self.seven_Entry_jg.delete("0", 'end')
            self.seven_Entry_jg.insert(tk.INSERT, '采购信息表出错,请重新选择')
            self.seven_Entry_jg.configure(fg='red')
            return
        try:
            rows2 = list(
                python_calamine.CalamineWorkbook.from_path(self.seven_Entry2.get()).get_sheet_by_index(0).to_python())
        except:
            self.seven_Entry2.configure(fg='red')
            self.seven_Entry_jg.delete("0", 'end')
            self.seven_Entry_jg.insert(tk.INSERT, '发货信息表出错,请重新选择')
            self.seven_Entry_jg.configure(fg='red')
            return
        for every_row in rows1:
            for value in every_row:
                if value in dict1:
                    dict1[value] += 1
                else:
                    dict1[value] = 1
        for every_row in rows2:
            for value in every_row:
                if value in dict2:
                    dict2[value] += 1
                else:
                    dict2[value] = 1
        del dict1['']
        del dict2['']
        df1 = pl.DataFrame(data=[[x for x in dict1.keys()], [x for x in dict1.values()]],
                           schema=['商品名称', '采购数量'])
        df2 = pl.DataFrame(data=[[x for x in dict2.keys()], [x for x in dict2.values()]],
                           schema=['商品名称', '发货数量'])
        df_end = df1.join(df2, on='商品名称', how='full')
        df_end = df_end.with_columns(
            pl.when(pl.col('商品名称').is_null()).then(pl.col('商品名称_right')).otherwise(pl.col('商品名称')).alias(
                '商品名称'))
        df_end = df_end.fill_null(0)
        df_end = df_end.with_columns((pl.col('采购数量') - pl.col('发货数量')).alias('剩余数量'))
        df_end = df_end.sort('商品名称').select(['商品名称', '采购数量', '发货数量', '剩余数量'])

        full_path = os.path.join(os.path.dirname(self.seven_Entry1.get()),
                                 self.seven_Entry_save.get() + '.xlsx')
        if not all_use.polarsdf_openpyxl_to_excel(df_end, full_path):
            self.seven_Entry_jg.delete("0", 'end')
            self.seven_Entry_jg.insert(tk.INSERT, '文件保存失败')
            self.seven_Entry_jg.configure(fg='red')
            return
        self.seven_Entry_jg.delete("0", 'end')
        self.seven_Entry_jg.insert(tk.INSERT, '共统计了{}种商品,保存到{}中'.format(len(df_end), self.seven_Entry_save.get()))
        self.seven_Entry_jg.configure(fg='green')

