# -*- coding=utf-8-*-
# 功能2：根据发货单号整理制单信息(末尾有未匹配出的制单信息)
from tkinter import ttk
import tkinter as tk
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use
import polars as pl


class match:
    def __init__(self, root):
        self.use = all_use.Use(root)
        self.c1 = tk.Label(root, text='功能2：整理信息：')
        self.c2 = tk.Label(root, text='要发货单号:')
        self.c3 = tk.Label(root, text='功能1生成的文件:')
        self.e1 = tk.Entry(root, name='match_e1', width=56, highlightthickness=1.5, highlightcolor='blue',
                           validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.choose1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e1))
        self.e2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='match_e2',
                           validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.choose2 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e2))
        self.cBtton = ttk.Button(root, text='开始整理信息', command=self.work)
        self.jg = tk.Label(root, text='结果提示：')
        self.e10 = tk.Entry(root, width=56)

    # 布局函数
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.c1.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.c1)
        self.c2.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.c2)
        self.e1.delete("0", 'end')
        self.e1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.e1)
        self.choose1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.choose1)
        self.c3.grid(row=4, column=0, sticky='e')
        all_use.tk_list.append(self.c3)
        self.e2.delete("0", 'end')
        self.e2.grid(row=4, column=1, sticky='w')
        all_use.tk_list.append(self.e2)
        self.choose2.grid(row=4, column=2, sticky='w')
        all_use.tk_list.append(self.choose2)

        self.cBtton.grid(row=5, column=1)
        all_use.tk_list.append(self.cBtton)
        self.jg.grid(row=6, column=0, sticky='e')
        all_use.tk_list.append(self.jg)
        self.e10.delete("0", 'end')
        self.e10.grid(row=6, column=1, sticky='w')
        all_use.tk_list.append(self.e10)

    # 数据处理函数
    # 注意：最后利用了’产品名称‘一列来排序，若’产品名称‘一列改变须更改代码
    @creat_thread
    def work(self):  # polars
        df1 = all_use.read_excel_to_dataframe(self.e1, '要发货的单号', 0, [], self.e10)
        if df1 is None:
            return
        df2 = all_use.read_excel_to_dataframe(self.e2, '功能一生成的文件', 1, ['订单号'], self.e10)
        if df2 is None:
            return
        df1 = df1.filter(pl.col('column_0') != '')
        df1 = df1.select(['column_0'])
        df1 = df1.unique('column_0')
        df1 = df1.rename({'column_0': '订单号'})
        df_end = df1.join(df2.unique('订单号'), on='订单号', how='left')
        df_end = df_end.sort(by=['产品名称', '订单号'], nulls_last=[True, True])
        if not all_use.polarsdf_openpyxl_to_excel(df_end, self.e2.get()):
            self.e10.delete("0", 'end')
            self.e10.insert(tk.INSERT, '保存失败！！！')
            self.e10.configure(fg='red')
            return
        self.e10.delete("0", 'end')
        self.e10.insert(tk.INSERT, '匹配完成，已保存到文件:' + self.e2.get() + '中')
        self.e10.configure(fg='green')