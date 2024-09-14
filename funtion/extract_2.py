# -*- coding=utf-8-*-
# 功能10：提取相关信息
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
import polars as pl
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use


class Extract_2:
    def __init__(self, root):
        use = all_use.Use(root)
        self.eleven_Label1 = tk.Label(root, text='商户汇总账单：')
        self.eleven_Label2 = tk.Label(root, text='前两个字符：')
        self.eleven_Label_save = tk.Label(root, text='保存到:')
        self.eleven_Label_jg = tk.Label(root, text='结果提示:')
        self.eleven_Labeltip = tk.Label(root, text='功能10：提取相关信息')
        self.eleven_Choose1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.eleven_Entry1))
        self.eleven_Choose_work = ttk.Button(root, text='开始提取', command=self.work)
        self.eleven_Entry1 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                      name='eleven_Entry1',
                                      validate="key", relief='sunken',
                                      validatecommand=(use.check_open_file1, '%P', '%W'))
        self.eleven_Entry2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                      name='eleven_Entry2',
                                      relief='sunken')
        self.eleven_Entry_save = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.eleven_Entry_jg = tk.Entry(root, width=56)

    # 布局函数
    # @grid_progressbar
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.eleven_Labeltip.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.eleven_Labeltip)
        self.eleven_Label1.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.eleven_Label1)
        self.eleven_Entry1.delete("0", 'end')
        self.eleven_Entry1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.eleven_Entry1)
        self.eleven_Choose1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.eleven_Choose1)
        self.eleven_Label2.grid(row=4, column=0, sticky='e')
        all_use.tk_list.append(self.eleven_Label2)
        self.eleven_Entry2.delete("0", 'end')
        self.eleven_Entry2.grid(row=4, column=1, sticky='w')
        all_use.tk_list.append(self.eleven_Entry2)

        self.eleven_Label_save.grid(row=5, column=0, sticky='e')
        all_use.tk_list.append(self.eleven_Label_save)
        self.eleven_Entry_save.delete("0", 'end')
        self.eleven_Entry_save.grid(row=5, column=1, sticky='w')
        all_use.tk_list.append(self.eleven_Entry_save)
        self.eleven_Choose_work.grid(row=6, column=1)
        all_use.tk_list.append(self.eleven_Choose_work)
        self.eleven_Label_jg.grid(row=7, column=0, sticky='e')
        all_use.tk_list.append(self.eleven_Label_jg)
        self.eleven_Entry_jg.delete("0", 'end')
        self.eleven_Entry_jg.grid(row=7, column=1, sticky='w')
        all_use.tk_list.append(self.eleven_Entry_jg)
        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.eleven_Entry_save.delete("0", 'end')
        self.eleven_Entry_save.insert(tk.INSERT, '提取-' + now_time)

    # 数据处理函数
    @creat_thread
    def work(self):  # polars
        df1 = all_use.read_excel_to_dataframe(self.eleven_Entry1, '商户汇总账单', 1, ['订单号'], self.eleven_Entry_jg)

        if df1 is None:
            return
        need_string = self.eleven_Entry2.get()
        df_end = df1.filter(pl.col('订单号').str.starts_with(need_string))

        sss = self.eleven_Entry_save.get()
        full_path = os.path.join(os.path.dirname(self.eleven_Entry1.get()), sss + '.xlsx')

        if not all_use.polarsdf_openpyxl_to_excel(df_end, full_path):
            self.eleven_Entry_jg.delete("0", 'end')
            self.eleven_Entry_jg.insert(tk.INSERT, '文件保存失败')
            self.eleven_Entry_jg.configure(fg='red')
            return
        self.eleven_Entry_jg.delete("0", 'end')
        self.eleven_Entry_jg.insert(tk.INSERT, '找到{}个目标，已保存到{}.xlsx中'.format(len(df_end), sss))
        self.eleven_Entry_jg.configure(fg='green')