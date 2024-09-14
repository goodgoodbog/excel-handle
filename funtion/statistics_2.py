# -*- coding=utf-8-*-
# 功能9：统计各商品数量
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
import python_calamine
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use
import polars as pl

class statistics_2:
    def __init__(self, root):
        use = all_use.Use(root)
        self.ten_Label_tip = tk.Label(root, text='功能9：统计各商品数量')
        self.ten_Label_file = tk.Label(root, text='对照表:')
        self.ten_Button_Choose_file1 = ttk.Button(root, text='选择文件',
                                                  command=lambda: choose_flie(self.ten_Entry_file1))
        self.ten_Entry_file1 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                        name='ten_Entry_file1', validate="key", relief='sunken',
                                        validatecommand=(use.check_open_file1, '%P', '%W'))
        self.ten_Button_Choose_work = ttk.Button(root, text='开始统计', command=self.work)
        self.ten_Label_save = tk.Label(root, text='保存到:')
        self.ten_Entry_save = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.ten_Label_jg = tk.Label(root, text='结果提示:')
        self.ten_Entry_jg = tk.Entry(root, width=56)

    # 布局函数
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.ten_Label_tip.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.ten_Label_tip)
        self.ten_Label_file.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.ten_Label_file)
        self.ten_Entry_file1.delete("0", 'end')
        self.ten_Entry_file1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.ten_Entry_file1)
        self.ten_Button_Choose_file1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.ten_Button_Choose_file1)

        self.ten_Label_save.grid(row=5, column=0, sticky='e')
        all_use.tk_list.append(self.ten_Label_save)
        self.ten_Entry_save.delete("0", 'end')
        self.ten_Entry_save.grid(row=5, column=1, sticky='w')
        all_use.tk_list.append(self.ten_Entry_save)
        self.ten_Button_Choose_work.grid(row=6, column=1)
        all_use.tk_list.append(self.ten_Button_Choose_work)
        self.ten_Label_jg.grid(row=7, column=0, sticky='e')
        all_use.tk_list.append(self.ten_Label_jg)
        self.ten_Entry_jg.delete("0", 'end')
        self.ten_Entry_jg.grid(row=7, column=1, sticky='w')
        all_use.tk_list.append(self.ten_Entry_jg)
        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.ten_Entry_save.delete("0", 'end')
        self.ten_Entry_save.insert(tk.INSERT, '各商品数量统计表-' + now_time)

    # 数据处理函数
    @creat_thread
    def work(self):  # 已polars
        try:
            rows = list(
                python_calamine.CalamineWorkbook.from_path(self.ten_Entry_file1.get()).get_sheet_by_index(
                    0).to_python())
        except:
            self.ten_Entry_jg.delete("0", 'end')
            self.ten_Entry_jg.insert(tk.INSERT, '对照表出错,请重新选择')
            self.ten_Entry_jg.configure(fg='red')
            return
        dict1 = {}
        for every_row in rows:
            # 去掉每行数据的前三列
            for value in every_row[3:]:
                if value in dict1:
                    dict1[value] += 1
                else:
                    dict1[value] = 1
        del dict1['']
        keys, values = [x for x in dict1.keys() if x != ''], [x for x in dict1.values()]

        df = pl.DataFrame(data=[keys, values], schema=['商品名称', '数量'])
        df = df.sort('商品名称')

        full_path = os.path.join(os.path.dirname(self.ten_Entry_file1.get()),
                                 self.ten_Entry_save.get() + '.xlsx')
        if not all_use.polarsdf_openpyxl_to_excel(df, full_path):
            return

        self.ten_Entry_jg.delete("0", 'end')
        self.ten_Entry_jg.insert(tk.INSERT, '共有{}种商品,保存到{}中'.format(len(dict1), self.ten_Entry_save.get()))
        self.ten_Entry_jg.configure(fg='green')
