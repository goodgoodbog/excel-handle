# -*- coding=utf-8-*-
# 功能1：订单列表信息提取
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import numpy as np
import polars as pl
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use


class add_file:
    def __init__(self, root):
        self.a_list = 0
        self.use = all_use.Use(root)

        self.a1 = tk.Label(root, text='功能1：提取信息')
        self.Extract_compare_Entry2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.a2 = tk.Label(root, text='订单列表1:')
        self.a3 = tk.Label(root, text='订单列表2:')
        self.a4 = tk.Label(root, text='订单列表3:')
        self.a5 = tk.Label(root, text='订单列表4:')
        self.a6 = tk.Label(root, text='订单列表5:')
        self.a7 = tk.Label(root, text='订单列表6:')
        self.a8 = tk.Label(root, text='订单列表7:')
        self.a9 = tk.Label(root, text='订单列表8:')
        self.a10 = tk.Label(root, text='订单列表9:')
        self.a11 = tk.Label(root, text='保存到:')
        self.jg = tk.Label(root, text='结果提示：')
        self.ABtton = ttk.Button(root, text='订单列表数+1', command=self.a_add)
        self.one_Button_choose_files = ttk.Button(root, text='选择多个表', command=self.one_fun_choose_files)
        self.aBtton = ttk.Button(root, text='开始提取信息', command=self.work_1)
        self.choose1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e1))
        self.choose2 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e2))
        self.choose3 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e3))
        self.choose4 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e4))
        self.choose5 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e5))
        self.choose6 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e6))
        self.choose7 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e7))
        self.choose8 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e8))
        self.choose9 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.e9))
        # 9个文本框
        self.e1 = tk.Entry(root, name='e1', width=56, highlightthickness=1.5, highlightcolor='blue', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e2', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e3 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e3', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e4 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e4', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e5 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e5', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e6 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e6', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e7 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e7', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e8 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e8', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e9 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue', name='e9', validate="key",
                           relief='sunken',
                           validatecommand=(self.use.check_open_file1, '%P', '%W'))
        self.e10 = tk.Entry(root, width=56)
        self.e11 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')  # 保存文件名

    # 布局函数
    def show(self):
        while len(all_use.tk_list):
            t = all_use.tk_list.pop()
            if t.widgetName == 'entry':
                t.delete("0", 'end')
            t.grid_forget()
        self.a1.grid(row=2, column=0, sticky='w', columnspan=2)
        all_use.tk_list.append(self.a1)

        self.a2.grid(row=3, column=0, sticky='e')
        all_use.tk_list.append(self.a2)
        self.e1.delete("0", 'end')
        self.e1.grid(row=3, column=1, sticky='w')
        all_use.tk_list.append(self.e1)
        self.choose1.grid(row=3, column=2, sticky='w')
        all_use.tk_list.append(self.choose1)
        self.one_Button_choose_files.grid(row=3, column=3, sticky='w')
        all_use.tk_list.append(self.one_Button_choose_files)
        self.ABtton.grid(row=3, column=4, sticky='w')
        all_use.tk_list.append(self.ABtton)
        self.aBtton.grid(row=12, column=1, sticky='')
        all_use.tk_list.append(self.aBtton)
        self.e11.delete("0", 'end')
        self.e11.grid(row=13, column=1, sticky='w')
        all_use.tk_list.append(self.e11)
        self.a11.grid(row=13, column=0, sticky='e')
        all_use.tk_list.append(self.a11)
        self.jg.grid(row=14, column=0, sticky='e')
        all_use.tk_list.append(self.jg)
        self.e10.delete("0", 'end')
        self.e10.grid(row=14, column=1, sticky='w')
        all_use.tk_list.append(self.e10)
        self.a_list = 8

        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.e11.delete("0", 'end')
        self.e11.insert(tk.INSERT, '订单合并表-' + now_time)

    # 文件数加1
    def a_add(self):
        i = self.a_list
        if i == 8:
            self.a3.grid(row=4, column=0, sticky='e')
            all_use.tk_list.append(self.a3)
            self.e2.delete("0", 'end')
            self.e2.grid(row=4, column=1, sticky='w')
            all_use.tk_list.append(self.e2)
            self.choose2.grid(row=4, column=2, sticky='w')
            all_use.tk_list.append(self.choose2)
            self.a_list = self.a_list - 1
        elif i == 7:
            self.a4.grid(row=5, column=0, sticky='e')
            all_use.tk_list.append(self.a4)
            self.e3.delete("0", 'end')
            self.e3.grid(row=5, column=1, sticky='w')
            all_use.tk_list.append(self.e3)
            self.choose3.grid(row=5, column=2, sticky='w')
            all_use.tk_list.append(self.choose3)
            self.a_list = self.a_list - 1
        elif i == 6:
            self.a5.grid(row=6, column=0, sticky='e')
            all_use.tk_list.append(self.a5)
            self.e4.delete("0", 'end')
            self.e4.grid(row=6, column=1, sticky='w')
            all_use.tk_list.append(self.e4)
            self.choose4.grid(row=6, column=2, sticky='w')
            all_use.tk_list.append(self.choose4)
            self.a_list = self.a_list - 1
        elif i == 5:
            self.a6.grid(row=7, column=0, sticky='e')
            all_use.tk_list.append(self.a6)
            self.e5.delete("0", 'end')
            self.e5.grid(row=7, column=1, sticky='w')
            all_use.tk_list.append(self.e5)
            self.choose5.grid(row=7, column=2, sticky='w')
            all_use.tk_list.append(self.choose5)
            self.a_list = self.a_list - 1
        elif i == 4:
            self.a7.grid(row=8, column=0, sticky='e')
            all_use.tk_list.append(self.a7)
            self.e6.delete("0", 'end')
            self.e6.grid(row=8, column=1, sticky='w')
            all_use.tk_list.append(self.e6)
            self.choose6.grid(row=8, column=2, sticky='w')
            all_use.tk_list.append(self.choose6)
            self.a_list = self.a_list - 1
        elif i == 3:
            self.a8.grid(row=9, column=0, sticky='e')
            all_use.tk_list.append(self.a8)
            self.e7.delete("0", 'end')
            self.e7.grid(row=9, column=1, sticky='w')
            all_use.tk_list.append(self.e7)
            self.choose7.grid(row=9, column=2, sticky='w')
            all_use.tk_list.append(self.choose7)
            self.a_list = self.a_list - 1
        elif i == 2:
            self.a9.grid(row=10, column=0, sticky='e')
            all_use.tk_list.append(self.a9)
            self.e8.delete("0", 'end')
            self.e8.grid(row=10, column=1, sticky='w')
            all_use.tk_list.append(self.e8)
            self.choose8.grid(row=10, column=2, sticky='w')
            all_use.tk_list.append(self.choose8)
            self.a_list = self.a_list - 1
        elif i == 1:
            self.a10.grid(row=11, column=0, sticky='e')
            all_use.tk_list.append(self.a10)
            self.e9.delete("0", 'end')
            self.e9.grid(row=11, column=1, sticky='w')
            all_use.tk_list.append(self.e9)
            self.choose9.grid(row=11, column=2, sticky='w')
            all_use.tk_list.append(self.choose9)
            self.a_list = self.a_list - 1

    # 一次性选择多个表
    def one_fun_choose_files(self):
        file_paths = filedialog.askopenfilenames()
        if file_paths == '':
            return
        self.a_list = 8
        more_list = [[self.a3, self.e2, self.choose2],
                     [self.a4, self.e3, self.choose3],
                     [self.a5, self.e4, self.choose4],
                     [self.a6, self.e5, self.choose5],
                     [self.a7, self.e6, self.choose6],
                     [self.a8, self.e7, self.choose7],
                     [self.a9, self.e8, self.choose8],
                     [self.a10, self.e9, self.choose9]
                     ]

        more_list = iter(more_list)
        self.e1.delete("0", 'end')
        self.e1.insert(tk.INSERT, file_paths[0])
        now_list = next(more_list)
        list2 = [self.e2,
                 self.e3,
                 self.e4,
                 self.e5,
                 self.e6,
                 self.e7,
                 self.e8,
                 self.e9,
                 ]

        for file_path, entry in zip(file_paths[1:], list2):
            entry.delete("0", 'end')
            self.a_add()
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

    # 数据处理函数
    @creat_thread
    def work_1(self):  # 已polars
        df1 = all_use.read_excel_to_dataframe(entry_list=[self.e1, self.e2, self.e3,
                                                          self.e4, self.e5,
                                                          self.e6, self.e7,
                                                          self.e8, self.e9],
                                              entry_name_list=['订单列表1', '订单列表2', '订单列表3', '订单列表4',
                                                               '订单列表5', '订单列表6',
                                                               '订单列表7', '订单列表8', '订单列表9'], title_began=2,
                                              title_need_list=['订单编号', '名', '姓', '国家/地区', '街道', '寓所',
                                                               '城市', '省/州', '邮编',
                                                               '联系电话', '联系邮箱'],
                                              jg_entry=self.e10)

        if df1 is None:
            return
        df1 = df1.unique('订单编号', keep='first')  # 去重
        # 最终dataframe
        last_column = {'订单号', '交货仓', '产品名称', '收件人姓名', '收件人电话', '收件人邮箱', '收件人税号',
                       '收件人公司',
                       '收件人国家', '收件人省/州', '收件人城市', '收件人邮编', '收件人地址', '收件人门牌号',
                       '寄件人税号信息', '包装尺寸【长】cm', '包装尺寸【宽】cm', '包装尺寸【高】cm', '收款到账日期',
                       '币种类型',
                       '是否含电', '拣货单信息', 'IOSS税号', '中文品名1', '英文品名1', '单票数量1', '重量1(g)',
                       '申报价值1',
                       '商品材质1', '商品海关编码1', '商品链接1', '中文品名2', '英文品名2', '单票数量2', '重量2(g)',
                       '申报价值2', '商品材质2', '商品海关编码2', '商品链接2', '中文品名3', '英文品名3', '单票数量3',
                       '重量3(g)', '申报价值3', '商品材质3', '商品海关编码3', '商品链接3', '中文品名4', '英文品名4',
                       '单票数量4', '重量4(g)', '申报价值4', '商品材质4', '商品海关编码4', '商品链接4', '中文品名5',
                       '英文品名5', '单票数量5', '重量5(g)', '申报价值5', '商品材质5', '商品海关编码5', '商品链接5'}
        last_column1 = ['订单号', '平台交易号', '交货仓', '产品名称', '收件人姓名', '收件人电话', '收件人邮箱',
                        '收件人税号',
                        '收件人公司',
                        '收件人国家', '收件人省/州', '收件人城市', '收件人邮编', '收件人地址', '收件人门牌号',
                        '寄件人税号信息', '包装尺寸【长】cm', '包装尺寸【宽】cm', '包装尺寸【高】cm', '收款到账日期',
                        '币种类型',
                        '是否含电', '拣货单信息', 'IOSS税号', '中文品名1', '英文品名1', '单票数量1', '重量1(g)',
                        '申报价值1',
                        '商品材质1', '商品海关编码1', '商品链接1', '中文品名2', '英文品名2', '单票数量2', '重量2(g)',
                        '申报价值2', '商品材质2', '商品海关编码2', '商品链接2', '中文品名3', '英文品名3', '单票数量3',
                        '重量3(g)', '申报价值3', '商品材质3', '商品海关编码3', '商品链接3', '中文品名4', '英文品名4',
                        '单票数量4', '重量4(g)', '申报价值4', '商品材质4', '商品海关编码4', '商品链接4', '中文品名5',
                        '英文品名5', '单票数量5', '重量5(g)', '申报价值5', '商品材质5', '商品海关编码5', '商品链接5']
        df1 = df1.cast({"街道": pl.String, '寓所': pl.String})
        df1 = df1.with_columns(
            pl.lit('燕文专线追踪-普货').alias('产品名称'),
            (pl.col('名') + ' ' + pl.col('姓')).alias('收件人姓名'),
            pl.when(pl.col('联系电话') != '').then(pl.col('联系电话').str.replace_all(r'[+ ]', '')).otherwise(
                pl.lit('000000000')).alias('收件人电话'),
            (pl.col('街道') + ' ' + pl.col('寓所')).alias('收件人地址'),
            pl.col('城市').str.replace_all(r'[^\w\s]', '').alias('收件人城市'),
            pl.col('联系邮箱').alias('收件人邮箱'),
            pl.col('订单编号').alias('订单号'),
            pl.col('国家/地区').alias('收件人国家'),
            pl.col('省/州').alias('收件人省/州'),
            pl.col('邮编').alias('收件人邮编'),
            pl.lit('美元').alias('币种类型'),
            pl.lit('否').alias('是否含电'),
            pl.lit('骰子').alias('中文品名1'),
            pl.lit('dice').alias('英文品名1'),
            pl.lit(1).alias('单票数量1'),
            pl.lit(10).alias('重量1(g)'),
            pl.when(pl.col('国家/地区').is_in(['美国', '澳大利亚', 'United States', 'Australia'])).then(30).when(
                pl.col('国家/地区').is_in(['加拿大', 'Canada'])).then(13).otherwise(5).alias('申报价值1'),
            pl.lit(np.nan).alias('平台交易号')
        )
        need_column = last_column.difference(set(df1.columns))
        for column1 in need_column:
            df1 = df1.with_columns(pl.lit(None).alias(column1))
        df_end = df1.select(last_column1)
        df_end = df_end.sort(by='订单号')

        # 使用openpyxl写入excel比dataframe.to_excel()快一点
        # wb.save(r'D:\Users\17219\Desktop\00.xlsx')
        full_path = os.path.join(os.path.dirname(self.e1.get()), self.e11.get() + '.xlsx')

        if not all_use.polarsdf_openpyxl_to_excel(df_end, full_path):
            self.e10.delete("0", 'end')
            self.e10.insert(tk.INSERT, '文件保存失败')
            self.e10.configure(fg='red')
            return
        q = str(df_end.shape[0])  # 获取行数
        self.e10.delete("0", 'end')
        self.e10.insert(tk.INSERT, '共提取了' + q + '个信息' + '，已保存到文件：' + self.e11.get() + '.xlsx中')
        self.e10.configure(fg='green')
