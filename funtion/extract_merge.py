# -*- coding=utf-8-*-
# 功能1+4：提取并对照信息
# @Time : 2024/9/14 15:33
# @Author : BQ
# @File : extract_merge.py
# @Software: PyCharm
import os
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import python_calamine
import numpy as np
import polars as pl
from common.all_use import choose_flie
from common.all_use import creat_thread
from common import all_use


class Extract_Merge:
    def __init__(self, root):
        use = all_use.Use(root)
        self.f1 = tk.Label(root, text='功能1+4：提取并合并信息：')
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
        self.f_jg_label = tk.Label(root, text='结果提示:')
        self.Extract_compare_Label1 = tk.Label(root, text='订单合并表保存到：')
        self.Extract_compare_Entry1 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.Extract_compare_Label2 = tk.Label(root, text='对照表保存到：')
        self.Extract_compare_Entry2 = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue')
        self.Extract_compare_Choose_work = ttk.Button(root, text='开始合并对照', command=self.work)
        self.Extract_compare_Choose_more_file = ttk.Button(root, text='选择多个订单列表',
                                                           command=self.Extract_compare_Choose_files)
        self.f_duizhao_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                        name='extract_merge_duizhao_entry', validate="key", relief='sunken',
                                        validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan1_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan1_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_jg_entry = tk.Entry(root, width=56)
        self.f_dingdan2_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan2_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan3_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan3_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan4_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan4_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan5_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan5_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan6_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan6_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan7_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan7_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan8_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan8_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.f_dingdan9_entry = tk.Entry(root, width=56, highlightthickness=1.5, highlightcolor='blue',
                                         name='extract_merge_dingdan9_entry', validate="key", relief='sunken',
                                         validatecommand=(use.check_open_file1, '%P', '%W'))
        self.e10 = tk.Entry(root, width=56)
        self.f_choose_duizhao = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_duizhao_entry))
        self.f_choose_add = ttk.Button(root, text='订单列表+1', command=self.Extract_compare_add)
        self.f_choose_dingdan1 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan1_entry))
        self.f_choose_dingdan2 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan2_entry))
        self.f_choose_dingdan3 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan3_entry))
        self.f_choose_dingdan4 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan4_entry))
        self.f_choose_dingdan5 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan5_entry))
        self.f_choose_dingdan6 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan6_entry))
        self.f_choose_dingdan7 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan7_entry))
        self.f_choose_dingdan8 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan8_entry))
        self.f_choose_dingdan9 = ttk.Button(root, text='选择文件', command=lambda: choose_flie(self.f_dingdan9_entry))

    # 布局函数
    # @grid_progressbar
    def show(self):
        global f_list
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

        self.Extract_compare_Choose_more_file.grid(row=4, column=3, sticky='w')
        all_use.tk_list.append(self.Extract_compare_Choose_more_file)
        self.f_choose_add.grid(row=4, column=4, sticky='w')
        all_use.tk_list.append(self.f_choose_add)

        self.Extract_compare_Label1.grid(row=14, column=0, sticky='e')
        all_use.tk_list.append(self.Extract_compare_Label1)
        self.Extract_compare_Entry1.delete("0", 'end')
        self.Extract_compare_Entry1.grid(row=14, column=1, sticky='w')
        all_use.tk_list.append(self.Extract_compare_Entry1)
        self.Extract_compare_Label2.grid(row=15, column=0, sticky='e')
        all_use.tk_list.append(self.Extract_compare_Label2)
        self.Extract_compare_Entry2.delete("0", 'end')
        self.Extract_compare_Entry2.grid(row=15, column=1, sticky='w')
        all_use.tk_list.append(self.Extract_compare_Entry2)

        self.Extract_compare_Choose_work.grid(row=16, column=1)
        all_use.tk_list.append(self.Extract_compare_Choose_work)

        self.f_jg_label.grid(row=17, column=1, sticky='w')
        all_use.tk_list.append(self.f_jg_label)
        self.f_jg_entry.delete("0", 'end')
        self.f_jg_entry.grid(row=17, column=1, sticky='w')
        all_use.tk_list.append(self.f_jg_entry)

        f_list = 8
        # 获取时间
        now_time = time.strftime('%M分%S秒', time.localtime(time.time()))
        # 显示生成的名称
        self.Extract_compare_Entry1.delete("0", 'end')
        self.Extract_compare_Entry1.insert(tk.INSERT, '订单合并表-' + now_time)
        self.Extract_compare_Entry2.delete("0", 'end')
        self.Extract_compare_Entry2.insert(tk.INSERT, '对照-' + now_time)

    # 一次性选择多个文件
    def Extract_compare_Choose_files(self):
        global f_list
        file_paths = filedialog.askopenfilenames()
        if file_paths == '':
            return
        f_list = 8
        more_list = [[self.f_dingdan2_label, self.f_dingdan2_entry, self.f_choose_dingdan2],
                     [self.f_dingdan3_label, self.f_dingdan3_entry, self.f_choose_dingdan3],
                     [self.f_dingdan4_label, self.f_dingdan4_entry, self.f_choose_dingdan4],
                     [self.f_dingdan5_label, self.f_dingdan5_entry, self.f_choose_dingdan5],
                     [self.f_dingdan6_label, self.f_dingdan6_entry, self.f_choose_dingdan6],
                     [self.f_dingdan7_label, self.f_dingdan7_entry, self.f_choose_dingdan7],
                     [self.f_dingdan8_label, self.f_dingdan8_entry, self.f_choose_dingdan8],
                     [self.f_dingdan9_label, self.f_dingdan9_entry, self.f_choose_dingdan9], '']
        more_list = iter(more_list)
        self.f_dingdan1_entry.delete("0", 'end')
        self.f_dingdan1_entry.insert(tk.INSERT, file_paths[0])
        now_list = next(more_list)
        list2 = [self.f_dingdan2_entry, self.f_dingdan3_entry, self.f_dingdan4_entry, self.f_dingdan5_entry,
                 self.f_dingdan6_entry,
                 self.f_dingdan7_entry,
                 self.f_dingdan8_entry, self.f_dingdan9_entry]
        for file_path, entry in zip(file_paths[1:], list2):
            entry.delete("0", 'end')
            self.Extract_compare_add()
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
    def Extract_compare_add(self):
        global f_list
        i = f_list
        if i == 8:
            self.f_dingdan2_label.grid(row=5, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan2_label)
            self.f_dingdan2_entry.delete("0", 'end')
            self.f_dingdan2_entry.delete("0", 'end')
            self.f_dingdan2_entry.grid(row=5, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan2_entry)
            self.f_choose_dingdan2.grid(row=5, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan2)
            f_list = f_list - 1
        elif i == 7:
            self.f_dingdan3_label.grid(row=6, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan3_label)
            self.f_dingdan3_entry.delete("0", 'end')
            self.f_dingdan3_entry.delete("0", 'end')
            self.f_dingdan3_entry.grid(row=6, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan3_entry)
            self.f_choose_dingdan3.grid(row=6, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan3)
            f_list = f_list - 1
        elif i == 6:
            self.f_dingdan4_label.grid(row=7, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan4_label)
            self.f_dingdan4_entry.delete("0", 'end')
            self.f_dingdan4_entry.delete("0", 'end')
            self.f_dingdan4_entry.grid(row=7, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan4_entry)
            self.f_choose_dingdan4.grid(row=7, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan4)
            f_list = f_list - 1
        elif i == 5:
            self.f_dingdan5_label.grid(row=8, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan5_label)
            self.f_dingdan5_entry.delete("0", 'end')
            self.f_dingdan5_entry.delete("0", 'end')
            self.f_dingdan5_entry.grid(row=8, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan5_entry)
            self.f_choose_dingdan5.grid(row=8, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan5)
            f_list = f_list - 1
        elif i == 4:
            self.f_dingdan6_label.grid(row=9, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan6_label)
            self.f_dingdan6_entry.delete("0", 'end')
            self.f_dingdan6_entry.delete("0", 'end')
            self.f_dingdan6_entry.grid(row=9, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan6_entry)
            self.f_choose_dingdan6.grid(row=9, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan6)
            f_list = f_list - 1
        elif i == 3:
            self.f_dingdan7_label.grid(row=11, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan7_label)
            self.f_dingdan7_entry.delete("0", 'end')
            self.f_dingdan7_entry.delete("0", 'end')
            self.f_dingdan7_entry.grid(row=11, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan7_entry)
            self.f_choose_dingdan7.grid(row=11, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan7)
            f_list = f_list - 1
        elif i == 2:
            self.f_dingdan8_label.grid(row=12, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan8_label)
            self.f_dingdan8_entry.delete("0", 'end')
            self.f_dingdan8_entry.delete("0", 'end')
            self.f_dingdan8_entry.grid(row=12, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan8_entry)
            self.f_choose_dingdan8.grid(row=12, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan8)
            f_list = f_list - 1
        elif i == 1:
            self.f_dingdan9_label.grid(row=13, column=0, sticky='e')
            all_use.tk_list.append(self.f_dingdan9_label)
            self.f_dingdan9_entry.delete("0", 'end')
            self.f_dingdan9_entry.delete("0", 'end')
            self.f_dingdan9_entry.grid(row=13, column=1, sticky='w')
            all_use.tk_list.append(self.f_dingdan9_entry)
            self.f_choose_dingdan9.grid(row=13, column=2, sticky='w')
            all_use.tk_list.append(self.f_choose_dingdan9)
            f_list = f_list - 1

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
                                             ['订单编号', '名', '姓', '国家/地区', '街道', '寓所', '城市', '省/州',
                                              '邮编',
                                              '联系电话', '联系邮箱', '物流费用', '商品名称', '款式1', '款式2',
                                              '款式3'],
                                             self.f_jg_entry)
        if df is None:
            return
        df = df.cast(
            {"商品名称": pl.String, '款式1': pl.String, '款式2': pl.String, '款式3': pl.String},
            strict=False)

        ##合并订单
        if df is None:
            return
        df1 = df.unique('订单编号', keep='first')  # 去重
        # 最终dataframe
        # 原本列
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
        # 最终列
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
        df_end1 = df1.select(last_column1)
        df_end1 = df_end1.sort(by='订单号')

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
                    dict_end[list1[0]] = []
                    if list1[1] in dict1:
                        dict_end[list1[0]].extend(dict1[list1[1]])
                    else:
                        dict_end[list1[0]].append(list1[1])
                list1[2] = list1[2] - 1

        list_end = []
        for key, value in dict_end.items():
            value.insert(0, key)
            list_end.append(value)

        # 创建字典：订单创建时间，物流费用
        df1 = df.unique('订单编号', keep='first')
        df1 = df1.cast({'物流费用': pl.Float64})
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

        # 先按第一第二个字母排。再按时间排
        # 需要先将列表对齐，补全，才能list->pl.DataFrame(polars要求传入的list必须长度一致)
        max_len = 0
        for list2 in list_end:
            max_len = len(list2) if len(list2) > max_len else max_len
        for index, list3 in enumerate(list_end):
            list_end[index].extend([None] * (max_len - len(list3)))
        df_end2 = pl.DataFrame(data=list_end, orient='row')

        # 按第一列第二列排序
        df_end2 = df_end2.sort(by=['column_0', 'column_1'], descending=[False, False])

        # 保存文件
        full_path1 = os.path.join(os.path.dirname(self.f_dingdan1_entry.get()),
                                  self.Extract_compare_Entry1.get() + '.xlsx')
        full_path2 = os.path.join(os.path.dirname(self.f_dingdan1_entry.get()),
                                  self.Extract_compare_Entry2.get() + '.xlsx')
        # 保存订单合并表
        if not all_use.polarsdf_openpyxl_to_excel(df_end1, full_path1):
            self.e10.delete("0", 'end')
            self.e10.insert(tk.INSERT, '订单合并表保存失败')
            self.e10.configure(fg='red')
            return

        # 保存对照表
        if not all_use.polarsdf_openpyxl_to_excel(df_end2, full_path2, need_title=False):
            self.e10.delete("0", 'end')
            self.e10.insert(tk.INSERT, '对照表保存失败')
            self.e10.configure(fg='red')
            return

        self.f_jg_entry.delete("0", 'end')
        self.f_jg_entry.insert(tk.INSERT, '已完成，已保存到：' + full_path1 + '和' + full_path2)
        self.f_jg_entry.configure(fg='green')
        return