# -*- coding=utf-8-*-
import tkinter as tk
from funtion import combined
from funtion import extract
from funtion import match
from funtion import merge
from funtion import scan
from funtion import splice
from funtion import statistics
from funtion import statistics_2
from funtion import supplement
from funtion import extract_2
from funtion import statistics_3
from funtion import extract_merge
from funtion import supplement_plus
from funtion import extract_merge_plus


if __name__ == '__main__':
    print('ppp')
    root = tk.Tk()
    root.title('Excel表格定制处理')
    root.geometry('900x600')
    fun_1 = extract.add_file(root)
    fun_2 = match.match(root)
    fun_3 = supplement.Supplement(root)
    fun_4 = merge.merge(root)
    fun_5 = combined.combined(root)
    fun_6 = statistics.statistics(root)
    fun_7 = splice.splice(root)
    fun_8 = scan.scan(root)
    fun_9 = statistics_2.statistics_2(root)
    fun_10 = extract_2.Extract_2(root)
    fun_11 = statistics_3.Statistics_3(root)
    fun_12 = extract_merge.Extract_Merge(root)
    fun_13 = supplement_plus.Supplement_plus(root)
    fun_14 = extract_merge_plus.Extract_Merge_plus(root)

    t1 = tk.Label(root,
                  text='文件操作区域：注意文件必须为xlsx文件，且不能设置为“只读"，文件名为红色时代表该文件无法正常打开，请检查该文件')
    t1.grid(row=0, column=0, columnspan=4, sticky='w')
    tk.Label(root, text='功能选择区域：').grid(row=200, column=0, columnspan=2, sticky='w')
    # 所有功能选择按钮
    Rad1 = tk.Radiobutton(root, text='功能1：订单列表信息提取', value=1, indicatoron=False, activeforeground='blue',
                          selectcolor='darkgrey', command=fun_1.show)
    Rad1.grid(row=201, column=0, sticky='w')
    Rad2 = tk.Radiobutton(root, text='功能2：根据发货单号整理制单信息(末尾有未匹配出的制单信息)', value=2,
                          indicatoron=False,
                          selectcolor='darkgrey', command=fun_2.show)
    Rad2.grid(row=201, column=1, sticky='w')
    Rad3 = tk.Radiobutton(root, text='功能3：回填运单号订单列表处理', value=3, indicatoron=False, selectcolor='darkgrey',
                          command=fun_3.show)
    Rad3.grid(row=202, column=0, sticky='w')
    Rad4 = tk.Radiobutton(root, text='功能4：对同一运单号的商品进行合并', value=4, indicatoron=False,
                          selectcolor='darkgrey',
                          command=fun_4.show)
    Rad4.grid(row=202, column=1, sticky='w')
    Rad5 = tk.Radiobutton(root, text='功能5：对多个订单与运单号进行合并', value=5, indicatoron=False,
                          selectcolor='darkgrey',
                          command=fun_5.show)
    Rad5.grid(row=203, column=0, sticky='w')
    Rad6 = tk.Radiobutton(root, text='功能6：统计商品数量', value=6, indicatoron=False, selectcolor='darkgrey',
                          command=fun_6.show)
    Rad6.grid(row=203, column=1, sticky='w')
    Rad7 = tk.Radiobutton(root, text='功能7：扫码前订单拼接', value=7, indicatoron=False, selectcolor='darkgrey',
                          command=fun_7.show)
    Rad7.grid(row=204, column=0, sticky='w')
    Rad8 = tk.Radiobutton(root, text='功能8：扫码', value=8, indicatoron=False, selectcolor='darkgrey',
                          command=fun_8.show)
    Rad8.grid(row=204, column=1, sticky='w')
    Rad9 = tk.Radiobutton(root, text='功能9：统计各商品数量', value=9, indicatoron=False, selectcolor='darkgrey',
                          command=fun_9.show)
    Rad9.grid(row=205, column=0, sticky='w')
    Rad10 = tk.Radiobutton(root, text='功能10：提取相关信息', value=10, indicatoron=False, selectcolor='darkgrey',
                           command=fun_10.show)
    Rad10.grid(row=205, column=1, sticky='w')
    Rad11 = tk.Radiobutton(root, text='功能11：统计信息', value=11, indicatoron=False, selectcolor='darkgrey',
                           command=fun_11.show)
    Rad11.grid(row=206, column=0, sticky='w')
    Rad12 = tk.Radiobutton(root, text='功能1+4：提取并对照信息', value=12, indicatoron=False, selectcolor='darkgrey',
                           command=fun_12.show)
    Rad12.grid(row=206, column=1, sticky='w')
    Rad13 = tk.Radiobutton(root, text='功能3plus：回填运单号订单列表处理', value=13, indicatoron=False,
                           selectcolor='darkgrey',
                           command=fun_13.show)
    Rad13.grid(row=207, column=0, sticky='w')
    Rad14 = tk.Radiobutton(root, text='功能1+4plus：提取并对照信息', value=14, indicatoron=False, selectcolor='darkgrey',
                           command=fun_14.show)
    Rad14.grid(row=207, column=1, sticky='w')
    tk.mainloop()
