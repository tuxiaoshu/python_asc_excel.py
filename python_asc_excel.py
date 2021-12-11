# 这是一个示例 Python 脚本。
import re
import os
# import csv
# import pathlib
# from os import statvfs_result
# from typing import Union

import pandas as pd
import numpy as np


# 按 Shift+F10 执行或将其替换为您的代码。
# 按 按两次 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。
def get_logfile():
    directory = os.getcwd()
    # flists = os.listdir(directory)
    # print(os.path.basename(directory))
    # os.listdir(directory):#

    for a, b, filenames in os.walk(directory):
        # print(filenames)
        for file in filenames:
            i = 1
            if not file.endswith('.asc'):
                #print("this file not .asc")
            else:
                flist = file
                print(flist)
                write_excel(get_asclog(flist), flist)
                exe_table(flist)
                i += i
    # print(filenames)
    # print(flist)
    return


def get_asclog(file_list):
    # with open(file_list,'r') as result_log:
    result_log = open(file_list, 'r')
    a = result_log.read()
    a_temp = re.compile(r'SUMMARY:(.*?)TEST END TIME:', re.S)
    asc_log_temp = re.findall(a_temp, a)
    # print(asc_log_temp)
    asc_log_temp_a = re.compile(r'DUT (.*?)DIEX=(.*?)DIEY=(.*?)SBIN=(.*?)HBIN=', re.S)
    asc_log = re.findall(asc_log_temp_a, str(asc_log_temp))
    # print(asc_log)
    return asc_log


def write_excel(asc_log, flist):
    fliename = flist.split(".asc")
    # print(fliename[0])
    result_excel = open(fliename[0] + '.csv', 'w+')
    result_excel.write('DUT'',''X'',''Y'',''SBIN''\n')
    for m in range(len(asc_log)):
        for n in range(len(asc_log[m])):
            result_excel.write(str(asc_log[m][n]))
            result_excel.write(',')
        result_excel.write('\n')
    result_excel.close()
    return


def exe_table(flist):
    fliename = flist.split(".asc")
    # print(fliename[0])
    p_data = pd.read_csv(fliename[0] + '.csv')
    pt = pd.pivot_table(p_data, values='Y', index='DUT', columns='X', aggfunc=np.sum)

    writer = pd.ExcelWriter(fliename[0] + '.xlsx', engine='xlsxwriter')
    pt.to_excel(writer, sheet_name=fliename[0] + '透视图')

    worksheet = writer.sheets[fliename[0] + '透视图']

    worksheet.conditional_format(1, 1, pt.shape[0], pt.shape[1], {'type': '3_color_scale'})

    writer.save()

    # print(pt)
    # pt.to_csv(fliename[0]+'.csv')
    # pt.to_excel(fliename[0] + '.xlsx')
    # print(pt)
    return


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    print_hi('PyCharm')
    get_logfile()
    # fliename= flists.split(".asc")
    # #print(fliename[0])
    # write_excel(get_asclog(flists))
    # exe_table()
# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
