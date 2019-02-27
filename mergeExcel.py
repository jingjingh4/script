# -*- coding:utf-8*-

import os
import pandas as pd
import xlrd

def MergeExcel(filepath,outfile):
    # read title name
    rowname = []
    files = os.listdir(filepath)
    for i in files:
        wb = xlrd.open_workbook(filepath+'/%s' % i)
        sh = wb.sheet_by_index(0)
        for j in sh.row_values(0):
            rowname.append(j)
    rowname = list(set(rowname))  # 列名去重
    print("获取列名成功！")

    for i, j in enumerate(files):
        print(j, "开始获取数据！")
        data = pd.read_excel(filepath+'/%s' % j)
        if i == 0:
            dff = pd.DataFrame(data, columns=rowname)
        if i != 0:
            dff = dff.append(data, ignore_index=True)
        print(j, "获取数据成功！")
    print("正在合成！")
    # 保存到一个文件下
    dff.to_csv(outfile)


if __name__ == '__main__':
    filepath ='D:/Users/install/Desktop/Jing/JJ/JD whole year promotion/raw data'
    outfile ="result.csv"
    MergeExcel(filepath,outfile)