import xlrd
from glob2 import glob
import pandas as pd
from pandas import DataFrame

# merge1()使用情况：
# 适用于合并一个工作簿中的子表
def merge1():
    excel_name = r'C:\Users\A\Desktop\WorksheetMerge\file.xlsx'
    wb = xlrd.open_workbook(excel_name)

    # 获取workbook中所有的表格
    sheets = wb.sheet_names()

    # 循环遍历所有sheet
    df_merge = DataFrame()
    for i in range(len(sheets)):
        # skiprows=[1],忽略第二行
        # converters是转换列数据的数据类型的
        df = pd.read_excel(excel_name, sheet_name=i, index=False, encoding='utf8', converters={'账号ID': str})
        df_merge = df_merge.append(df)

    df_merge.to_excel(r'C:\Users\A\Desktop\WorksheetMerge\合并.xlsx', sheet_name='合并', index=False)

# merge2()使用情况：
# 每个表中有相同的子表，且名称相同
def merge2():
    # 获取文件夹内所有.xlsx文件
    file_list = glob(r'C:\Users\A\Desktop\WorksheetMerge' + '\*.xlsx')
    print(file_list)
    for name in file_list:
        print(name.split('\\')[-1].replace('.xlsx', ''))

    FB_merge = DataFrame()
    GG_merge = DataFrame()
    for name in file_list:
        wb = xlrd.open_workbook(name)
        # 获取workbook中所有的表格
        sheets = wb.sheet_names()

        for sheet_name in sheets:
            if sheet_name == "FB":
                df = pd.read_excel(name, sheet_name=sheet_name, encoding='utf8', converters={'账号ID': str})
                FB_merge = FB_merge.append(df)
            else:
                df = pd.read_excel(name, sheet_name=sheet_name, encoding='utf8', converters={'账号ID': str})
                GG_merge = GG_merge.append(df)

    # 创建工作簿
    with pd.ExcelWriter(r'C:\Users\A\Desktop\WorksheetMerge\合并.xlsx') as writer:
        FB_merge.to_excel(writer, index=False)
        GG_merge.to_excel(writer, sheet_name='GG', index=False)
        pass

def merge3():
    # 获取文件夹内所有.xlsx文件
    file_list = glob(r'C:\Users\A\Desktop\WorksheetMerge' + '\*.xlsx')
    print(file_list)
    df_merge = DataFrame()
    for name in file_list:
        print(name.split('\\')[-1].replace('.xlsx', ''))
        df = pd.read_excel(name, encoding='utf8', converters={'账号ID': str})
        df_merge = df_merge.append(df)

    # 创建工作簿
    with pd.ExcelWriter(r'C:\Users\A\Desktop\WorksheetMerge\合并.xlsx') as writer:
        df_merge.to_excel(writer, index=False)
        pass

if __name__ == '__main__':
    # 需要哪个l类型的就调用对应的函数
    merge1()
    merge2()
    merge3()
