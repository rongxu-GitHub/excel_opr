import os
import pandas as pd


def file_merge(file_path_list, str_col):
    """
    不同文件薄合并
    """
    # 定义一个列表存放读取的数据
    data_list = []

    # 遍历所有文件，进行合并操作
    for file_path in file_path_list:
        # 依次读取每个文件，并追加到data_list中，如果有账户ID，记得转换成str类型，header指定第几行作为表头
        df = pd.read_excel(file_path, converters={str_col: str})
        data_list.append(df)

    # 输出到指定目录下
    data = pd.concat(data_list)
    return data


def sheet_merge(file_path, output_path, str_col):
    """
    同文件薄下多个sheet表合并
    """
    # 读取文件
    file = pd.ExcelFile(file_path)
    # 创建一个列表，用来合并读取的每个sheet表
    dfs = []
    # 遍历所有sheet表，添加到dfs[]
    for sheet_name in file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, converters={str_col: str})
        dfs.append(df)

    # 使用pd.concat()方法拼接成一个DataFrame
    data = pd.concat(dfs)
    # 输出到指定文件，index设置不显示
    data.to_excel(output_path, index=False)


# 获取目录内的所有文件绝对路径（不包括子目录）
def get_file_path_list(dir_path):
    """
    :param dir_path: 目录
    :return: 目录内所有文件绝对路径列表（包括子目录中的文件）
    """
    # 获取所有文件的绝对路径
    file_path_list = []
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_path_list.append(file_path)
    return file_path_list



if __name__ == '__main__':

    """
    不同文件合并
    """
    # 获取所有要合并的文件的绝对路径
    dir_path = r'C:\Users\A\Desktop\WorksheetMerge'
    file_path_list = get_file_path_list(dir_path)

    # 进行合并操作
    str_col = '帐户编号'
    data = file_merge(file_path_list, str_col)

    # 输出并指定路径文件名
    data.to_excel(r'C:\Users\A\Desktop\WorksheetMerge\海外户3月消耗.xlsx', index=False)
    print('合并完成！')


    # """
    # 同个文件不同sheet表合并
    # """
    # # 定义文件绝对路径
    # file_path = r'C:\Users\A\Desktop\WorksheetMerge\1月-FB广告账户消耗明细-深圳三部阿米巴A组.xlsx'
    # output_path = r'C:\Users\A\Desktop\WorksheetMerge\合并结果.xlsx'
    # str_col = '账户ID'
    # sheet_merge(file_path, output_path, str_col)
    # print('合并完成！')



