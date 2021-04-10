import pandas as pd
import os


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


# 遍历目录内的文件并拆分
def file_split(file_path, split_column, acc_id):
    """
    :param file_path: 要拆分的文件绝对路径
    :param split_column: 分列依据列名
    :param acc_id: 需要转换文本格式的列，通常为账户ID列
    :return: None
    """
    print('正在拆分 %s' % file_path)
    # 设置一个拆分文件计数变量
    file_counts = 0
    # 读取文件
    data = pd.read_excel(file_path, encode='gbk', converters={acc_id: str})

    # 获取分列列的不重复值列表，只会识别出最新创建的工作表，建议汇总工作薄中只保留一份工作表。
    class_list = list(data[split_column].unique())

    # 按客户名称拆分后，各自保存到对应的对接人文件夹中，先定义各对接人文件夹名称
    split_path = file_path[:-5]
    # 判断文件路径是否存在，不存在则创建
    if not os.path.exists(split_path):
        os.makedirs(split_path)

    for i in class_list:
        # 判断筛选同名行
        data_cut = data[data[split_column] == i]
        # 文件名称
        name = '%s.xlsx' % (i)

        # 保存文件到目录下
        data_cut.to_excel(split_path + '\\' + name, encoding='utf-8', index=False)
        file_counts += 1

    print("拆分文件个数为:", file_counts)


# 表格求和函数
def sheet_sum(file_path, acc_id, total, sum_column):
    """
    :param file_path: 当前需要求和的文件所在目录
    :param acc_id: 需要设置文本格式的列，通常为账号ID
    :param total: “总计”写入所在列名
    :param sum_column: 求和列所在列名
    :return: None
    """

    # 读取文件
    df = pd.read_excel(file_path, converters={acc_id: str})

    # 获取求和值行号，求和值所在行号=文件总行数+1
    sum_row_index = df.shape[0] + 1

    # 求和并写入指定列求和值行号位置（即指定行列位置）
    df.loc[sum_row_index, total] = "总计"
    df.loc[sum_row_index, sum_column] = df.loc[:sum_row_index, sum_column].sum()

    # 保留2位小数
    df = df.round(2)

    # 输出到源文件，覆盖即可
    df.to_excel(file_path, index=False)
    print("%s求和写入成功！" % file_path.split("\\")[-1])

# 获取子目录绝对路径
def get_subdir_path_list(dir_path):
    """
    :param dir_path: 目录绝对路径
    :return: subdir_path_list[]
    """
    subdir_path_list = []
    for root, dirs, files in os.walk(dir_path):
        for dir in dirs:
            # 获取孙目录绝对路径
            subdir_path = os.path.join(root, dir)
            subdir_path_list.append(subdir_path)
    return subdir_path_list

if __name__ == '__main__':

    print("开始第一次拆分，按运管对接人")
    # 根据运管对接人进行第一次拆分
    dir_path = r'C:\Users\A\Desktop\WorksheetSplite'
    file_path_list = get_file_path_list(dir_path)
    split_column1 = '负责运管'
    # acc_id需要从表格中复制到这里，否则会导致第二次拆分出现问题
    acc_id = '账户ID'
    for file_path in file_path_list:
        file_split(file_path, split_column1, acc_id)

    # print("开始第二次拆分，按客户名称")
    # print("-" * 20)
    # # 根据客户名称进行第二次拆分
    # # 先获取子目录绝对路径
    # subdir_path_list = get_subdir_path_list(dir_path)
    # split_column2 = '客户名称'
    # # 求和变量
    # total = "客户名称"
    # sum_column = "金额"
    # # 定义一个统计拆分文件总个数的变量
    # count = 0
    # for subdir_path in subdir_path_list:
    #     # 获取子目录下所有文件绝对路径
    #     file_path_list = get_file_path_list(subdir_path)
    #     # 拆分
    #     for file_path in file_path_list:
    #         # file_path是第一次拆分后的文件的绝对路径，
    #         file_split(file_path, split_column2, acc_id)
    #         print("-" * 20)
    #
    #     # 求和
    #     subdir_path_list = get_subdir_path_list(subdir_path)
    #     for subdir_path in subdir_path_list:
    #         # 输出要求和的目录
    #         print(subdir_path)
    #         file_path_list = get_file_path_list(subdir_path)
    #         for file_path in file_path_list:
    #             # sheet_sum(file_path, acc_id, total, sum_column)
    #             count +=1
    #         print("-" * 20)
    #
    # print("共拆分出%s个文件。" % count)