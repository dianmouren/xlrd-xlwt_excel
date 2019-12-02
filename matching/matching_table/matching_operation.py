

import xlrd


def ini_table(file_name, sheet_name):
    """ 读取表格信息，获取sheet """
    data = xlrd.open_workbook(file_name)

    table = data.sheet_by_name(sheet_name)

    return table


def read_model_table(model_table):
    """读取汇总模板表格的表头信息"""
    model_first_data = model_table.row_values(rowx=0)

    return model_first_data


def read_hz_table(hz_table):
    """获取中间表格信息，并将信息保存为大列表"""
    # 编写循环获取相应行的数据，将数据存储为大列表
    data_list = []
    for i in range(0, 16):
        data = hz_table.col_values(colx=i)
        data_list.append(data)

    return data_list


def write_table(worksheet, message, col):
    """将一列信息写入相应匹配表中"""
    a = 1
    m1 = message
    for n in m1:
        worksheet.write(a, col, n)
        a += 1




