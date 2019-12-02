

import xlwt

from matching.matching_table.matching_operation import *


def write_model_table(sheet_name, write_data, path, col_data):
    """书写信息，包括表头信息和内容"""
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet(sheet_name)
    a = 1
    b = 0
    for i in write_data:
        worksheet.write(0, b, i)
        b += 1

    # todo 使用列表索引操作，将大列表中的小列表（一列元素）取出来
    # todo 使用切片操作，将一个单元格的数据取出来
    # todo 使用循环执行第三步，将每一列的数据写入

    write_table(worksheet, col_data[1], 0)
    write_table(worksheet, col_data[2], 17)
    write_table(worksheet, col_data[4], 16)
    write_table(worksheet, col_data[5], 4)
    write_table(worksheet, col_data[6], 14)
    write_table(worksheet, col_data[10], 18)
    write_table(worksheet, col_data[11], 13)
    write_table(worksheet, col_data[12], 8)
    write_table(worksheet, col_data[13], 15)
    write_table(worksheet, col_data[14], 6)
    write_table(worksheet, col_data[15], 19)

    workbook.save(path)
    print('堡垒机表格信息匹配成功！')


def matching_bl():
    file_name = r'D:\githubproject\excel\files\汇总表\堡垒机.xlsx'  # 模板文件地址
    sheet_name = 'host'
    save_path = r'D:\githubproject\excel\files\匹配表\堡垒机.xls'  # 最终匹配文件存储地址

    hz_file_name = r'D:\githubproject\excel\files\result\堡垒机.xls'  # 中间表格地址
    hz_sheet_name = '堡垒机'  # 中间表格sheet名
    ini_model_table = ini_table(file_name, sheet_name)
    read_data = read_model_table(ini_model_table)
    hz_table = ini_table(hz_file_name, hz_sheet_name)
    col_data = read_hz_table(hz_table)
    write_model_table(sheet_name, read_data, save_path, col_data)





