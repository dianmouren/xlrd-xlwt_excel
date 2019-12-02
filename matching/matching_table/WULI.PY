

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

    write_table(worksheet, col_data[1], 12)
    write_table(worksheet, col_data[2], 16)
    write_table(worksheet, col_data[4], 17)
    write_table(worksheet, col_data[5], 11)
    write_table(worksheet, col_data[6], 19)
    write_table(worksheet, col_data[7], 13)
    write_table(worksheet, col_data[9], 22)
    write_table(worksheet, col_data[10], 34)
    write_table(worksheet, col_data[11], 40)
    write_table(worksheet, col_data[12], 41)
    write_table(worksheet, col_data[14], 2)
    write_table(worksheet, col_data[15], 32)

    workbook.save(path)
    print('物理设备表格信息匹配成功！')


def matching_wl():
    file_name = r'D:\githubproject\excel\files\汇总表\物理设备.xlsx'
    sheet_name = 'host'
    save_path = r'D:\githubproject\excel\files\匹配表\物理设备.xls'

    hz_file_name = r'D:\githubproject\excel\files\result\物理存储器.xls'
    hz_sheet_name = '物理设备'
    ini_model_table = ini_table(file_name, sheet_name)
    read_data = read_model_table(ini_model_table)
    hz_table = ini_table(hz_file_name, hz_sheet_name)
    col_data = read_hz_table(hz_table)
    write_model_table(sheet_name, read_data, save_path, col_data)



