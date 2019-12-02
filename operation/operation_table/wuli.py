
from operation.write import write_excel


def operation_wli(table, sheet_rows, worksheet_wl, workbook_wl, index):

    a = 2
    d = 0
    while a < sheet_rows:
        b = table.row_values(rowx=a)  # 读取一行数据
        if index in b[1]:
            write_excel(b, d, worksheet_wl)  # 调用写入方法，将数据写入
            d += 1  # 设定写入行数为每次增加1
        a += 1  # 设定循环次数，保证遍历所有行

    workbook_wl.save(r'd:\githubproject\excel\files\result\物理存储器.xls')  # 将表格存为XXX.xls
    print('物理服务器数据已经汇总完毕！')






