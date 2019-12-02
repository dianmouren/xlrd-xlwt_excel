import xlwt


def add_sheet(sheet_name):
    workbook = xlwt.Workbook(encoding='utf-8')  # 规定写的书写格式为utf-8
    worksheet = workbook.add_sheet(sheet_name)  # 在表格中添加名字为sheet_name的表

    return worksheet, workbook


def write_excel(col_data, col, worksheet):
    c = 0
    for i in col_data:
        c += 1
        worksheet.write(col, c-1, i)

