import xlrd


def initialize(excel, sheet):
    data = xlrd.open_workbook(excel)
    table = data.sheet_by_name(sheet)  # 获取表格中的sheet

    if data.sheet_loaded(sheet):  # 加载sheet
        print('{}sheet已经上载完毕'.format(sheet))

    sheet_rows = table.nrows  # 获取sheet有效行数
    sheet_cols = table.ncols  # 获取sheet有效列数
    print('云计算部资产梳理表格共有{}行'.format(sheet_rows))
    print('云计算部资产梳理表格共有{}列'.format(sheet_cols))
    print('********************')
    print('开始进行云计算部资产梳理表格汇总工作！！')
    print('********************')
    return table, sheet_cols, sheet_rows
