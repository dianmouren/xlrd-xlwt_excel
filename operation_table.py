
from operation.initialize import initialize
from operation.operation_table.baolei import operation_bl
from operation.operation_table.cunchu import operation_cc
from operation.operation_table.fanghq import operation_fhq
from operation.operation_table.fuzai import operation_fz
from operation.operation_table.jiaohuan import operation_jh
from operation.operation_table.luyou import operation_ly
from operation.operation_table.wuli import operation_wli
from operation.write import add_sheet


def operation_table():
    op_excel = r'D:\githubproject\excel\files\云计算部资产梳理.xlsx'
    op_sheet = '汇总'
    ini_table, ini_cols, ini_rows = initialize(op_excel, op_sheet)

    # 物理设备信息汇总操作
    ini_sheet = '物理设备'  # 定义sheet名称
    index_word = '物理服务器'  # 定义匹配字段
    ini_worksheet_wl, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_wli(ini_table, ini_rows, ini_worksheet_wl, ini_workbook, index_word)  # 业务逻辑处理

    # 堡垒机信息汇总操作
    ini_sheet = '堡垒机'  # 定义sheet名称
    index_word = '堡垒机'  # 定义匹配字段
    ini_worksheet_bl, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_bl(ini_table, ini_rows, ini_worksheet_bl, ini_workbook, index_word)  # 业务逻辑处理

    # 存储设备信息汇总操作
    ini_sheet = '存储设备'  # 定义sheet名称
    index_word = '存储'  # 定义匹配字段
    ini_worksheet_cc, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_cc(ini_table, ini_rows, ini_worksheet_cc, ini_workbook, index_word)  # 业务逻辑处理

    # 防火墙信息汇总操作
    ini_sheet = '防火墙'  # 定义sheet名称
    index_word = '防火墙'  # 定义匹配字段
    ini_worksheet_fhq, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_fhq(ini_table, ini_rows, ini_worksheet_fhq, ini_workbook, index_word)  # 业务逻辑处理

    # 负载均衡信息汇总操作
    ini_sheet = '负载均衡'  # 定义sheet名称
    index_word = '负载'  # 定义匹配字段
    ini_worksheet_fz, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_fz(ini_table, ini_rows, ini_worksheet_fz, ini_workbook, index_word)  # 业务逻辑处理

    # 交换机信息汇总操作
    ini_sheet = '交换机'  # 定义sheet名称
    index_word = '交换'  # 定义匹配字段
    ini_worksheet_jh, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_jh(ini_table, ini_rows, ini_worksheet_jh, ini_workbook, index_word)  # 业务逻辑处理

    # 路由器信息汇总操作
    ini_sheet = '路由器'  # 定义sheet名称
    index_word = '路由'  # 定义匹配字段
    ini_worksheet_ly, ini_workbook = add_sheet(ini_sheet)  # 创建sheet操作
    operation_ly(ini_table, ini_rows, ini_worksheet_ly, ini_workbook, index_word)  # 业务逻辑处理




