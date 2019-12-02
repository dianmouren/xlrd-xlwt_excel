from matching.matching_table.BAOLEI import matching_bl
from matching.matching_table.CUNCHU import matching_cc
from matching.matching_table.FANGHQ import matching_fhq
from matching.matching_table.FUZAI import matching_fz
from matching.matching_table.JIAOHUAN import matching_jhj
from matching.matching_table.LUYOU import matching_lyq
from matching.matching_table.WULI import matching_wl


def matching_table():

    # 物理设备信息匹配操作
    matching_wl()

    # 存储设备信息匹配操作
    matching_cc()

    # 堡垒机信息匹配操作
    matching_bl()

    # 防火墙信息匹配操作
    matching_fhq()

    # 交换机信息匹配操作
    matching_jhj()

    # 路由器信息匹配操作
    matching_lyq()

    # 负载均衡信息匹配操作
    matching_fz()





