# -*- coding:utf-8 -*-
import xlwt
import xlrd
from decimal import Decimal
import math

def read_fourth_layer():
    temp_list = []
    data = xlrd.open_workbook('CUMCM-2018-Problem-A-Chinese-Appendix.xlsx')
    table = data.sheets()[1]
    nrows = table.nrows
    for i in xrange(2,nrows):
        row = table.row_values(i)
        temp_list.append(row[4])
    return temp_list

if __name__ == '__main__':
    layer = read_fourth_layer()
    temp = 37
    for i in xrange(1,len(layer)):
        # 0.028 * (list_K[each_mm - 2] - list_K[each_mm-1]) / (1.18 * 1005 * 0.0001)
        # Q = 1.18 * 1.005 * 0.0000001 * (layer[i]-layer[i-1])
        # t = Q * 1.0 / (0.000028 * 5)
        t = ((1377 * 300 * 0.0006) * (layer[i]-layer[i-1]) / 0.082) + layer[i]
        if t > 75 or t < 0:
            t = temp
        temp = t
        print t