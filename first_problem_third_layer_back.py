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
        temp_list.append(row[2])
    return temp_list

if __name__ == '__main__':
    layer = read_fourth_layer()
    temp = 37
    for i in xrange(1,len(layer)):
        # 0.045 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (74.2 * 1726 * 0.0001)
        t = ((74.2 * 1726 * 0.000036) * (layer[i]-layer[i-1]) / 0.045) + layer[i]
        if t > 75 or t < 0:
            t = temp
        temp = t
        print t