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
        temp_list.append(row[3])
    return temp_list

if __name__ == '__main__':
    temp = 37
    layer = read_fourth_layer()
    for i in xrange(1,len(layer)):
        t = ((862 * 2100 * 0.006) * (layer[i]-layer[i-1]) / 0.37) + layer[i]
        if t > 75 or t < 0:
            t = temp
        temp = t
        print t