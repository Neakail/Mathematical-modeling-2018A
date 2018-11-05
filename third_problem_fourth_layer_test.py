# -*- coding:utf-8 -*-
import xlwt
import xlrd
from decimal import Decimal
import math

def read_third_layer():
    temp_list = []
    data = xlrd.open_workbook('third_problem_third_layer_test.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    for i in xrange(1,nrows):
        row = table.row_values(i)
        temp_list.append(row[-1])
    return temp_list

def write_in_xlwt(data,filename):
    wx = xlwt.Workbook()
    ws = wx.add_sheet('1',cell_overwrite_ok=True)
    ws.write(0, 0, 'second'.decode('utf-8'))
    ws.write(0, 1, '1mm'.decode('utf-8'))
    ws.write(0, 2, '2mm'.decode('utf-8'))
    ws.write(0, 3, '3mm'.decode('utf-8'))
    ws.write(0, 4, '4mm'.decode('utf-8'))
    ws.write(0, 5, '5mm'.decode('utf-8'))
    ws.write(0, 6, '6mm'.decode('utf-8'))
    # ws.write(0, 7, 'avg'.decode('utf-8'))
    number = 1
    for each_data in data:
        print each_data
        count = 0
        for i in each_data:
            ws.write(number, count, i)
            count += 1
        number += 1
    wx.save(filename)


if __name__ == '__main__':
    all_data = []
    third_layer = read_third_layer()
    list_K = [37] * 39
    number = 1
    for i in third_layer:
        temp_data = []
        temp_list = []
        for each_mm in xrange(1, 40):
            if third_layer.index(i) == 0:
                if each_mm == 1:
                    # Q = 0.000028 * 0.1 * (i - list_K[each_mm-1])
                    temp_t = (1+(each_mm * 1.0/50)) * 0.028 * (i - list_K[each_mm - 1]) / (1.18 * 1005 * 0.0001)
                else:
                    # Q = 0.000028 * 0.1 * (list_K[each_mm-2] - list_K[each_mm-1])
                    temp_t = (1+(each_mm * 1.0/50)) * 0.028 * (list_K[each_mm - 2] - list_K[each_mm-1]) / (1.18 * 1005 * 0.0001)
            else:
                if each_mm == 1:
                    temp_t = (1+(each_mm * 1.0/36)) * 0.028 * (i - list_K[each_mm - 1]) / (1.18 * 1005 * 0.0001)
                else:
                    # temp_t = 0.028 * (all_data[-1][-1] - list_K[each_mm-1]) / (1.18 * 1005 * 0.0001)
                    temp_t = (1+(each_mm * 1.0/36)) * 0.028 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (1.18 * 1005 * 0.0001)
                    # temp_t = temp_t * 1.0 / 2
            # temp_t = Q * 1.0 * 1000000 / (1.18 * 1.005)
            temp_list.append(list_K[each_mm - 1] + temp_t)
        list_K = temp_list
            # print temp_t
        print "第"+str(number)+"秒："
        # print list_K
        temp = 0
        for i in xrange(39):
            temp += list_K[i]
        number += 1
        temp_data.append(number-1)
        temp_data[1:40] = list_K
        temp_data.append((temp_data[1]+temp_data[39]) / 2)
        print temp_data
        all_data.append(temp_data)

    write_in_xlwt(all_data, 'third_problem_fourth_layer_test.xls')