# -*- coding:utf-8 -*-
import xlwt
import xlrd

def read_first_layer():
    temp_list = []
    data = xlrd.open_workbook('first_problem_first_layer_6.xls')
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
    first_layer = read_first_layer()
    list_K = [37] * 60
    number = 1
    for i in first_layer:
        temp_data = []
        temp_list = []
        for each_mm in xrange(1, 61):
            if first_layer.index(i) == 0:
                if each_mm == 1:
                    # Q = 0.00037 * 0.1 * (i - list_K[each_mm-1])
                    temp_t = (1+(each_mm * 1.0/60)) * 0.37 * (i - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                else:
                    # Q = 0.00037 * 0.1 * (list_K[each_mm-2] - list_K[each_mm-1])
                    temp_t = (1+(each_mm * 1.0/60)) * 0.37 * (list_K[each_mm-2] - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                # temp_t = Q * 1.0 / (8.62 * 2.1 * 0.00001)
            else:
                if each_mm == 1:
                    temp_t = (1+(each_mm * 1.0/60)) * 0.37 * (i - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                else:
                    temp_t = (1+(each_mm * 1.0/60)) * 0.37 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (862 * 2100 * 0.0001)
                    # temp_t = (((all_data[-1][-1]-list_K[each_mm - 1]) * 1.0 / all_data[-1][-1]) + 1) * 0.37 * (list_K[each_mm-2] - list_K[each_mm - 1]) / (862 * 2100 * 0.0001)
                    # temp_t = temp_t * 1.0 / 2
            temp_list.append(list_K[each_mm-1]+temp_t)
        list_K = temp_list
        print "第"+str(number)+"秒："
        print list_K
        temp = 0
        for i in xrange(60):
            temp += list_K[i]
        number += 1
        temp_data.append(number-1)
        temp_data[1:61] = list_K
        temp_data.append((temp_data[1]+temp_data[60]) / 2)
        # print temp_data
        all_data.append(temp_data)
    write_in_xlwt(all_data, 'first_problem_second_layer_7.xls')