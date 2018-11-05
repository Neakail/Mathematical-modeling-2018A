# -*- coding:utf-8 -*-
import xlwt
import xlrd

def read_second_layer():
    temp_list = []
    data = xlrd.open_workbook('second_problem_second_layer_1.xls')
    table = data.sheets()[0]
    ncols = table.ncols
    for i in xrange(1,ncols):
        row = table.col_values(i)
        temp_list.append(row)
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
            ws.write(count, number, i)
            count += 1
        number += 1
    wx.save(filename)


if __name__ == '__main__':
    second_layers = read_second_layer()
    number = 1
    all_data = []
    for second_layer in second_layers:
        # if number < 11:
            list_K = [37] * 36
            experiment_num = second_layer[0]
            print experiment_num
            second_layer = second_layer[1:]
            temp_data = [experiment_num]
            # print second_layer
            for i in second_layer:

                temp_list = []
                for each_mm in xrange(1, 37):
                    # 更新温度
                    if second_layer.index(i) == 0:
                        if each_mm == 1:
                            # Q = 0.000045 * 0.1 * (i - list_K[each_mm-1])
                            temp_t = (1+(each_mm * 1.0/36)) * 0.045 * (i - list_K[each_mm - 1]) / (74.2 * 1726 * 0.0001)
                        else:
                            # Q = 0.000045 * 0.1 * (list_K[each_mm-2] - list_K[each_mm-1])
                            temp_t = (1+(each_mm * 1.0/36)) * 0.045 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (74.2 * 1726 * 0.0001)
                        # temp_t = Q * 1.0 / (7.42 * 1.726 * 0.000001)
                    else:
                        if each_mm == 1:
                            temp_t = (1+(each_mm * 1.0/36)) * 0.045 * (i - list_K[each_mm - 1]) / (74.2 * 1726 * 0.0001)
                        else:
                            # temp_t = 0.045 * (all_data[-1][-1] - list_K[each_mm-1]) / (74.2 * 1726 * 0.0001)
                            temp_t = (1+(each_mm * 1.0/36)) * 0.045 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (74.2 * 1726 * 0.0001)
                            # temp_t = temp_t * 1.0 / 2
                    temp_list.append(list_K[each_mm - 1] + temp_t)
                list_K = temp_list
                # print "第"+str(number)+"秒："
                # print list_K
                temp_data.append((list_K[0]+list_K[-1]) * 1.0 /2)
            all_data.append(temp_data)
        # number += 1
    write_in_xlwt(all_data, 'second_problem_third_layer_1.xls')