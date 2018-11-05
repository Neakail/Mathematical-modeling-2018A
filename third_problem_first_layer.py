# -*- coding:utf-8 -*-
import xlwt

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
    ws.write(0, 7, 'avg'.decode('utf-8'))
    number = 1
    for each_data in data:
        for i in each_data:
            ws.write(number, each_data.index(i), i)
        number += 1
    wx.save(filename)

if __name__ == '__main__':
    all_data = []
    list_K = [80,37,37,37,37,37,37]
    for each_time in xrange(1,1801):
        temp_data = []
        temp_list = [80]
        for each_mm in xrange(1,7):
            # Q = 0.000082 * 0.1 * (list_K[each_mm-1] - list_K[each_mm])
            # temp_t = Q * 1.0 / (1.377 * 0.3 * 0.0001)
            # print temp_t
            temp_t = 0.082 * (list_K[each_mm-1] - list_K[each_mm]) / (1377 * 300 * 0.0001)
            temp_list.append(list_K[each_mm] + temp_t)
        list_K = temp_list
        print "第"+str(each_time)+"秒："
        print list_K
        temp = 0
        for i in xrange(1,7):
            temp += list_K[i]
        temp_data = list_K[1:]
        print "平均温度："+str(temp*1.00/6)
        temp_data[0] = each_time
        temp_data[1:7] = list_K[1:]
        temp_data.append(temp*1.00/6)
        all_data.append(temp_data)
    write_in_xlwt(all_data,'third_problem_first_layer.xls')