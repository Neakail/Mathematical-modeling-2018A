# -*- coding:utf-8 -*-
import xlwt
import xlrd

def read_first_layer():
    temp_list = []
    data = xlrd.open_workbook('third_problem_third_layer_5.xls')
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
            ws.write(count,number,i)
            count += 1
        number += 1
    wx.save(filename)


if __name__ == '__main__':
    thirds_layer = read_first_layer()
    for mm in xrange(1, 65):
        all_data = []
        print str(mm) + '层'
        for third_layer in thirds_layer:
            list_K = [37] * mm
            temp_data = [third_layer[0]]
            third_layer = third_layer[1:]
            for i in third_layer:
                temp_list = []
                for each_mm in xrange(1, mm+1):
                    if third_layer.index(i) == 0:
                        if each_mm == 1:
                            temp_t = (1+(each_mm * 1.0/mm)) * 0.37 * (i - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                        else:
                            temp_t = (1+(each_mm * 1.0/mm)) * 0.37 * (list_K[each_mm-2] - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                    else:
                        if each_mm == 1:
                            temp_t = (1+(each_mm * 1.0/mm)) * 0.37 * (i - list_K[each_mm-1]) / (862 * 2100 * 0.0001)
                        else:
                            temp_t = (1+(each_mm * 1.0/mm)) * 0.37 * (list_K[each_mm - 2] - list_K[each_mm - 1]) / (862 * 2100 * 0.0001)

                    temp_list.append(list_K[each_mm-1]+temp_t)
                list_K = temp_list
                temp_data.append((list_K[0]+list_K[-1]) / 2)
            # print temp_data
            all_data.append(temp_data)
        write_in_xlwt(all_data, 'third_problem/5/'+str(mm)+'.xls')