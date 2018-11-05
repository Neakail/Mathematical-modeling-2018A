import xlrd
import xlwt
max_t = 37.4445048659
min_t = 37.3542778901
# 37.3542778901
def thickness_normalization(length1,length2):
    return (length1+length2-12)*1.0/(314-12)

def weight_normol(length1,length2):
    return (length1*862+length2*1.18-1.18*6-862*6)/(250*862+64*1.18-1.18*6-862*6)

def t_normol(t):
    return (t-min_t) * 1.0 / (max_t-min_t)

def read_fourth_layer(filename):
    temp_list = []
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    ncols = table.ncols
    for i in xrange(6,ncols):
        row = table.col_values(i)
        temp_list.append(row)
    return temp_list

# 34.0 64 0.0615335321581
# 0.0905712412862
min_a = 0.0872337738598
min_length1 = 0
min_length2 = 0
wx = xlwt.Workbook()
ws = wx.add_sheet('1', cell_overwrite_ok=True)
number = 0
for length2 in xrange(6,65):
    filename = 'third_problem/5/' + str(length2) + '.xls'
    print filename
    layers = read_fourth_layer(filename)
    for layer in layers:
        sum_ = 0
        length1 = layer[0]
        for i in layer[1:]:
            sum_ += i
        avg = sum_ * 1.0 / len(layer[1:])
        ws.write(number, 0, length1)
        ws.write(number, 1, length2)
        ws.write(number, 2, thickness_normalization(length1,length2))
        ws.write(number, 3, weight_normol(length1,length2))
        ws.write(number, 4, t_normol(avg))
        a = 0.197 * thickness_normalization(length1, length2) + 0.311 * weight_normol(length1, length2) + 0.491 * t_normol(avg)
        if a < min_a:
            min_a = a
            min_length1 = length1
            min_length2 = length2
        ws.write(number, 5, a)
        number += 1
wx.save("optimization_15.xls")
print min_length1,min_length2,min_a
