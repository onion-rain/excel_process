# coding=utf-8

import xlrd, xlwt
from xlrd import *

import win32com.client
import csv
import sys

def get_table(excel_path):
    data = xlrd.open_workbook(excel_path)
    table = data.sheets()[0]
    # table = data.sheet_by_index(0)
    return table

def write_row(ws, sh, index, cursor):
    for i, v in enumerate(sh.row_values(index)):
        ws.write(cursor, i, v)
    cursor = cursor + 1
    return cursor

if __name__ == '__main__':
    in1_path = 'io/2019年档案入库统计.xls' # in
    in2_path = 'io/2019年新增就业登记情况_nopassword.xlsx' # in
    out_path = 'io/2019年网签情况.xls' # out
    in1 = get_table(in1_path)
    in2 = get_table(in2_path)
    out = get_table(out_path)

    in1_idcard_value = in1.col_values(2)[1:] # 含空
    in2_idcard_value = in2.col_values(0)[1:] # 全
    out_idcard_value = out.col_values(5)[1:] # 全

    # 仅凭id card求得交集
    union_id = [i1 for i1 in in1_idcard_value if i1 in in2_idcard_value]

    # 提出out中的id
    final_id = [x for x in union_id if x not in out_idcard_value]

    new_in1 = xlwt.Workbook()
    new_in1_sheet0 = new_in1.add_sheet('sheet0')
    new_in1_cursor = write_row(new_in1_sheet0, in1, 0, 0)
    new_in1_cursor = 1
    new_in2 = xlwt.Workbook()
    new_in2_sheet0 = new_in2.add_sheet('sheet0')
    new_in2_cursor = write_row(new_in2_sheet0, in2, 0, 0)
    new_in2_cursor = 1
    # new_out = xlwt.Workbook()
    # new_out_sheet = new_out.add_sheet('sheet0')
    # new_out_cursor = 0
    for id in final_id:
        new_in1_cursor = write_row(new_in1_sheet0, in1, in1_idcard_value.index(id)+1, new_in1_cursor)
        new_in2_cursor = write_row(new_in2_sheet0, in2, in2_idcard_value.index(id)+1, new_in2_cursor)
        # new_out_cursor = write_row(new_out_sheet, out, out_idcard_value.index(id)+1, new_out_cursor)

    # new_out.save('new_out.xls')
    print("id done")

    


    # 第一轮比较完后剩余需要比对的行号
    in1_rest_rawx = [i for i, x in enumerate(in1_idcard_value) if x=='']
    in2_rest_rawx = [i for i, x in enumerate(in2_idcard_value) if x not in union_id]

    # 排除out与in1、in2 重合id
    in1_id_in_out = [x for x in in1_idcard_value if x in out_idcard_value]
    in2_id_in_out = [x for x in in2_idcard_value if x in out_idcard_value]
    in1_id_in_out_rawx = []
    in2_id_in_out_rawx = []
    for id in in1_id_in_out:
        in1_id_in_out_rawx.append(in1_idcard_value.index(id))
    for id in in2_id_in_out:
        in2_id_in_out_rawx.append(in2_idcard_value.index(id))
    in1_rest_rawx = [x for x in in1_rest_rawx if x not in in1_id_in_out_rawx]
    in2_rest_rawx = [x for x in in2_rest_rawx if x not in in2_id_in_out_rawx]

    # 求name交集
    in1_name_value = in1.col_values(0)[1:]
    in2_name_value = in2.col_values(1)[1:]
    in1_name_rest_value = [x for i, x in enumerate(in1_name_value) if i in in1_rest_rawx]
    in2_name_rest_value = [x for i, x in enumerate(in2_name_value) if i in in2_rest_rawx]
    union_name = [i1 for i1 in in1_name_rest_value if i1 in in2_name_rest_value]

    # name交集所在行号
    union_in1_name_rawx = []
    union_in2_name_rawx = []
    for name in union_name:
        in1_name_rawx = [i for i, x in enumerate(in1_name_value) if x==name]
        in2_name_rawx = [i for i, x in enumerate(in2_name_value) if x==name]
        union_1_raw_x = [x for x in in1_name_rawx if x in in1_rest_rawx]
        union_2_raw_x = [x for x in in2_name_rawx if x in in2_rest_rawx]

        if len(union_1_raw_x) > 0 and len(union_2_raw_x) > 0:
            union_in1_name_rawx += union_1_raw_x
            union_in2_name_rawx += union_2_raw_x

    
    new_in1_sheet1 = new_in1.add_sheet('sheet1')
    new_in1_cursor = write_row(new_in1_sheet1, in1, 0, 0)
    new_in1_cursor = 1
    new_in2_sheet1 = new_in2.add_sheet('sheet1')
    new_in2_cursor = write_row(new_in2_sheet1, in2, 0, 0)
    new_in2_cursor = 1
    for rawx in union_in1_name_rawx:
        new_in1_cursor = write_row(new_in1_sheet1, in1, rawx+1, new_in1_cursor)
    for rawx in union_in2_name_rawx:
        new_in2_cursor = write_row(new_in2_sheet1, in2, rawx+1, new_in2_cursor)

    new_in1.save('io/new_in1.xls')
    new_in2.save('io/new_in2.xls')