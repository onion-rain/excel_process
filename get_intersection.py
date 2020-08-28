# coding=utf-8

import xlrd, xlwt
from xlutils.copy import copy
from xlrd import *

import win32com.client
import csv
import sys
import traceback

# D:\Python\excel\Scripts\Activate.ps1
# pyinstaller --onefile --nowindowed get_intersection.py


def get_r_sheet(excel_path, sheet=0):
    data = xlrd.open_workbook(excel_path)
    sheet = data.sheets()[sheet]
    # table = data.sheet_by_index(0)
    return data, sheet

# def get_w_table(excel_path, sheet=0):
#     data = xlwt.Workbook(encoding = 'ascii')
#     sheet = data.sheets()[sheet]
#     return sheet

def write_row(ws, sh, index, cursor):
    for i, v in enumerate(sh.row_values(index)):
        ws.write(cursor, i, v)
    cursor = cursor + 1
    return cursor

def check(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)

def get_style():
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5 		 # 5 背景颜色为黄色
    #1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon,
    # 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray
    style = xlwt.XFStyle()
    style.pattern = pattern
    return style

if __name__ == '__main__':
    print("default:")
    print("   excel1 path: io/in1.xlsx")
    print("   excel2 path: io/in2.xlsx")
    print("   save root: io")
    print("   excel1 ID NO col: 0")
    print("   excel2 ID NO col: 0")
    in1_path = input("excel1 path:")
    if in1_path == "":
        in1_path = 'io/in1.xlsx' # in1
    in2_path = input("excel2 path:")
    if in2_path == "":
        in2_path = 'io/in2.xlsx' # in2
    in1_path = in1_path.replace("\\", "/")
    in2_path = in2_path.replace("\\", "/")
    try:
        in1_ws, in1 = get_r_sheet(in1_path)
        in2_ws, in2 = get_r_sheet(in2_path)
    except FileNotFoundError:
        input(traceback.print_exc())
        exit(0)
    in1_ws = copy(in1_ws)
    in1_w = in1_ws.get_sheet(0)
    in2_ws = copy(in2_ws)
    in2_w = in2_ws.get_sheet(0)

    save_root = input("save root:")
    if save_root == "":
        save_root = "io"
    save_root = save_root.replace("\\", "/") + "/"
    check(save_root)

    in1_col = input("excel1 ID NO col:")
    if in1_col == "":
        in1_col = '0' # in1 ID NO col
    in2_col = input("excel2 ID NO col:")
    if in2_col == "":
        in2_col = '0' # in2 ID NO col
    in1_col = int(in1_col)
    in2_col = int(in2_col)

    in1_idcard_value = in1.col_values(in1_col)[1:]
    in2_idcard_value = in2.col_values(in2_col)[1:]

    print("\nfound: 0", end="")

    # 仅凭id card求得交集
    union_id = [i1 for i1 in in1_idcard_value if i1 in in2_idcard_value]

    # 新建
    new_in1 = xlwt.Workbook()
    new_in1_sheet0 = new_in1.add_sheet('sheet0')
    new_in1_cursor = write_row(new_in1_sheet0, in1, 0, 0)
    new_in1_cursor = 1
    new_in2 = xlwt.Workbook()
    new_in2_sheet0 = new_in2.add_sheet('sheet0')
    new_in2_cursor = write_row(new_in2_sheet0, in2, 0, 0)
    new_in2_cursor = 1
    num = 1
    for id in union_id:
        print("\rfound: {}".format(num), end="")
        new_in1_cursor = write_row(new_in1_sheet0, in1, in1_idcard_value.index(id)+1, new_in1_cursor)
        new_in2_cursor = write_row(new_in2_sheet0, in2, in2_idcard_value.index(id)+1, new_in2_cursor)
        num += 1
        
        # 原表高亮
        style = get_style()
        in1_w.write(in1_idcard_value.index(id)+1, in1_col, id, style)
        in2_w.write(in2_idcard_value.index(id)+1, in2_col, id, style)
    

    new_in1.save(save_root + 'new_in1.xls')
    new_in2.save(save_root + 'new_in2.xls')
    in1_ws.save(save_root + 'highlight_in1.xls')
    in2_ws.save(save_root + 'highlight_in2.xls')

    print("\n\ndone")
    input("任意键结束...")