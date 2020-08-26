import xlrd, xlwt
from xlrd import *

import win32com.client
import csv
import sys
import shutil
import traceback

# D:\Python\excel\Scripts\Activate.ps1
# pyinstaller --onefile --nowindowed wjx_background.py

def get_table(excel_path, sheet=0):
    data = xlrd.open_workbook(excel_path)
    table = data.sheets()[sheet]
    # table = data.sheet_by_index(0)
    return table

def write_row(ws, sh, index, cursor):
    for i, v in enumerate(sh.row_values(index)):
        ws.write(cursor, i, v)
    cursor = cursor + 1
    return cursor

def check_and_clear(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)
    else:
        shutil.rmtree(dir)
        os.makedirs(dir)

downloads_labels = ['6、身份证/护照：',
                    '9、最高学历证书或职称证书：',
                    '10、最高学位证书：',
                    '19、与现工作单位签订的劳动合同1',
                    '20、与现工作单位签订的劳动合同2',
                    '21、与现工作单位签订的劳动合同3',
                    '22、与现工作单位签订的劳动合同4',
                    '24、所在用人单位推荐意见']

if __name__ == '__main__':
    print("default:")
    print("   excel path: io/in.xlsx")
    print("   accessory roo: io/accessory")
    print("   save roo: io")

    in_path = input("excel path:")
    if in_path == "":
        in_path = 'io/in.xlsx' # in
    accessory_root = input("accessory root:")
    if accessory_root == "":
        accessory_root = 'io/accessory'
    save_root = input("save root:")
    if save_root == "":
        save_root = "io"
    in_path = in_path.replace("\\", "/")
    accessory_root = accessory_root.replace("\\", "/") + "/"
    save_root = save_root.replace("\\", "/") + "/"
    
    try:
        in_table = get_table(in_path)
        accessory_list = os.listdir(accessory_root)
    except FileNotFoundError:
        input(traceback.print_exc())
        exit(0)
    labels = in_table.row_values(0)
    for row_id in range(1, in_table.nrows):
        in_row_value = in_table.row_values(row_id)
        name = in_row_value[7]
        print("proccessing: {}".format(name))
        id = in_row_value[10]
        idx = int(in_row_value[0])
        root = save_root + name + "_" + id + "/"
        check_and_clear(root)

        # new excel
        new_book = xlwt.Workbook()
        new_sheet0 = new_book.add_sheet('sheet0')
        new_cursor = write_row(new_sheet0, in_table, row_id, 0)
        new_book.save(root+'complete_data.xls')

        # accessory
        idx_path = []
        for path in accessory_list:
            if "序号"+str(idx) in path:
                # idx_path.append(path)
                shutil.copyfile(accessory_root+path, root+path)


        # # print "downloading with requests"
        # for label in downloads_labels:
        #     url_id = labels.index(label)
        #     url = in_row_value[url_id]
        #     # r = requests.get(url)
        #     # save_name = label.replace("/", "_")
        #     # with open(root+save_name, "wb") as f:
        #     #     f.write(r.content)
        #     WJX_Tomas.login(url, "plzzbrcb", "rcb5644952")
    print("done")
    input("任意键结束...")
