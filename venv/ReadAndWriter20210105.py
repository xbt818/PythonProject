# -*- coding: utf-8 -*-
import os
from openpyxl import load_workbook
from openpyxl import Workbook

def in_sheet_out(sheet_name, need_sheet):
    global nowTitle
    nowTitle=''
    for i in range(0, 15):
        if sheet_name[i:i + 2] == need_sheet:
            print('i=' + str(i) + '-'+now_file_name+'-' + sheet_name[i:i + 8])
            nowTitle = sheet_name[i:i + 8]
    return nowTitle

#输入一个'文件全路径'，输出遍历的所有sheet标签名,读取原始档，输出数据到out.xlsx
def read_one_file(file_name,need_sheet):
    global now_file_name
    wb_read = load_workbook(filename=file_name, read_only=True)
    for sheet in wb_read:
        now_file_name = os.path.basename(file_name)[::-1][5:20][::-1]
        in_sheet_out(sheet.title, need_sheet)
        if nowTitle[0:2]==need_sheet:
            now_file_name=os.path.basename(file_name)[::-1][5:20][::-1]
            ws=wb_read[sheet.title]
            for row in range(4, 35):
                ws_A=ws.cell(row=row, column=1).value
                ws_H = ws.cell(row=row, column=8).value
                if ws_H==None:
                    ws_H=0
                print(str(ws_A)+'-'+str(ws_H)+'-'+now_file_name+'-'+nowTitle,nowTitle+str(ws_A))
                input_list=[str(ws_A),'','','','','','',str(ws_H),now_file_name,nowTitle,nowTitle+str(ws_A)]
                ws_out.append(input_list)

#输入一个str(文件夹路径)，遍历文件夹，输出数组names_list[文件全路径名]
def read_name(file_dir):
    global names_dir_list
    read_path=''
    names_dir_list=[]
    for root, dirs, names_list in os.walk(file_dir):
        read_path=root
    for name in names_list:
        names_dir_list.append(read_path + name)
    print(names_list)
    return names_dir_list,names_list


if __name__ == "__main__":
    wb_out = Workbook()
    ws_out = wb_out.active
    read_name(r'PATH')
    for every in names_dir_list:
        read_one_file(every,'A')
    wb_out.save(r'PATH')