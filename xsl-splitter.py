#!/usr/bin/env python3

import os
import argparse
from itertools import groupby
from openpyxl import Workbook, load_workbook

# command line parsing
parser = argparse.ArgumentParser(description='Optional app description')
# Required positional argument
parser.add_argument('source_file', help='源数据文件名称')
parser.add_argument('col_name', nargs='?', default='汇算单位名称', help='分组列名称，缺省为“汇算单位名称”')
args = parser.parse_args()

def load_source_data(filename):
    workbook = load_workbook(filename=filename)
    sheet = workbook.active
    data = []
    for row in sheet.values:
        rowdata = []
        for value in row:
            rowdata.append(value)
        data.append(rowdata)
    return data

def get_col_index(line, pred):
    for i in range(len(line)):
        if line[i] and pred(line[i]):
            return i
    raise Exception("找不到指定的列")

def group_data(source_data):
    unit_name_col = get_col_index(source_data[0], lambda line:args.col_name in line)
    keyfunc = lambda r:r[unit_name_col]
    grouped_data = groupby(sorted(source_data[1:], key=keyfunc), keyfunc)
    target_data = {}
    for key, rows_gen in grouped_data:
        rows = []
        for row in rows_gen:
            rows.append(row)
        target_data[key] = rows
    return target_data

def write_one_target_xls(filename, header_line, rows):
    wb = Workbook()
    s = wb.active
    s.append(header_line)
    for r in rows:
        s.append(r)
    wb.save(filename=filename)

def write_target_xls(header_line, data):
    folder_name = "target/" + os.path.splitext(args.source_file)[0]
    os.makedirs(folder_name)
    for k in data.keys():
        print("正在写入", k, len(data[k]), "行数据")
        write_one_target_xls(folder_name + "/" + k + ".xlsx", header_line, data[k])

data = load_source_data(args.source_file)
write_target_xls(data[0], group_data(data))
