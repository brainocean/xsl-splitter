#!/usr/bin/env python3

from itertools import groupby
from openpyxl import Workbook, load_workbook

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
    raise Exception("Can't find needed column")

def group_data(source_data):
    unit_name_col = get_col_index(source_data[0], lambda line:"单位名称" in line)
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
    for k in data.keys():
        print("Writing", k, len(data[k]), "rows.")
        write_one_target_xls("target/"+k+".xlsx", header_line, data[k])

data = load_source_data("data.xlsx")
write_target_xls(data[0], group_data(data))
