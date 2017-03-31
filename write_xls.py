# -*- coding: utf-8 -*-
import xlrd
from xlutils.copy import copy

def write(target, open_src, sheet_name):
    oldWb = xlrd.open_workbook(target, formatting_info=True)
    shwt = oldWb.sheet_by_name(sheet_name)
    newWb = copy(oldWb)
    sheet = newWb.get_sheet(sheet_name)
    src_id = [[] for i in range(2)]
    li = range(open_src.nrows)
    del li[0]

    for irow in li:
        c_row = open_src.row(irow)
        src_id[0].append(c_row[0].value)
        src_id[1].append(c_row[2].value)

    for jrow in range(shwt.nrows):
        d_row = shwt.row(jrow)
        for srci_row in range(len(src_id[0])):
            if src_id[0][srci_row] == d_row[6].value:
                if len(src_id[1][srci_row]) > 0:
                    sheet.write(jrow, 8, src_id[1][srci_row])
    newWb.save(target)
