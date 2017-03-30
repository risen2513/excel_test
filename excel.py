# -*- coding: utf-8 -*-
import xlrd
# import xlwt
from xlutils.copy import copy

def excel_job(src, target):
    src_book = xlrd.open_workbook(src)
    sh = src_book.sheet_by_name(u'テスト')
    # value = sh.cell_value(rowx=1, colx=0)

    oldWb = xlrd.open_workbook(target, formatting_info=True)
    shwt = oldWb.sheet_by_name(u'hello')
    newWb = copy(oldWb)
    sheet = newWb.get_sheet('hello')
    src_id = [[] for i in range(2)]
    for irow in range(sh.nrows):
        c_row = sh.row(irow)
        src_id[0].append(str(c_row[0].value))
        src_id[1].append(c_row[1].value)

    for jrow in range(shwt.nrows):
        d_row = shwt.row(jrow)
        for srci_row in range(len(src_id[0])):
            if str(src_id[0][srci_row]) == str(d_row[0].value):
                print str(src_id[0][srci_row])
                sheet.write(jrow, 1, src_id[1][srci_row])

    newWb.save(target)


if __name__ == '__main__':
    excel_job('src.xls', 'target.xls')
