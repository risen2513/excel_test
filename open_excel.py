# -*- coding: utf-8 -*-
import xlrd

def read(src, sheet_name):
    src_book = xlrd.open_workbook(src, sheet_name)
    sh = src_book.sheet_by_name(sheet_name)
    return sh
