# -*- coding: utf-8 -*-
import datetime
import open_excel
import write_xls

def searchandcopy(src, target):
    starttime = datetime.datetime.now()
    open_src = open_excel.read(src, u'Sheet2')
    write_xls.write(target, open_src, u'台帳')
    endtime = datetime.datetime.now()
    print "用时：", (endtime - starttime).seconds, "秒"

if __name__ == '__main__':
    searchandcopy('./testdata/zsrc.xlsx', './testdata/ztarget.xls')
