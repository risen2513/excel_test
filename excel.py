import xlrd
import xlwt


def excel_job(src,target):

    src_book = xlrd.open_workbook(src)
    sh = src_book.sheet_by_index(0)
    value = sh.cell_value(rowx=1, colx=0)

    target_book = xlwt.Workbook()
    sheet2 = target_book.get_sheet(0)
    sheet2.row(1).write(0, value)
    sheet2.flush_row_data()
    target_book.save(target)




if __name__ == '__main__':
    excel_job('src.xls', 'target')