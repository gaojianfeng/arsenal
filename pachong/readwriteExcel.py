# -*- coding: utf-8 -*-
import xdrlib, sys

import os
import xlrd
import xlwt
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

allset = set()
pricedit = {}
qualitydit = {}
commonnames = set()
vendors = set()
status = set()
specifics = set()
factors = set()
regions = set()
listregion = []
prices = set()


def open_excel(path):
    try:
        data = xlrd.open_workbook('D:/test/original/' + path.decode('utf-8'))
        return data
    except Exception, e:
        print str(e)


def excel_table_byname(path, f):
    data = open_excel(path)
    table = data.sheet_by_index(0)
    nrows = table.nrows  # 行数

    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            if (row[16].find(u'执行中') == -1):
                continue
            if (row[19] == '' or row[19] == '-'):
                continue
            allset.add(row[7] + '$' + row[0] + '$' + row[3] + '$' + str(row[4]))
            pricedit[row[7] + '$' + row[0] + '$' + row[3] + '$' + str(row[4]) + '$' + row[9]] = str(row[19])
            qualitydit[row[7] + '$' + row[0] + '$' + row[3] + '$' + str(row[4]) + '$' + row[9]] = row[17]

            regions.add(row[9])

    listregion = list(regions)
    listregion.sort()

    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;');  # 80% like

    # 生成第一lie
    sheet1.write(0, 0, u'生产企业', styleBlueBkg)
    sheet1.write(1, 0, u'通用名', styleBlueBkg)
    sheet1.write(2, 0, u'规格', styleBlueBkg)
    sheet1.write(3, 0, u'转换系数', styleBlueBkg)
    index = int(4)
    for region in listregion:
        sheet1.write(index, 0, region)
        index = index + 1
    colindex = int(1)
    listresult = list(allset)
    listresult.sort()
    for result in listresult:
        listvalue = result.split('$')
        rowindex = int(0)
        for value in listvalue:
            sheet1.write(rowindex, colindex, value, styleBlueBkg)
            sheet1.write(rowindex, colindex + 1, value, styleBlueBkg)
            rowindex = rowindex + 1
        colindex = colindex + 2
        if colindex >= 254:
            print path + "exceed 256 columns"
            break

    rows = int(4)
    for region in listregion:
        cols = int(1)
        for result in listresult:
            if result + '$' + region in pricedit:
                sheet1.write(rows, cols, pricedit[result + '$' + region])
                sheet1.write(rows, cols + 1, qualitydit[result + '$' + region])
                # else:
                # print 'the key not exist:'+result+','+region
            cols = cols + 2
            if cols >= 254:
                print  path + "exceed 256 columns"
                break
        rows = rows + 1


def eachFile():
    pathlist = set()
    pathDir = os.listdir(u'C:\\Users\\pc\\Desktop\\华招竞品- 原始数据')
    for allDir in pathDir:
        child = os.path.join('%s' % (allDir))
        pathlist.add(child)

    return pathlist


def main():
    f17 = xlwt.Workbook()  # 创建工作簿
    excel_table_byname('厄贝沙坦.xls', f17)
    f17.save(u'D:/test/result/厄贝沙坦.xls')  # 保存文件


if __name__ == "__main__":
    main()
