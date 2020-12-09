# coding=utf-8
import xlrd
import datetime
from datetime import date
import xlwt

def read_excel():
    # 打开文件
    wb1 = xlrd.open_workbook(r'changjing.xlsx')
    sheet1 = wb1.sheet_by_index(0)
    manager_dict = dict()
    for i in range(1, sheet1.nrows):
        manager_dict[sheet1.cell(i, 4).value] = sheet1.cell(i, 11).value
    # print(manager_dict['10000'])

    wb2 = xlrd.open_workbook(r'in.xlsx')
    sheet2 = wb2.sheet_by_index(0)
    # num_dict = dict()
    # for i in range(1, sheet2.nrows):
    #     num_dict[sheet2.cell(i, 3).value] = sheet2.cell(i, 9).value
    # print(num_dict)


    f = xlwt.Workbook()
    sheets = f.add_sheet('1', cell_overwrite_ok=True)
    wb = xlrd.open_workbook(r'tts.xlsx')
    sheet = wb.sheet_by_index(0)
    rowNum = sheet.nrows
    colNum = sheet.ncols
    print('行数：'+str(rowNum)+'\t'+'列数：'+str(colNum))
    count = 0
    fcount = 0
    tcount = 0
    # 获取第二行第一个值
    for i in range(1, rowNum):
        if sheet.cell(i, 4).value == sheet.cell(i, 5).value:
            count += 1
            print(sheet.cell(i, 4).value)
            for j in range(colNum):
                sheets.write(count, j, sheet.cell(i, j).value)
            if sheet.cell(i, 4).value in manager_dict:
                sheets.write(count, colNum, manager_dict[sheet.cell(i, 4).value])
            temp = 0
            for z in range(1, sheet2.nrows):
                if sheet.cell(i, 4).value == sheet2.cell(z, 3).value:
                    temp += 1
                    if temp == 1:
                        sheets.write(count, colNum + 1, sheet2.cell(z, 9).value)
                        # print(sheet.cell(i, 4).value+'--1--'+sheet2.cell(z, 9).value)
                    if temp > 1:
                        count += 1
                        print(sheet.cell(i, 4).value+'----'+sheet2.cell(z, 9).value)
                        for m in range(colNum):
                            sheets.write(count, m, sheet.cell(i, m).value)
                        sheets.write(count, colNum, manager_dict[sheet.cell(i, 4).value])
                        sheets.write(count, colNum+1, sheet2.cell(z, 9).value)
            tcount += temp
            saveexcel(f)
        else:
            fcount += 1
    print(count)
    print(fcount)
    print(tcount)
# 保存表格
def saveexcel(f):
    f.save(r'C:\Users\wwh05\Documents\工作资料\python\python-test\SAAS\TTS整理\result.xlsx')

if __name__ == '__main__':
    read_excel()

