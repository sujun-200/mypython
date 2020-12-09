# coding=utf-8
import xlrd
import datetime
from datetime import date
import xlwt

def read_excel():
    # sum=2
    # 打开文件
    wb = xlrd.open_workbook(r'C:\Users\wwh05\Documents\test\test.xlsx')
    # 读sheet name
    # s1 = wb.sheet_names()
    # print(s1)

    # 获取第一/2个sheet的内容
    sheet1 = wb.sheet_by_index(0)
    sheet2 = wb.sheet_by_index(1)

    # 行数和列数
    # rowNum = sheet1.nrows
    # colNum = sheet1.ncols
    # print('行数：'+str(rowNum)+'\t'+'列数：'+str(colNum))

    # 获取第二行第一个值
    # s = sheet1.cell(1, 0).value
    # print(s)
    # 获取获取第二行第一个值的类型；ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    # ty = sheet1.cell(1, 0).ctype
    # print(ty)

    one = {}
    two = {}
    f = xlwt.Workbook()
    sheet = wtexcelHeader(f)
    for x in range(1, sheet1.nrows):
        one['things'] = sheet1.cell(x, 0).value
        one['words'] = sheet1.cell(x, 1).value
        one['robots-words'] = sheet1.cell(x, 2).value
        print(one)
        # print(sheet2.nrows)
        for y in range(1, sheet2.nrows):
            two['things'] = sheet2.cell(y, 0).value
            two['words'] = sheet2.cell(y, 1).value
            two['robots-words'] = sheet2.cell(y, 2).value
            print(two)
            wtexcel(sheet, one, two, y*2+1)
            saveexcel(f)
            # two.clear()
        one.clear()

# 设置表格样式
def set_style(name, height, bold=False):
    # 初始化样式
    style = xlwt.XFStyle()

    # 创建字体
    font = xlwt.Font()
    font.bold = bold
    font.colour_index = 0
    font.height = height
    font.name = name
    style.font = font
    return style


# 写入表格头
def wtexcelHeader(f):

    sheet1 = f.add_sheet('用例', cell_overwrite_ok=True)

    rows = ['意图', '实际步骤', '实际结果', '预期结果', '是否通过', 'senderid', 'topicCode']

    for i in range(len(rows)):
        sheet1.write(0, i, rows[i], set_style("Time New Roman", 220, True))

    saveexcel(f)
    return sheet1

# 写入表格
def wtexcel(sheet, one, two, c):
    # 打开文件
    wb1 = xlrd.open_workbook(r'C:\Users\wwh05\Documents\test\data.xlsx')
    # 获取第一/2个sheet的内容
    sheet1 = wb1.sheet_by_index(0)
    # 行数和列数
    rowNum1 = sheet1.nrows
    print(rowNum1)
    print(two['words'])
    sheet.write_merge(rowNum1 + 1, rowNum1 + 2, 0, 0, one['things'] + ' - ' + two['things'], set_style("Time New Roman", 220, True))
    sheet.write(rowNum1 + 1, 1, one['words'], set_style("Time New Roman", 220, True))
    sheet.write(rowNum1 + 1, 3, one['robots-words'], set_style("Time New Roman", 220, True))
    sheet.write(rowNum1 + 2, 1, two['words'], set_style("Time New Roman", 220, True))
    sheet.write(rowNum1 + 2, 3, two['robots-words'], set_style("Time New Roman", 220, True))


# 保存表格
def saveexcel(f):
    f.save(r'C:\Users\wwh05\Documents\test\data.xlsx')

if __name__ == '__main__':
    read_excel()

