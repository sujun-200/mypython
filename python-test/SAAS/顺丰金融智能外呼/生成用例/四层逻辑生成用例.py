# coding=utf-8
import xlrd
import datetime
from datetime import date
import xlwt

def read_excel():
    # sum=2
    # 打开文件
    wb = xlrd.open_workbook(r'test1.xlsx')
    # 读sheet name
    # s1 = wb.sheet_names()
    # print(s1)

    # 获取第一/2个sheet的内容
    sheet1 = wb.sheet_by_index(0)
    sheet2 = wb.sheet_by_index(1)
    sheet3 = wb.sheet_by_index(2)
    sheet4 = wb.sheet_by_index(2)


    one = {}
    two = {}
    three = {}
    four = {}
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
            flag = 0
            for z in range(1, sheet3.nrows):
                three['things'] = sheet3.cell(z, 0).value
                three['words'] = sheet3.cell(z, 1).value
                three['robots-words'] = sheet3.cell(z, 2).value
                print(three)
                if three['things'].find(two['things']) != -1:
                    print(three['things']+"-----"+two['things'])
                    for h in range(1, sheet4.nrows):
                        four['things'] = sheet4.cell(h, 0).value
                        four['words'] = sheet4.cell(h, 1).value
                        four['robots-words'] = sheet4.cell(h, 2).value
                        print(four)
                        if four['things'].find(two['things']) != -1:
                            print(four['things'] + "-----" + three['things'])

                            wtexcel3(sheet, one, two, three, z * 3 + 1)
                        else:
                            print(four['things'] + "--2---" + three['things'])
                            flag = 1
                            # break
                        saveexcel(f)
                    wtexcel3(sheet, one, two, three, z*3+1)
                else:
                    print(three['things'] + "--2---" + two['things'])
                    flag = 1
                    # break
                saveexcel(f)
            if flag == 1:
                wtexcel2(sheet, one, two, y * 2 + 1)
                saveexcel(f)
            # two.clear()
        # one.clear()

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
def wtexcel2(sheet, one, two, c):
    # 打开文件
    wb1 = xlrd.open_workbook(r'C:\Users\wwh05\Documents\test\data1.xlsx')
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


# 写入表格
def wtexcel3(sheet, one, two, three, c):
    # 打开文件
    wb1 = xlrd.open_workbook(r'C:\Users\wwh05\Documents\test\data1.xlsx')
    # 获取第一/2个sheet的内容
    sheet1 = wb1.sheet_by_index(0)
    # 行数和列数
    rowNum1 = sheet1.nrows
    print(rowNum1)
    # print(two['words'])
    if three['things'].find(two['things']) != -1:
        if two['words'] == 'signal=silence':
            two['robots-words'] = one['robots-words']
        if three['words'] == 'signal=silence':
            three['robots-words'] = two['robots-words']

        sheet.write_merge(rowNum1 + 1, rowNum1 + 3, 0, 0, one['things'] + ' - ' + two['things'] + ' - ' + three['things'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 1, 1, one['words'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 1, 3, one['robots-words'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 2, 1, two['words'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 2, 3, two['robots-words'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 3, 1, three['words'], set_style("Time New Roman", 220, True))
        sheet.write(rowNum1 + 3, 3, three['robots-words'], set_style("Time New Roman", 220, True))


# 保存表格
def saveexcel(f):
    f.save(r'C:\Users\wwh05\Documents\test\data1.xlsx')

if __name__ == '__main__':
    read_excel()

