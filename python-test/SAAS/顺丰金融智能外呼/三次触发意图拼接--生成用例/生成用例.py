# coding=utf-8
import xlrd
import datetime
from datetime import date
import xlwt

def read_excel():
    # sum=2
    # 打开文件
    wb = xlrd.open_workbook(r'test2.xlsx')
    # 读sheet name
    # s1 = wb.sheet_names()

    # 获取第一/2个sheet的内容
    sheet1 = wb.sheet_by_index(1)

    one = {}
    two = {}
    three = {}
    four = {}
    f = xlwt.Workbook()
    sheet = wtexcelHeader(f)
    num = 0
    for x in range(0, sheet1.nrows):
        if sheet1.cell(x, 3).value == 1:
            for y in range(0, sheet1.nrows):
                if sheet1.cell(y, 3).value == 1:
                    for h in range(0, sheet1.nrows):
                        sheet.write(num*6, 0, '开场白')
                        sheet.write(num*6, 1, 'signal=newCall')
                        sheet.write(num*6, 3, '您好，请问是张三先生吗？')
                        sheet.write(num*6+1, 0, '本人接听-确认本人')
                        sheet.write(num*6+1, 1, '是的')
                        sheet.write(num*6+1, 3, '我是顺丰金融客服，您本月一共有5笔借据的贷款已逾期，总还款金额1000元，请您今晚六点前尽快把欠款处理一下好吧？')
                        sheet.write(num * 6+2, 0, sheet1.cell(x, 0).value)
                        sheet.write(num * 6+2, 1, sheet1.cell(x, 1).value)
                        sheet.write(num * 6+2, 3, sheet1.cell(x, 2).value)
                        sheet.write(num * 6 + 3, 0, sheet1.cell(y, 0).value)
                        sheet.write(num * 6 + 3, 1, sheet1.cell(y, 1).value)
                        sheet.write(num * 6 + 3, 3, sheet1.cell(y, 2).value)
                        sheet.write(num * 6 + 4, 0, sheet1.cell(h, 0).value)
                        sheet.write(num * 6 + 4, 1, sheet1.cell(h, 1).value)
                        sheet.write(num * 6 + 4, 3, sheet1.cell(h, 2).value)
                        num += 1
                        saveexcel(f)
                else:
                    sheet.write(num * 6, 0, '开场白')
                    sheet.write(num * 6, 1, 'signal=newCall')
                    sheet.write(num * 6, 3, '您好，请问是张三先生吗？')
                    sheet.write(num * 6 + 1, 0, '本人接听-确认本人')
                    sheet.write(num * 6 + 1, 1, '是的')
                    sheet.write(num * 6 + 1, 3, '我是顺丰金融客服，您本月一共有5笔借据的贷款已逾期，总还款金额1000元，请您今晚六点前尽快把欠款处理一下好吧？')
                    sheet.write(num * 6 + 2, 0, sheet1.cell(x, 0).value)
                    sheet.write(num * 6 + 2, 1, sheet1.cell(x, 1).value)
                    sheet.write(num * 6 + 2, 3, sheet1.cell(x, 2).value)
                    sheet.write(num * 6 + 3, 0, sheet1.cell(y, 0).value)
                    sheet.write(num * 6 + 3, 1, sheet1.cell(y, 1).value)
                    sheet.write(num * 6 + 3, 3, sheet1.cell(y, 2).value)
                    num += 1
                    saveexcel(f)
        else:
            sheet.write(num * 6, 0, '开场白')
            sheet.write(num * 6, 1, 'signal=newCall')
            sheet.write(num * 6, 3, '您好，请问是张三先生吗？')
            sheet.write(num * 6 + 1, 0, '本人接听-确认本人')
            sheet.write(num * 6 + 1, 1, '是的')
            sheet.write(num * 6 + 1, 3, '我是顺丰金融客服，您本月一共有5笔借据的贷款已逾期，总还款金额1000元，请您今晚六点前尽快把欠款处理一下好吧？')
            sheet.write(num * 6 + 2, 0, sheet1.cell(x, 0).value)
            sheet.write(num * 6 + 2, 1, sheet1.cell(x, 1).value)
            sheet.write(num * 6 + 2, 3, sheet1.cell(x, 2).value)
            num += 1
            saveexcel(f)


# 写入表格头
def wtexcelHeader(f):

    sheet1 = f.add_sheet('用例', cell_overwrite_ok=True)

    # rows = ['意图', '实际步骤', '实际结果', '预期结果', '是否通过', 'senderid', 'topicCode']

    # for i in range(len(rows)):
        # sheet1.write(0, i, rows[i], set_style("Time New Roman", 220, True))

    # saveexcel(f)
    return sheet1



# 保存表格
def saveexcel(f):
    f.save(r'C:\Users\wwh05\Documents\test\data.xlsx')

if __name__ == '__main__':
    read_excel()

