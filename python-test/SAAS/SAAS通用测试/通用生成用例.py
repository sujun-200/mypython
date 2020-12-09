import xlrd
import xlwt

def read_excel():
    wb = xlrd.open_workbook(r'eg.xlsx')
    sheet1 = wb.sheet_by_index(0)

    doit_things = {}
    for i in range(1, sheet1.nrows):
        if sheet1.cell(i, 0).value != '' and sheet1.cell(i, 1).value !='':
            doit_things[sheet1.cell(i, 0).value] = sheet1.cell(i, 1).value
    print(doit_things)

    lists = sheet1.merged_cells[:]
    more_load_rows = {}
    more_load_things = []
    for i in sheet1.merged_cells:
        temp = i
        if temp[2] != 1 and temp[3] != 2:
            inx = lists.index(i)
            # print(inx)
            del lists[inx]
        else:
            more_load_rows[sheet1.cell(temp[0], 0).value] = [temp[0], temp[1]]
            more_load_things.append(sheet1.cell(temp[0], 0).value)
    print(lists)
    print(more_load_rows)
    print(more_load_things)
    for i1 in range(1,sheet1.nrows):
        if sheet1.cell(i1, 4) != 'end':
            print(sheet1.cell(i1, 3))
    readtree(sheet1)


def readtree(sheet1):
    print()








# 保存表格
def saveexcel(f):
    f.save(r'data.xlsx')


if __name__ == '__main__':
    read_excel()