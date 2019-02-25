# -*- coding: utf-8 -*-
import xlrd
import xlwt

data = xlrd.open_workbook("17.xlsx")
table = data.sheets()[0]
nrows = table.nrows
print(nrows)

wbk = xlwt.Workbook()
sheet = wbk.add_sheet("sheet1")


insert_row_number = 1
for i in range(nrows):
    if i == 0 or i == 1:
        continue
    # print(i, table.row_values(i)[:])
    for j in range(2, 12):
        # print(table.row_values(i)[1], table.row_values(0)[j], table.row_values(i)[j])
        sheet.write(insert_row_number, 0, "2018-2019")
        sheet.write(insert_row_number, 1, u"上学期")
        sheet.write(insert_row_number, 2, "药学系")
        sheet.write(insert_row_number, 3, "17药剂1班")
        sheet.write(insert_row_number, 4, table.row_values(i)[1])
        sheet.write(insert_row_number, 5, table.row_values(0)[j])
        sheet.write(insert_row_number, 7, table.row_values(i)[j])
        sheet.write(insert_row_number, 8, "期末考试")
        insert_row_number = insert_row_number + 1
wbk.save("text.xls")