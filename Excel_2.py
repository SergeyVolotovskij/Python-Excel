import xlrd
import openpyxl

#создадим файл
#wb = openpyxl.Workbook()
#wb.save('МойТест_2.xlsx')

location = 'МойТест_2.xlsx'

wb = xlrd.open_workbook(location) #открыли
#print(wb)

sheet = wb.sheet_by_index(0) #присвоили 1-й лист в переменную
#print(sheet)

#извлекаем наименования колонок
for i in range(sheet.ncols):
    print(sheet.cell_value(0,i))

#выведем первую строку на первом листе
print(sheet.row_values(1))

