#импортируем необходимые библиотеки
from openpyxl import workbook
from openpyxl.styles import Font, Color, colors
from openpyxl import load_workbook

#для удобства указали название файла
filename = 'МойТест_3.xlsx'

#вытянули данные с документа
active_excel = load_workbook(filename=filename, data_only=True)

#делаем или смотрим активный лист
active_sheet = active_excel.active

#меняем наименование листа
active_sheet.title = 'MyGameDesign'

#нужно понять максимальный размер данных на листе

def max_min_data(x):
    if x == 'max_row':
        max_row = int(active_sheet.max_row)
        return max_row
    else:
        max_column = int(active_sheet.max_column)
        return max_column
    #print('СТРОК: ' + str(max_row) + '\nКОЛОНОК: ' + str(max_column))

#создадим список, который запишем в столбик
Spisok1 = [
    'ПОШЛЕМ ТРВ',
    1,
    2,
    3,
    4
]

Spisok2 = [
    'Фамилия',
    'Имя',
    'Отчество',
    'Год рождения',
    'Зарплата'
]

Spisok3 = [
    'Запишем',
    'все',
    'в',
    'одну',
    'строку'
]

#с помощью for запишем наши данные (по умолчанию в 1 колонке после записей
#либо с самого начала если документ пустой)
for row in range(len(Spisok1)):
    active_sheet.append([Spisok1[row]])

#с помощью for запишем наши данные
# - развернем в строку)
a = max_min_data('max_row')
for i in range (len(Spisok3)):
    _= active_sheet.cell(column= i + 1, row=a + 1, value=Spisok3[i])


#запишем данные в столбец после max_column - данных
b = max_min_data('max_column')
for i in range (len(Spisok2)):
    _= active_sheet.cell(column= b + 1, row=i + 1, value=Spisok2[i])

#запишем 2 списка сразу в 2 колонки
if active_excel.sheetnames == 'MyGameDesign_2':
    pass
else:
    active_sheet2 = active_excel.create_sheet(title='MyGameDesign_2') #создали 2 лист в документе

Spisok4 = [
    1,
    2,
    3,
    4,
    5
]
Spisok5 = [
    1,
    2,
    3,
    4,
    5
]
S_45 = [Spisok4, Spisok5]

for i in range (len(Spisok4)):
    for j in range (len(S_45)):
        _=active_sheet2.cell(column=j+2, row=i+1, value=S_45[j][i])

#применение форматирование
style_1 = Font(name='Calibri', color=colors.BLUE,
               bold=True, size=12, underline='double')
style_2 = Font(name='Calibri', color=colors.RED,
               bold=True, size=10, underline='single')
style_3 = Font(name='Calibri', color=colors.RED,
               bold=True, size=14, underline='double')

#отформатируем шапку
a1 = active_sheet2['A1']
b1 = active_sheet2['B1']
a1.font = style_1
b1.font = style_1

#далее отформатируем остальную часть таблицы
for i in range(2,6):
    a = active_sheet2['B' + str(i)]
    b = active_sheet2['C' + str(i)]
    a.font = style_2
    b.font = style_2

#применем формулу СУММА
active_sheet2["A6"] = 'СУММА:'
active_sheet2["B6"] = '=SUM(B1:B5)'
active_sheet2["C6"] = '=SUM(C1:C5)'

a6 = active_sheet2['A6']
a6.font = style_3
# print(active_sheet2["A6"])

active_excel.save('МойТест_3.xlsx') #сохраняем все изменения
print('ИЗМЕНЕНИЯ СОХРАНЕНЫ')
