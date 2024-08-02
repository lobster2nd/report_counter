import os

import openpyxl

file_path = 'отчёт.xlsx'

if not os.path.exists(file_path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    cell_names = [
        'Наименование',
        '1',
        'Всего',
        'ОГК',
        'Костно-мышечной системы',
        'Конечности',
        'Таза и тазобедренных суставов',
        'Шейные позвонки',
        'Грудные позвонки',
        'Поясничные позвонки',
        'Рёбра и грудина',
        'Черепа и челюстно-лицевой области',
        'Зубы',
        'Челюстей',
        'Околоносовых пазух',
        'Череп'
    ]
    cell = sheet['A1']
    cell.value = 'Количество процедур по исследованиям'
    start = 5
    for value in cell_names:
        cell = sheet['A' + str(start)]
        cell.value = value
        start += 1

    wb.save(file_path)

