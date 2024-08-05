import os

import openpyxl
from openpyxl.styles import Alignment

file_path = 'отчёт.xlsx'


if not os.path.exists(file_path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20

    cell_names = [
        'Наименование',
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

    cell_values = {
        'B4': 'Всего исследований',
        'C4': 'Всего снимков',
        'B5': '=B7 + B8 + B15',
        'C5': '=C7 + C8 + C15',
        'B8': '=B9 + B14',
        'C8': '=C9 + C14',
        'B15': '=SUM(B16:B19)',
        'C15': '=SUM(C16:C19)'
    }

    for cell, formula in cell_values.items():
        sheet[cell].value = formula

    for col in ['B', 'C']:
        for cell in sheet[col]:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    wb.save(file_path)
