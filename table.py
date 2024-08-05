import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment

file_path = 'отчёт.xlsx'


def create_table():
    """Создаёт шаблон таблицы отчёта"""
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
            'Череп',
            'Брюшная полость',
            '',
            'Сохранено: '
        ]

        start = 5
        for value in cell_names:
            cell = sheet['A' + str(start)]
            cell.value = value
            start += 1

        cell_values = {
            'A1': 'Количество процедур по исследованиям',
            'B4': 'Всего исследований',
            'C4': 'Всего снимков',
            'B5': '=B7 + B8 + B15',
            'C5': '=C7 + C8 + C15',
            'B8': '=B9 + B14',
            'C8': '=C9 + C14',
            'B15': '=SUM(B16:B20)',
            'C15': '=SUM(C16:C20)',
            'B22': f'{datetime.now().strftime("%d-%m-%y %H:%M")}'
        }

        for cell, formula in cell_values.items():
            sheet[cell].value = formula

        for row in sheet.iter_rows(min_row=5, max_row=20, min_col=2,
                                   max_col=3):
            for cell in row:
                if cell.value is None:
                    cell.value = 0

        for col in ['B', 'C']:
            for cell in sheet[col]:
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')

        wb.save(file_path)


def modify_table(values: dict):
    """Внесение изменений в таблицу"""
    pass
