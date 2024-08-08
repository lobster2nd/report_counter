import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment


TOTAL_TABLE_VALUES = {    # Дописать формулы заполнения всех ячеек
    'A1': 'Количество процедур по исследованиям',
    'A5': 'Наименование',
    'A6': 'Всего',
    'A7': 'ОГК',
    'A8': 'Костно-мышечной системы',
    'A9': 'Конечности',
    'A10': 'Таза и тазобедренных суставов',
    'A11': 'Шейные позвонки',
    'A12': 'Грудные позвонки',
    'A13': 'Поясничные позвонки',
    'A14': 'Рёбра и грудина',
    'A15': 'Черепа и челюстно-лицевой области',
    'A16': 'Зубы',
    'A17': 'Челюстей',
    'A18': 'Околоносовых пазух',
    'A19': 'Череп',
    'A20': 'Брюшная полость',
    'A22': 'Сохранено: ',
    'B5': 'Всего исследований',
    'C5': 'Всего снимков',
    'B6': '=B7 + B8 + B15',
    'C6': '=C7 + C8 + C15',
    'B8': '=SUM(B9:B14)',
    'C8': '=SUM(C9:C14)',
    'B15': '=SUM(B16:B19)',
    'C15': '=SUM(C16:C19)',
    'B22': f'{datetime.now().strftime("%d-%m-%y %H:%M")}'
}

file_path = f'Отчёт_по_исследованиям_{datetime.now().year}_год.xlsx'


def create_month_template(month_name, start_row):
    return {
        f'A{start_row + 22}': month_name,
        f'A{start_row + 24}': 'Наименование',
        f'A{start_row + 25}': 'Всего',
        f'A{start_row + 26}': 'ОГК',
        f'A{start_row + 27}': 'Костно-мышечной системы',
        f'A{start_row + 28}': 'Конечности',
        f'A{start_row + 29}': 'Таза и тазобедренных суставов',
        f'A{start_row + 30}': 'Шейные позвонки',
        f'A{start_row + 31}': 'Грудные позвонки',
        f'A{start_row + 32}': 'Поясничные позвонки',
        f'A{start_row + 33}': 'Рёбра и грудина',
        f'A{start_row + 34}': 'Черепа и челюстно-лицевой области',
        f'A{start_row + 35}': 'Зубы',
        f'A{start_row + 36}': 'Челюстей',
        f'A{start_row + 37}': 'Околоносовых пазух',
        f'A{start_row + 38}': 'Череп',
        f'A{start_row + 39}': 'Брюшная полость',
        f'A{start_row + 41}': 'Сохранено: ',
        f'B{start_row + 24}': 'Всего исследований',
        f'C{start_row + 24}': 'Всего снимков',
        f'B{start_row + 25}': f'=B{start_row + 26} + B{start_row + 27} + B{start_row + 34} + B{start_row + 39}',
        f'C{start_row + 25}': f'=C{start_row + 26} + C{start_row + 27} + C{start_row + 34} + C{start_row + 39}',
        f'B{start_row + 27}': f'=SUM(B{start_row + 28}:B{start_row + 33})',
        f'C{start_row + 27}': f'=SUM(C{start_row + 28}:C{start_row + 33})',
        f'B{start_row + 34}': f'=SUM(B{start_row + 35}:B{start_row + 38})',
        f'C{start_row + 34}': f'=SUM(C{start_row + 35}:C{start_row + 38})',
        f'B{start_row + 41}': f'{datetime.now().strftime("%d-%m-%y %H:%M")}'
    }


months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
CELL_VALUES = {}
start_from_row = 1

for i in months:
    CELL_VALUES.update(create_month_template(i, start_from_row))
    start_from_row += 22


def create_table():
    """Создаёт шаблон таблицы отчёта"""
    if not os.path.exists(file_path):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.column_dimensions['A'].width = 40
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 20

        for cell, formula in CELL_VALUES.items():
            sheet[cell].value = formula

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        #
        # for row in sheet.iter_rows(min_row=30, max_row=45, min_col=1,
        #                            max_col=3):
        #     for cell in row:
        #         cell.border = thin_border
        #         if cell.value is None:
        #             cell.value = 0
        #         if cell.row in [31, 33, 40, 45, 53, 55, 62, 67]:
        #             cell.font = Font(bold=True)

        for row in sheet.iter_rows(min_row=5, max_row=282,
                                   min_col=1, max_col=3):
            pass

        for col in ['B', 'C']:
            for cell in sheet[col]:
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')

        wb.save(file_path)
