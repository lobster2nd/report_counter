import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from fields import values

RESEARCH_TYPES = [v[0].value for v in values]

COLUMN_CONFIG = {
    'A': {'name': 'Дата', 'width': 18, 'type': 'date'},
}

# Начинаем с индекса 1, так как колонка A уже занята
col_index = 1
RESEARCH_COL_MAP = {}

for research_name in RESEARCH_TYPES:
    # Правильный расчет индексов колонок
    scan_col = get_column_letter(
        col_index + 1)  # +1 потому что начинаем со следующей после A
    img_col = get_column_letter(col_index + 2)

    COLUMN_CONFIG[scan_col] = {
        'name': research_name,
        'width': 10,
        'type': 'scan',
        'group': research_name
    }
    COLUMN_CONFIG[img_col] = {
        'name': research_name,
        'width': 10,
        'type': 'img',
        'group': research_name
    }

    RESEARCH_COL_MAP[research_name] = (scan_col, img_col)
    col_index += 2  # Увеличиваем на 2 для следующей пары колонок

# Колонки для итогов
TOTAL_SCAN_COL = get_column_letter(col_index + 1)
TOTAL_IMG_COL = get_column_letter(col_index + 2)

COLUMN_CONFIG[TOTAL_SCAN_COL] = {
    'name': 'Итого',
    'width': 12,
    'type': 'total_scan',
    'group': 'Итого'
}
COLUMN_CONFIG[TOTAL_IMG_COL] = {
    'name': 'Итого',
    'width': 12,
    'type': 'total_img',
    'group': 'Итого'
}

MONTHS = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
]


def get_days_in_month(year: int, month_num: int) -> int:
    month_days = {
        1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30,
        7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31
    }

    if month_num == 2:
        if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
            return 29

    return month_days.get(month_num, 30)


def get_month_sheet_name(year: int, month: str) -> str:
    return f"{month}_{year}"


def get_year_file_path(year: int) -> str:
    return f'Отчёт_по_исследованиям_{year}_год.xlsx'


def get_column_index(col_letter: str) -> int:
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def create_month_sheet(wb, year: int, month: str, month_num: int):
    sheet_name = get_month_sheet_name(year, month)

    if sheet_name in wb.sheetnames:
        return wb[sheet_name]

    sheet = wb.create_sheet(sheet_name)

    # Устанавливаем ширину колонок
    for col, config in COLUMN_CONFIG.items():
        sheet.column_dimensions[col].width = config['width']

    # Стили
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4',
                              fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center',
                                 wrap_text=True)

    subheader_font = Font(bold=True, size=10)
    subheader_fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5',
                                 fill_type='solid')

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    thick_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thick'),
        bottom=Side(style='thin')
    )

    # Сортируем колонки для правильного отображения
    sorted_cols = sorted(COLUMN_CONFIG.keys(), key=get_column_index)

    # Сначала заполняем все ячейки значениями
    # Заполняем первую строку
    for col in sorted_cols:
        config = COLUMN_CONFIG[col]
        cell = sheet[f'{col}1']

        if config['type'] == 'date':
            cell.value = 'Дата'
        elif config['type'] in ['scan', 'img']:
            cell.value = config['group']
        elif config['type'] in ['total_scan', 'total_img']:
            if config['type'] == 'total_scan':
                cell.value = 'Итого за день'

        # Применяем стили к ячейке
        if cell.value:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

    # Заполняем вторую строку
    for col, config in COLUMN_CONFIG.items():
        cell = sheet[f'{col}2']

        if config['type'] == 'date':
            cell.value = ''
        elif config['type'] == 'scan':
            cell.value = 'иссл.'
        elif config['type'] == 'img':
            cell.value = 'снимки'
        elif config['type'] == 'total_scan':
            cell.value = 'иссл.'
        elif config['type'] == 'total_img':
            cell.value = 'снимки'

        # Применяем стили ко второй строке
        if config['type'] != 'date':
            cell.font = subheader_font
            cell.fill = subheader_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        else:
            cell.border = thin_border

    # Теперь объединяем ячейки
    # Объединяем ячейки даты (A1:A2)
    sheet.merge_cells('A1:A2')

    # Объединяем ячейки групп исследований
    groups = {}
    for col in sorted_cols:
        config = COLUMN_CONFIG[col]
        if config['type'] in ['scan', 'img']:
            if config['group'] not in groups:
                groups[config['group']] = []
            groups[config['group']].append(col)

    for group_name, cols in groups.items():
        if cols and len(cols) > 1:
            start_col = cols[0]
            end_col = cols[-1]
            sheet.merge_cells(f'{start_col}1:{end_col}1')

    # Объединяем итоговые ячейки
    sheet.merge_cells(f'{TOTAL_SCAN_COL}1:{TOTAL_IMG_COL}1')

    # Заполняем дни месяца
    days_in_month = get_days_in_month(year, month_num)
    scan_cols = [col for col, cfg in COLUMN_CONFIG.items() if
                 cfg.get('type') == 'scan']
    img_cols = [col for col, cfg in COLUMN_CONFIG.items() if
                cfg.get('type') == 'img']

    for day in range(1, days_in_month + 1):
        row = day + 2

        # Дата
        date_cell = sheet[f'A{row}']
        date_cell.value = f'{day:02d}.{month_num:02d}.{year}'
        date_cell.border = thin_border
        date_cell.alignment = Alignment(horizontal='center', vertical='center')

        # Применяем границы ко всем ячейкам строки
        for col in COLUMN_CONFIG.keys():
            cell = sheet[f'{col}{row}']
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # Формулы для итогов
        if scan_cols:
            scan_formula = f'=SUM({",".join([f"{col}{row}" for col in scan_cols])})'
            sheet[f'{TOTAL_SCAN_COL}{row}'].value = scan_formula
        if img_cols:
            img_formula = f'=SUM({",".join([f"{col}{row}" for col in img_cols])})'
            sheet[f'{TOTAL_IMG_COL}{row}'].value = img_formula

    # Добавляем строку с итогами за месяц
    total_row = days_in_month + 4

    # Толстая граница перед итогами
    for col in COLUMN_CONFIG.keys():
        if col != 'A':  # Не применяем толстую границу к колонке с датами
            sheet[f'{col}{total_row - 1}'].border = thick_border

    # Заголовок итоговой строки
    sheet[f'A{total_row}'].value = 'Итого за месяц:'
    sheet[f'A{total_row}'].font = Font(bold=True)
    sheet[f'A{total_row}'].alignment = Alignment(horizontal='right')
    sheet[f'A{total_row}'].border = thick_border

    sheet.row_dimensions[total_row].height = 20

    # Формулы для итогов за месяц
    for col in COLUMN_CONFIG.keys():
        if col in [TOTAL_SCAN_COL, TOTAL_IMG_COL]:
            formula = f'=SUM({col}3:{col}{days_in_month + 2})'
            sheet[f'{col}{total_row}'].value = formula
            sheet[f'{col}{total_row}'].font = Font(bold=True)
            sheet[f'{col}{total_row}'].fill = PatternFill(start_color='FFC000',
                                                          end_color='FFC000',
                                                          fill_type='solid')
        elif col != 'A':
            formula = f'=SUM({col}3:{col}{days_in_month + 2})'
            sheet[f'{col}{total_row}'].value = formula
            sheet[f'{col}{total_row}'].font = Font(bold=True)

        cell = sheet[f'{col}{total_row}']
        cell.border = thick_border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Закрепляем первые две строки
    sheet.freeze_panes = 'A3'

    return sheet


def create_table(year: int = None):
    if year is None:
        year = datetime.now().year

    file_path = get_year_file_path(year)

    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
    else:
        wb = openpyxl.Workbook()
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']

    for month_num, month_name in enumerate(MONTHS, start=1):
        create_month_sheet(wb, year, month_name, month_num)

    wb.save(file_path)
    print(f"✅ Файл журнала создан: {file_path}")
    return file_path
