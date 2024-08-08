import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment


TOTAL_TABLE_VALUES = {
    'A1': 'Количество процедур по исследованиям',
    'B1': f'{datetime.now().year} год',
    'A3': 'Наименование',
    'A4': 'Всего',
    'B4': '=B26+B48+B70+B92+B114+B136+B158+B180+B202+B224+B246+B268',
    'C4': '=C26+C48+C70+C92+C114+C136+C158+C180+C202+C224+C246+C268',
    'A5': 'ОГК',
    'B5': '=B27+B49+B71+B93+B115+B137+B159+B181+B203+B225+B247+B269',
    'C5': '=C27+C49+C71+C93+C115+C137+C159+C181+C203+C225+C247+C269',
    'A6': 'Костно-мышечной системы',
    'B6': '=SUM(B7:B12)',
    'C6': '=SUM(C7:C12)',
    'A7': 'Конечности',
    'B7': '=B29+B51+B73+B95+B117+B139+B161+B183+B205+B227+B249+B271',
    'C7': '=C29+C51+C73+C95+C117+C139+C161+C183+C205+C227+C249+C271',
    'A8': 'Таза и тазобедренных суставов',
    'B8': '=B30+B52+B74+B96+B118+B140+B162+B184+B206+B228+B250+B272',
    'C8': '=C30+C52+C74+C96+C118+C140+C162+C184+C206+C228+C250+C272',
    'A9': 'Шейные позвонки',
    'B9': '=B31+B53+B75+B97+B119+B141+B163+B185+B207+B229+B251+B273',
    'C9': '=C31+C53+C75+C97+C119+C141+C163+C185+C207+C229+C251+C273',
    'A10': 'Грудные позвонки',
    'B10': '=B32+B54+B76+B98+B120+B142+B164+B186+B208+B230+B252+B274',
    'C10': '=C32+C54+C76+C98+C120+C142+C164+C186+C208+C230+C252+C274',
    'A11': 'Поясничные позвонки',
    'B11': '=B33+B55+B77+B99+B121+B143+B165+B187+B209+B231+B253+B275',
    'C11': '=C33+C55+C77+C99+C121+C143+C165+C187+C209+C231+C253+C275',
    'A12': 'Рёбра и грудина',
    'B12': '=B34+B56+B78+B100+B122+B144+B166+B188+B210+B232+B254+B276',
    'C12': '=C34+C56+C78+C100+C122+C144+C166+C188+C210+C232+C254+C276',
    'A13': 'Черепа и челюстно-лицевой области',
    'B13': '=SUM(B14:B17)',
    'C13': '=SUM(C14:C17)',
    'A14': 'Зубы',
    'B14': '=B36+B58+B80+B102+B124+B146+B168+B190+B212+B234+B256+B278',
    'C14': '=C36+C58+C80+C102+C124+C146+C168+C190+C212+C234+C256+C278',
    'A15': 'Челюстей',
    'B15': '=B37+B59+B81+B103+B125+B147+B169+B191+B213+B235+B257+B279',
    'C15': '=C37+C59+C81+C103+C125+C147+C169+C191+C213+C235+C257+C279',
    'A16': 'Околоносовых пазух',
    'B16': '=B38+B60+B82+B104+B126+B148+B170+B192+B214+B236+B258+B280',
    'C16': '=C38+C60+C82+C104+C126+C148+C170+C192+C214+C236+C258+C280',
    'A17': 'Череп',
    'B17': '=B39+B61+B83+B105+B127+B149+B171+B193+B215+B237+B259+B281',
    'C17': '=C39+C61+C83+C105+C127+C149+C171+C193+C215+C237+C259+C281',
    'A18': 'Брюшная полость',
    'B18': '=B40+B62+B84+B106+B128+B150+B172+B194+B216+B238+B260+B282',
    'C18': '=C40+C62+C84+C106+C128+C150+C172+C194+C216+C238+C260+C282',
    'A20': 'Сохранено: ',
    'B3': 'Всего исследований',
    'C3': 'Всего снимков',
    'B20': f'{datetime.now().strftime("%d-%m-%y %H:%M")}'
}

file_path = f'Отчёт_по_исследованиям_{datetime.now().year}_год.xlsx'


def create_month_template(month_name, start_row):
    """Возвращает шаблон таблицы на месяц"""
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

        for cell, formula in {**CELL_VALUES, **TOTAL_TABLE_VALUES}.items():
            sheet[cell].value = formula

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

        skip_rows = []
        for start in range(19, 262, 22):
            skip_rows += list(range(start, start + 6))

        bold_rows = months + ['Всего', 'Костно-мышечной системы',
                              'Черепа и челюстно-лицевой области',
                              'Брюшная полость']

        for row in sheet.iter_rows(min_row=3, max_row=282,
                                   min_col=1, max_col=3):
            for cell in row:
                if row[0].value in bold_rows:
                    cell.font = Font(bold=True)
                if row[0].row in skip_rows:
                    continue
                cell.border = thin_border
                if cell.value is None:
                    cell.value = 0

        for col in ['B', 'C']:
            for cell in sheet[col]:
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')

        wb.save(file_path)
