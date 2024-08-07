import openpyxl
import flet as ft

from fields import values
from tables import file_path, CELL_VALUES


def clear_fields(e, page):
    """Очистить поля"""
    for v in values:
        v[1].value = ''
        v[2].value = ''
    page.update()


def update_table_values(action):
    """Обновляет или перезаписывает значения в таблице"""

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    for val in values:
        section_name = val[0].value
        scan_cnt_str = val[1].value
        img_cnt_str = val[2].value

        try:
            scan_cnt = int(scan_cnt_str)
            img_cnt = int(img_cnt_str)
        except ValueError:
            continue

        for key, formula in CELL_VALUES.items():
            if section_name in formula and 'области' not in formula:
                cell_b = sheet['B' + key[1:]]
                cell_c = sheet['C' + key[1:]]

                if cell_b.value is None: cell_b.value = 0
                if cell_c.value is None: cell_c.value = 0

                match action:
                    case 'add':
                        cell_b.value += scan_cnt
                        cell_c.value += img_cnt
                    case 'rewrite':
                        cell_b.value = scan_cnt
                        cell_c.value = img_cnt

    wb.save(file_path)

    for v in values:
        v[1].value = ''
        v[2].value = ''


def add_to_table_values(e, page):
    """Прибавить переданные значения к таблице"""
    page.snack_bar = ft.SnackBar(ft.Text('Сохранено'))
    page.snack_bar.open = True

    update_table_values(action='add')
    page.update()


def rewrite_table_values(e, page):
    """Внести новые значения в таблицу, старые удаляются"""
    page.snack_bar = ft.SnackBar(ft.Text('Сохранено'))
    page.snack_bar.open = True

    update_table_values(action='rewrite')
    page.update()
