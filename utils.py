import openpyxl
import flet as ft

from fields import values
from tables import file_path, CELL_VALUES


def show_info(info, page):
    page.snack_bar = ft.SnackBar(ft.Text(f'{info}'))
    page.snack_bar.open = True
    page.update()


def clear_fields(e, page):
    """Очистить поля"""
    for v in values:
        v[1].value = ''
        v[2].value = ''
    page.update()


def validate_data(month, page):
    if month is None:
        show_info('Выберите месяц', page)
        return False
    for v in values:
        scan = v[1].value
        img = v[2].value

        if scan and img:
            try:
                int(scan)
                int(img)
            except ValueError:
                show_info('Значение должно быть целым числом', page)
                return False

        if (scan and not img) or (img and not scan):
            show_info('Введите количество исследований и снимков', page)
            return False
    return True


def update_table_values(action, month):
    """Обновляет или перезаписывает значения в таблице"""

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    start = 0

    for key, formula in CELL_VALUES.items():
        if formula == month:
            start = int(key[1:]) + 4
            break

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
            if section_name in formula \
                    and 'области' not in formula \
                    and start <= int(key[1:]) <= start + 13:
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


def add_to_table_values(e, page, month):
    """Прибавить переданные значения к таблице"""
    if validate_data(month=month, page=page):
        update_table_values(action='add', month=month)
        show_info('Значения добавлены', page)


def rewrite_table_values(e, page, month):
    """Внести новые значения в таблицу, старые удаляются"""
    if validate_data(month=month, page=page):

        update_table_values(action='rewrite', month=month)
        show_info('Значения обновлены', page)
