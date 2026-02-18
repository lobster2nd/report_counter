from datetime import datetime

import flet as ft
import openpyxl

from fields import values
from tables import months, file_path, CELL_VALUES


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


def get_month_from_date(date_str):
    """Извлекает название месяца из строки даты (дд.мм.гггг)"""
    try:
        if not date_str:
            return None
        date_obj = datetime.strptime(date_str, '%d.%m.%Y')
        month_num = date_obj.month - 1  # 0-indexed
        return months[month_num]
    except (ValueError, IndexError) as e:
        print(f"Ошибка преобразования даты: {e}")
        return None


def validate_data(date_str, page):
    if not date_str:
        show_info('Выберите дату', page)
        return False

    month = get_month_from_date(date_str)
    if not month:
        show_info('Некорректный формат даты', page)
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
    """Обновляет значения в таблице"""
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
            scan_cnt = int(scan_cnt_str) if scan_cnt_str else 0
            img_cnt = int(img_cnt_str) if img_cnt_str else 0
        except ValueError:
            continue

        for key, formula in CELL_VALUES.items():
            if section_name in formula \
                    and 'области' not in formula \
                    and start <= int(key[1:]) <= start + 13:
                cell_b = sheet['B' + key[1:]]
                cell_c = sheet['C' + key[1:]]

                if cell_b.value is None:
                    cell_b.value = 0
                if cell_c.value is None:
                    cell_c.value = 0

                if action == 'add':
                    cell_b.value += scan_cnt
                    cell_c.value += img_cnt
                elif action == 'rewrite':
                    cell_b.value = scan_cnt
                    cell_c.value = img_cnt

    # Обновляем дату сохранения
    for key in CELL_VALUES.keys():
        if CELL_VALUES[key] == month:
            row = int(key[1:]) + 19
            sheet[f'B{row}'] = datetime.now().strftime("%d-%m-%y %H:%M")
            break

    wb.save(file_path)

    for v in values:
        v[1].value = ''
        v[2].value = ''


def add_to_table_values(e, page, date_str):
    """Сохранить значения в таблицу"""
    if validate_data(date_str=date_str, page=page):
        month = get_month_from_date(date_str)
        update_table_values(action='add', month=month)
        show_info('Значения сохранены', page)