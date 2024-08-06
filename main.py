import flet as ft
import openpyxl

from fields import header, values
from table import create_table, file_path, CELL_VALUES


def main(page: ft.Page):
    page.window.width = 650
    page.window.height = 900

    create_table()

    def add_to_table_values(e):
        """Прибавить переданные значения к таблице"""

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

                    if cell_b.value is None:
                        cell_b.value = 0
                    if cell_c.value is None:
                        cell_c.value = 0

                    cell_b.value += scan_cnt
                    cell_c.value += img_cnt

        wb.save(file_path)

    def rewrite_table_values(e):
        """Внести новые значения в таблицу, старые удаляются"""

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

                    if cell_b.value is None:
                        cell_b.value = 0
                    if cell_c.value is None:
                        cell_c.value = 0

                    cell_b.value = scan_cnt
                    cell_c.value = img_cnt

        wb.save(file_path)

    page.add(
        ft.Row(
            [header], alignment=ft.MainAxisAlignment.CENTER,
        )
    )

    for value in values:
        page.add(
            ft.Row(
                [
                    ft.Column([value[0]], width=200),
                    ft.Column([value[1]], width=200),
                    ft.Column([value[2]], width=150)
                ], alignment=ft.MainAxisAlignment.CENTER
            )
        )

    add_btn = ft.ElevatedButton('Прибавить',
                                on_click=add_to_table_values
                                )
    rewrite_btn = ft.ElevatedButton('Сохранить новые значения',
                                    on_click=rewrite_table_values
                                    )

    page.add(ft.Row([add_btn, rewrite_btn],
                    alignment=ft.MainAxisAlignment.CENTER))


ft.app(target=main)
