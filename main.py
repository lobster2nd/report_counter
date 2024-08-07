import flet as ft

from fields import header, values
from utils import create_table, add_to_table_values, rewrite_table_values, \
    clear_fields


def main(page: ft.Page):
    page.window.width = 650
    page.window.height = 900
    page.theme_mode = 'dark'

    create_table()

    def change_theme(e):
        """Светлая/тёмная тема"""
        page.theme_mode = 'light' if page.theme_mode == 'dark' else 'dark'
        page.update()

    add_btn = ft.ElevatedButton('Прибавить',
                                on_click=lambda e: add_to_table_values(e, page)
                                )
    rewrite_btn = ft.ElevatedButton('Сохранить новые значения',
                                    on_click=lambda e: rewrite_table_values(e, page)
                                    )
    clear_page_btn = ft.ElevatedButton(text='Очистить',
                                       on_click=lambda e: clear_fields(e, page)
                                       )
    dark_mode_btn = ft.IconButton(ft.icons.SUNNY, on_click=change_theme)

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

    page.add(ft.Row([dark_mode_btn, add_btn, rewrite_btn,
                     clear_page_btn],
                    alignment=ft.MainAxisAlignment.CENTER))


ft.app(target=main)
