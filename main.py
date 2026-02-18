import flet as ft
from fields import values
from tables import create_table
from utils import add_to_table_values, clear_fields


def main(page: ft.Page):
    page.title = 'Количество процедур по исследованиям'
    page.icon = ft.icons.WORK
    page.window.width = 750
    page.window.height = 950
    page.theme_mode = 'dark'

    create_table()

    date_picker = ft.DatePicker(on_change=lambda e: update_date_field(e))
    page.overlay.append(date_picker)

    date_field = ft.TextField(
        label='Дата',
        hint_text='Выбрать дату',
        width=200,
        read_only=True,
    )

    def update_date_field(e):
        if date_picker.value:
            date_field.value = date_picker.value.strftime('%d.%m.%Y')
            page.update()

    def open_date_picker(e):
        page.open(date_picker)

    def change_theme(e):
        page.theme_mode = 'light' if page.theme_mode == 'dark' else 'dark'
        page.update()

    date_picker_btn = ft.IconButton(
        icon=ft.icons.CALENDAR_MONTH,
        on_click=open_date_picker,
        tooltip='Выбрать дату'
    )

    save_btn = ft.ElevatedButton(
        'Сохранить',
        on_click=lambda e: add_to_table_values(e, page, date_field.value)
    )

    clear_page_btn = ft.ElevatedButton(
        text='Очистить',
        on_click=lambda e: clear_fields(e, page)
    )

    dark_mode_btn = ft.IconButton(ft.icons.SUNNY, on_click=change_theme)

    page.add(
        ft.Row([date_field, date_picker_btn],
               alignment=ft.MainAxisAlignment.CENTER)
    )

    for value in values:
        page.add(
            ft.Row(
                [
                    ft.Column([value[0]], width=250),
                    ft.Column([value[1]], width=130),
                    ft.Column([value[2]], width=130)
                ],
                alignment=ft.MainAxisAlignment.CENTER
            )
        )

    page.add(
        ft.Row([dark_mode_btn, save_btn, clear_page_btn],
               alignment=ft.MainAxisAlignment.CENTER)
    )

    print("✅ Интерфейс отрисован")


ft.app(target=main, view=ft.WEB_BROWSER, port=8550)
