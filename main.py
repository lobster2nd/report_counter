import flet as ft
from fields import values
from utils import add_to_table_values, clear_fields


def main(page: ft.Page):
    page.title = 'Количество процедур по исследованиям'
    page.icon = ft.icons.LOCAL_HOSPITAL
    page.window.width = 650
    page.window.height = 1050
    page.theme_mode = 'dark'
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    date_picker = ft.DatePicker(on_change=lambda e: update_date_field(e))
    page.overlay.append(date_picker)

    date_field = ft.TextField(
        label='Дата',
        hint_text='Выбрать дату',
        width=250,
        read_only=True,
        border_radius=8,
        text_size=16,
    )

    def update_date_field(e):
        if date_picker.value:
            date_field.value = date_picker.value.strftime('%d.%m.%Y')
            page.update()

    def open_date_picker(e):
        page.open(date_picker)

    def change_theme(e):
        page.theme_mode = 'light' if page.theme_mode == 'dark' else 'dark'
        theme_btn.icon = ft.icons.SUNNY if page.theme_mode == 'dark' else ft.icons.DARK_MODE
        page.update()

    date_picker_btn = ft.IconButton(
        icon=ft.icons.CALENDAR_MONTH,
        on_click=open_date_picker,
        tooltip='Выбрать дату',
        icon_size=32,
    )

    # Верхняя панель с датой
    page.add(
        ft.Container(
            content=ft.Row(
                [date_field, date_picker_btn],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=10,
            ),
            margin=ft.margin.only(bottom=20),
        )
    )

    # Контейнер для таблицы с фиксированной шириной
    table_container = ft.Container(
        width=600,
        content=ft.Column(
            spacing=2,
            controls=[],
            scroll=ft.ScrollMode.AUTO,
            height=700,
        )
    )

    # Заголовки таблицы
    header_row = ft.Container(
        content=ft.Row(
            [
                ft.Container(
                    content=ft.Text(
                        "Наименование",
                        size=18,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=300,
                    padding=ft.padding.only(left=15),
                ),
                ft.Container(
                    content=ft.Text(
                        "Исследований",
                        size=18,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=150,
                    alignment=ft.alignment.center,
                ),
                ft.Container(
                    content=ft.Text(
                        "Снимков",
                        size=18,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=150,
                    alignment=ft.alignment.center,
                ),
            ],
            alignment=ft.MainAxisAlignment.START,
            spacing=0,
        ),
        bgcolor=ft.colors.with_opacity(0.1, ft.colors.PRIMARY),
        border_radius=ft.border_radius.only(top_left=8, top_right=8),
        padding=ft.padding.symmetric(vertical=12, horizontal=0),
        width=600,
    )

    table_container.content.controls.append(header_row)

    # Строки таблицы
    for i, value in enumerate(values):
        category_name = value[0].value
        scan_field = value[1]
        img_field = value[2]

        # Настраиваем поля ввода
        scan_field.width = 110
        scan_field.height = 40
        scan_field.text_align = ft.TextAlign.CENTER
        scan_field.border_radius = 6
        scan_field.label = None
        scan_field.hint_text = "0"
        scan_field.content_padding = ft.padding.symmetric(horizontal=8, vertical=6)
        scan_field.text_size = 16

        img_field.width = 110
        img_field.height = 40
        img_field.text_align = ft.TextAlign.CENTER
        img_field.border_radius = 6
        img_field.label = None
        img_field.hint_text = "0"
        img_field.content_padding = ft.padding.symmetric(horizontal=8, vertical=6)
        img_field.text_size = 16

        # Чередование фона для строк
        bg_color = ft.colors.with_opacity(0.03,
                                          ft.colors.PRIMARY) if i % 2 == 0 else None

        row_container = ft.Container(
            content=ft.Row(
                [
                    ft.Container(
                        content=ft.Text(
                            category_name,
                            size=16,
                            overflow=ft.TextOverflow.ELLIPSIS,
                        ),
                        width=300,
                        padding=ft.padding.only(left=15),
                    ),
                    ft.Container(
                        content=scan_field,
                        width=150,
                        alignment=ft.alignment.center,
                    ),
                    ft.Container(
                        content=img_field,
                        width=150,
                        alignment=ft.alignment.center,
                    ),
                ],
                alignment=ft.MainAxisAlignment.START,
                spacing=0,
            ),
            bgcolor=bg_color,
            padding=ft.padding.symmetric(vertical=6, horizontal=0),
            width=600,
            border_radius=4,
        )

        table_container.content.controls.append(row_container)

    # Добавляем таблицу на страницу
    page.add(
        ft.Container(
            content=table_container,
            alignment=ft.alignment.center,
        )
    )

    # Кнопки
    theme_btn = ft.IconButton(
        icon=ft.icons.SUNNY if page.theme_mode == 'dark' else ft.icons.DARK_MODE,
        on_click=change_theme,
        tooltip='Сменить тему',
        icon_size=28,
    )

    button_row = ft.Row(
        [
            ft.ElevatedButton(
                text="Сохранить",
                on_click=lambda e: add_to_table_values(e, page,
                                                       date_field.value),
                icon=ft.icons.SAVE,
                style=ft.ButtonStyle(
                    shape=ft.RoundedRectangleBorder(radius=6),
                    padding=ft.padding.symmetric(horizontal=20, vertical=12),
                ),
            ),
            ft.ElevatedButton(
                text="Очистить",
                on_click=lambda e: clear_fields(e, page),
                icon=ft.icons.CLEAR,
                style=ft.ButtonStyle(
                    shape=ft.RoundedRectangleBorder(radius=6),
                    padding=ft.padding.symmetric(horizontal=20, vertical=12),
                ),
            ),
            theme_btn,
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=15,
    )

    # Нижняя панель
    page.add(
        ft.Container(
            content=button_row,
            margin=ft.margin.only(top=25),
            alignment=ft.alignment.center,
        )
    )


ft.app(target=main, view=ft.WEB_BROWSER, port=8550)
