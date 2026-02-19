import flet as ft
from fields import values
from utils import add_to_table_values, clear_fields


def main(page: ft.Page):
    page.title = 'Количество процедур по исследованиям'
    page.icon = ft.icons.WORK
    page.window.width = 550
    page.window.height = 950
    page.theme_mode = 'dark'
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    date_picker = ft.DatePicker(on_change=lambda e: update_date_field(e))
    page.overlay.append(date_picker)

    date_field = ft.TextField(
        label='Дата',
        hint_text='Выбрать дату',
        width=200,
        read_only=True,
        border_radius=8,
        text_size=14,
    )

    def update_date_field(e):
        if date_picker.value:
            date_field.value = date_picker.value.strftime('%d.%m.%Y')
            page.update()

    def open_date_picker(e):
        page.open(date_picker)

    def change_theme(e):
        page.theme_mode = 'light' if page.theme_mode == 'dark' else 'dark'
        # Обновляем иконку кнопки темы
        theme_btn.icon = ft.icons.SUNNY if page.theme_mode == 'dark' else ft.icons.DARK_MODE
        page.update()

    date_picker_btn = ft.IconButton(
        icon=ft.icons.CALENDAR_MONTH,
        on_click=open_date_picker,
        tooltip='Выбрать дату',
        icon_size=28,
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
        width=500,
        content=ft.Column(
            spacing=2,
            controls=[],
            scroll=ft.ScrollMode.AUTO,  # Добавляем скролл для таблицы
            height=600,  # Фиксированная высота для скролла
        )
    )

    # Заголовки таблицы
    header_row = ft.Container(
        content=ft.Row(
            [
                ft.Container(
                    content=ft.Text(
                        "Наименование",
                        size=14,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=260,
                    padding=ft.padding.only(left=15),
                ),
                ft.Container(
                    content=ft.Text(
                        "Исследований",
                        size=14,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=120,  # Увеличиваем ширину для соответствия
                    alignment=ft.alignment.center,
                ),
                ft.Container(
                    content=ft.Text(
                        "Снимков",
                        size=14,
                        weight=ft.FontWeight.BOLD,
                        color=ft.colors.PRIMARY,
                    ),
                    width=120,  # Увеличиваем ширину для соответствия
                    alignment=ft.alignment.center,
                ),
            ],
            alignment=ft.MainAxisAlignment.START,
            spacing=0,  # Убираем spacing
        ),
        bgcolor=ft.colors.with_opacity(0.1, ft.colors.PRIMARY),
        border_radius=ft.border_radius.only(top_left=8, top_right=8),
        padding=ft.padding.symmetric(vertical=10, horizontal=0),
        width=500,
    )

    table_container.content.controls.append(header_row)

    # Строки таблицы
    for i, value in enumerate(values):
        category_name = value[0].value
        scan_field = value[1]
        img_field = value[2]

        # Настраиваем поля ввода
        scan_field.width = 90  # Увеличиваем ширину
        scan_field.height = 34
        scan_field.text_align = ft.TextAlign.CENTER
        scan_field.border_radius = 6
        scan_field.label = None
        scan_field.hint_text = "0"
        scan_field.content_padding = ft.padding.symmetric(horizontal=6,
                                                          vertical=4)
        scan_field.text_size = 13

        img_field.width = 90  # Увеличиваем ширину
        img_field.height = 34
        img_field.text_align = ft.TextAlign.CENTER
        img_field.border_radius = 6
        img_field.label = None
        img_field.hint_text = "0"
        img_field.content_padding = ft.padding.symmetric(horizontal=6,
                                                         vertical=4)
        img_field.text_size = 13

        # Чередование фона для строк
        bg_color = ft.colors.with_opacity(0.03,
                                          ft.colors.PRIMARY) if i % 2 == 0 else None

        row_container = ft.Container(
            content=ft.Row(
                [
                    ft.Container(
                        content=ft.Text(
                            category_name,
                            size=13,
                            overflow=ft.TextOverflow.ELLIPSIS,
                        ),
                        width=260,
                        padding=ft.padding.only(left=15),
                    ),
                    ft.Container(
                        content=scan_field,
                        width=120,  # Увеличиваем ширину
                        alignment=ft.alignment.center,
                    ),
                    ft.Container(
                        content=img_field,
                        width=120,  # Увеличиваем ширину
                        alignment=ft.alignment.center,
                    ),
                ],
                alignment=ft.MainAxisAlignment.START,
                spacing=0,  # Убираем spacing
            ),
            bgcolor=bg_color,
            padding=ft.padding.symmetric(vertical=4, horizontal=0),
            width=500,
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
        icon_size=22,
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
                    padding=ft.padding.symmetric(horizontal=16, vertical=8),
                ),
            ),
            ft.ElevatedButton(
                text="Очистить",
                on_click=lambda e: clear_fields(e, page),
                icon=ft.icons.CLEAR,
                style=ft.ButtonStyle(
                    shape=ft.RoundedRectangleBorder(radius=6),
                    padding=ft.padding.symmetric(horizontal=16, vertical=8),
                ),
            ),
            theme_btn,
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=12,
    )

    # Нижняя панель
    page.add(
        ft.Container(
            content=button_row,
            margin=ft.margin.only(top=20),
            alignment=ft.alignment.center,
        )
    )


ft.app(target=main, view=ft.WEB_BROWSER, port=8550)
