import flet as ft
import openpyxl

from table import create_table, file_path, CELL_VALUES

SCAN_CNT = 'Кол-во исследований'
IMG_CNT = 'Кол-во снимков'


def main(page: ft.Page):

    create_table()

    def modify_table(e):
        """Внесение изменений в таблицу"""

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

    page.window.width = 650
    page.window.height = 900

    header = ft.Text(value='Количество процедур по исследованиям')

    chest = ft.Text(value='ОГК')
    chest_scan_cnt = ft.TextField(label=SCAN_CNT)
    chest_img_cnt = ft.TextField(label=IMG_CNT)

    limbs = ft.Text(value='Конечности')
    limbs_scan_cnt = ft.TextField(label=SCAN_CNT)
    limbs_img_cnt = ft.TextField(label=IMG_CNT)

    pelvis = ft.Text(value='Таза и тазобедренных суставов')
    pelvis_scan_cnt = ft.TextField(label=SCAN_CNT)
    pelvis_img_cnt = ft.TextField(label=IMG_CNT)

    neck = ft.Text(value='Шейные позвонки')
    neck_scan_cnt = ft.TextField(label=SCAN_CNT)
    neck_img_cnt = ft.TextField(label=IMG_CNT)

    gop = ft.Text(value='Грудные позвонки')
    gop_scan_cnt = ft.TextField(label=SCAN_CNT)
    gop_img_cnt = ft.TextField(label=IMG_CNT)

    pkop = ft.Text(value='Поясничные позвонки')
    pkop_scan_cnt = ft.TextField(label=SCAN_CNT)
    pkop_img_cnt = ft.TextField(label=IMG_CNT)

    ribs = ft.Text(value='Рёбра и грудина')
    ribs_scan_cnt = ft.TextField(label=SCAN_CNT)
    ribs_img_cnt = ft.TextField(label=IMG_CNT)

    teeth = ft.Text(value='Зубы')
    teeth_scan_cnt = ft.TextField(label=SCAN_CNT)
    teeth_img_cnt = ft.TextField(label=IMG_CNT)

    jaw = ft.Text(value='Челюстей')
    jaw_scan_cnt = ft.TextField(label=SCAN_CNT)
    jaw_img_cnt = ft.TextField(label=IMG_CNT)

    ppn = ft.Text(value='Околоносовых пазух')
    ppn_scan_cnt = ft.TextField(label=SCAN_CNT)
    ppn_img_cnt = ft.TextField(label=IMG_CNT)

    scull = ft.Text(value='Череп')
    scull_scan_cnt = ft.TextField(label=SCAN_CNT)
    scull_img_cnt = ft.TextField(label=IMG_CNT)

    abdomen = ft.Text(value='Брюшная полость')
    abdomen_scan_cnt = ft.TextField(label=SCAN_CNT)
    abdomen_img_cnt = ft.TextField(label=IMG_CNT)

    values = [
        [chest, chest_scan_cnt, chest_img_cnt],
        [limbs, limbs_scan_cnt, limbs_img_cnt],
        [pelvis, pelvis_scan_cnt, pelvis_img_cnt],
        [neck, neck_scan_cnt, neck_img_cnt],
        [gop, gop_scan_cnt, gop_img_cnt],
        [pkop, pkop_scan_cnt, pkop_img_cnt],
        [ribs, ribs_scan_cnt, ribs_img_cnt],
        [teeth, teeth_scan_cnt, teeth_img_cnt],
        [jaw, jaw_scan_cnt, jaw_img_cnt],
        [ppn, ppn_scan_cnt, ppn_img_cnt],
        [scull, scull_scan_cnt, scull_img_cnt],
        [abdomen, abdomen_scan_cnt, abdomen_img_cnt]
    ]

    page.add(
        ft.Row(
            [header],
            alignment=ft.MainAxisAlignment.CENTER,
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

    save_btn = ft.ElevatedButton('Сохранить',
                                 on_click=modify_table
                                 )

    page.add(ft.Row([save_btn], alignment=ft.MainAxisAlignment.CENTER))


ft.app(target=main)
