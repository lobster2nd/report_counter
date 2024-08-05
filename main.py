import flet as ft

SCAN_CNT = 'Кол-во исследований'
IMG_CNT = 'Кол-во снимков'


def main(page: ft.Page):

    page.window_width = 650
    page.window_height = 900

    header = ft.Text(value='Количество процедур по исследованиям')

    chest = ft.Text(value='ОГК')
    chest_scan_cnt = ft.TextField(value=SCAN_CNT)
    chest_img_cnt = ft.TextField(value=IMG_CNT)

    limbs = ft.Text(value='Конечности')
    limbs_scan_cnt = ft.TextField(value=SCAN_CNT)
    limbs_img_cnt = ft.TextField(value=IMG_CNT)

    pelvis = ft.Text(value='Таз и ТБС')
    pelvis_scan_cnt = ft.TextField(value=SCAN_CNT)
    pelvis_img_cnt = ft.TextField(value=IMG_CNT)

    neck = ft.Text(value='ШОП')
    neck_scan_cnt = ft.TextField(value=SCAN_CNT)
    neck_img_cnt = ft.TextField(value=IMG_CNT)

    gop = ft.Text(value='ГОП')
    gop_scan_cnt = ft.TextField(value=SCAN_CNT)
    gop_img_cnt = ft.TextField(value=IMG_CNT)

    pkop = ft.Text(value='ПКОП')
    pkop_scan_cnt = ft.TextField(value=SCAN_CNT)
    pkop_img_cnt = ft.TextField(value=IMG_CNT)

    ribs = ft.Text(value='Рёбра и грудина')
    ribs_scan_cnt = ft.TextField(value=SCAN_CNT)
    ribs_img_cnt = ft.TextField(value=IMG_CNT)

    teeth = ft.Text(value='Зубы')
    teeth_scan_cnt = ft.TextField(value=SCAN_CNT)
    teeth_img_cnt = ft.TextField(value=IMG_CNT)

    jaw = ft.Text(value='Челюстей')
    jaw_scan_cnt = ft.TextField(value=SCAN_CNT)
    jaw_img_cnt = ft.TextField(value=IMG_CNT)

    ppn = ft.Text(value='Околоносовых пазух')
    ppn_scan_cnt = ft.TextField(value=SCAN_CNT)
    ppn_img_cnt = ft.TextField(value=IMG_CNT)

    scull = ft.Text(value='Череп')
    scull_scan_cnt = ft.TextField(value=SCAN_CNT)
    scull_img_cnt = ft.TextField(value=IMG_CNT)

    abdomen = ft.Text(value='ОБП')
    abdomen_scan_cnt = ft.TextField(value=SCAN_CNT)
    abdomen_img_cnt = ft.TextField(value=IMG_CNT)

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

    page.add(ft.Row([ft.ElevatedButton('Сохранить')],
                    alignment=ft.MainAxisAlignment.CENTER))


ft.app(target=main)
