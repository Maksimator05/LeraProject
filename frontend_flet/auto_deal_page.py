import flet as ft


class AutoDealPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        # Стиль для текстовых полей
        text_field_style = {
            "border_color": "#424242",
            "border_radius": 25,
            "bgcolor": "#424242",
            "border_width": 1
        }

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Container(
                        content=ft.Column([
                            ft.Row([
                                ft.Column([
                                    ft.Text("МАРКА", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                    ft.TextField(width=375, **text_field_style),
                                ], expand=True),
                                ft.Column([
                                    ft.Text("ГОД ВЫПУСКА", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                    ft.TextField(width=375, **text_field_style),
                                ], expand=True),
                                ft.Column([
                                    ft.Text("VIN", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                    ft.TextField(width=450, **text_field_style),
                                ], expand=True),
                            ], spacing=20),
                        ]),
                        padding=30,
                        border=ft.border.all(1, "#2C2C2C"),
                        border_radius=15,
                        bgcolor="#2C2C2C",
                        width=1600,
                    ),

                    ft.Container(height=20),

                    ft.Row([
                        ft.Column([
                            ft.Container(
                                content=ft.Column([
                                    ft.Column([
                                        ft.Text("ЦЕНА С ОПЦИЯМИ", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                        ft.TextField(width=375, **text_field_style),
                                    ], expand=True),
                                    ft.Column([
                                        ft.Text("ЗАКУПОЧНАЯ СТОИМОСТЬ", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                        ft.TextField(width=375, **text_field_style),
                                    ], expand=True),
                                    ft.Column([
                                        ft.Text("ДОПОЛНИТЕЛЬНЫЕ РАСХОДЫ", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                        ft.TextField(width=375, **text_field_style),
                                    ], expand=True),
                                ]),
                                padding=25,
                                border=ft.border.all(1, "#2C2C2C"),
                                border_radius=4,
                                bgcolor="#2C2C2C"
                            ),

                            ft.Container(height=30),

                            ft.Container(
                                content=ft.Row([
                                    ft.Container(width=500),
                                    ft.ElevatedButton(
                                        "ДОБАВИТЬ СДЕЛКУ",
                                        style=ft.ButtonStyle(
                                            color="white",
                                            bgcolor="#424242",
                                            padding=25,
                                        ),
                                        width=300
                                    ),
                                ]),
                            ),
                        ], expand=True),

                        ft.Container(
                            content=ft.Column([
                                ft.Text("КОММЕНТАРИЙ", size=12, weight=ft.FontWeight.BOLD, color="#424242"),
                                ft.TextField(
                                    multiline=True,
                                    min_lines=8,
                                    max_lines=8,
                                    width=500,
                                    **text_field_style
                                ),
                            ]),
                            width=500,
                            padding=20,
                            border=ft.border.all(1, "#2C2C2C"),
                            border_radius=8,
                            bgcolor="#2C2C2C",
                            margin=ft.margin.only(left=10, top=-100)  # Поднят на уровень с закупочной стоимостью
                        )
                    ], spacing=0)
                ]),
                padding=40,
                border=ft.border.all(1, "#424242"),
                border_radius=10,
                bgcolor="#2C2C2C",
                width=1500
            )
        ]
        self.spacing = 0
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER