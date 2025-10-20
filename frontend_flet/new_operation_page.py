import flet as ft


class NewOperationPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        text_field_style = {
            "border_color": "#424242",
            "border_radius": 25,
            "bgcolor": "#424242",
            "border_width": 1,
            "color": "white",
            "text_size": 12
        }

        # Стиль для выпадающих списков
        dropdown_style = {
            "border_color": "#424242",
            "border_radius": 25,
            "bgcolor": "#424242",
            "border_width": 1,
            "text_size": 16,
            "color": "white",  # Цвет текста в закрытом состоянии
            "width": 300,
            "focused_border_color": "#835DA3",
            "focused_bgcolor": "#835DA3",
        }

        # Стиль для кнопок
        button_style = ft.ButtonStyle(
            color="white",
            bgcolor="#424242",
            padding=20,
            shape=ft.RoundedRectangleBorder(radius=25),
            overlay_color="#835DA3",
        )

        # Функция для создания dropdown с цветным текстом выбранной опции
        def create_dropdown(label, options):
            dropdown = ft.Dropdown(
                label=label,
                **dropdown_style,
                options=[ft.dropdown.Option(option) for option in options],
            )

            # Обработчик изменения выбора
            def on_change(e):
                # При изменении выбора устанавливаем белый цвет текста
                dropdown.color = "white"
                dropdown.update()

            dropdown.on_change = on_change
            return dropdown

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                create_dropdown(
                                    "ОПЕРАЦИЯ",
                                    ["ПРИХОД", "РАСХОД"]
                                ),
                            ]),
                            expand=1,
                            padding=10
                        ),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("ВВЕДИТЕ СУММУ", size=16, weight=ft.FontWeight.BOLD, color="white"),
                                ft.TextField(**text_field_style),
                            ]),
                            expand=1,
                            padding=10
                        ),
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),

                    ft.Container(height=20),

                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                ft.Text("КОММЕНТАРИЙ", size=16, weight=ft.FontWeight.BOLD, color="white"),
                                ft.TextField(
                                    multiline=True,
                                    min_lines=8,
                                    max_lines=10,
                                    **text_field_style
                                ),
                            ]),
                            expand=1,
                            padding=10
                        ),

                        ft.Container(
                            content=ft.Column([
                                ft.Row([
                                    ft.Container(
                                        content=create_dropdown(
                                            "ТИП ОПЛАТЫ",
                                            ["Наличные", "Безнал", "Другое"]
                                        ),
                                        expand=1,
                                        padding=5
                                    ),
                                    ft.Container(
                                        content=create_dropdown(
                                            "КАТЕГОРИЯ",
                                            ["КЦ", "АРЕНДА", "РЕКЛАМА", "ЗП ОКЛАДНИКИ", "ЗП ПРОЦЕНТ",
                                             "КОМИССИЯ", "ДИЛЕРСТВО", "ДРУГОЕ"]
                                        ),
                                        expand=1,
                                        padding=5
                                    ),
                                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),

                                ft.Container(height=20),

                                # Кнопка ИСКЛЮЧИТЬ ОПЕРАЦИЮ
                                ft.Row([
                                    ft.TextButton(
                                        "ИСКЛЮЧИТЬ ОПЕРАЦИЮ",
                                        style=button_style
                                    ),
                                ], alignment=ft.MainAxisAlignment.CENTER),
                            ]),
                            expand=1,
                            padding=10
                        ),
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),

                    ft.Container(height=50),

                    # Кнопка ДОБАВИТЬ ОПЕРАЦИЮ
                    ft.Row([
                        ft.TextButton(
                            "ДОБАВИТЬ ОПЕРАЦИЮ",
                            style=button_style
                        ),
                    ], alignment=ft.MainAxisAlignment.CENTER)
                ]),
                padding=50,
                border=ft.border.all(1, "#424242"),
                border_radius=25,
                bgcolor="#2C2C2C",
                width=1500
            )
        ]
        self.spacing = 0
        self.alignment = ft.MainAxisAlignment.CENTER
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.expand = True