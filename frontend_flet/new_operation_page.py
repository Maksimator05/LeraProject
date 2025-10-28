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

        # Стиль для кнопок
        button_style = ft.ButtonStyle(
            color="white",
            bgcolor="#424242",
            padding=20,
            shape=ft.RoundedRectangleBorder(radius=25),
            overlay_color="#835DA3",
        )

        # Функция для создания кастомного dropdown через PopupMenuButton
        def create_custom_dropdown(label, options, default_value=None):
            if default_value is None and options:
                default_value = options[0]

            selected_value = ft.Text(
                value=default_value or "",
                color="white",
                size=16,
                expand=True
            )

            # Контейнер для отображения выбранного значения
            display_container = ft.Container(
                content=ft.Row([
                    selected_value,
                    ft.Text("▼", color="white", size=16),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                padding=ft.padding.symmetric(horizontal=20, vertical=15),
                border=ft.border.all(1, "#424242"),
                border_radius=25,
                bgcolor="#424242",
                width=300,
            )

            def on_item_selected(e):
                selected_value.value = e.control.text
                display_container.bgcolor = "#835DA3"
                display_container.border = ft.border.all(1, "#835DA3")
                display_container.update()
                selected_value.update()

            # Создаем PopupMenuButton - МИНИМАЛЬНАЯ ВЕРСИЯ
            popup_menu = ft.PopupMenuButton(
                content=display_container,
                items=[
                    ft.PopupMenuItem(
                        text=option,
                        on_click=on_item_selected,
                    ) for option in options
                ]
            )

            return ft.Column([
                ft.Text(label, size=12, color="white", weight=ft.FontWeight.BOLD),
                popup_menu,
            ], spacing=5)

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                create_custom_dropdown(
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
                                        content=create_custom_dropdown(
                                            "ТИП ОПЛАТЫ",
                                            ["Наличные", "Безнал", "Другое"]
                                        ),
                                        expand=1,
                                        padding=5
                                    ),
                                    ft.Container(
                                        content=create_custom_dropdown(
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