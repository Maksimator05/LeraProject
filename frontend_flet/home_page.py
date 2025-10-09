import flet as ft


class HomePage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        # Стиль для основных кнопок
        main_button_style = ft.ButtonStyle(
            color="white",
            bgcolor="#2196F3",
            padding=20,
            elevation=8,
            shape=ft.RoundedRectangleBorder(radius=10)
        )

        # Стиль для второстепенных кнопок
        secondary_button_style = ft.ButtonStyle(
            color="black",
            bgcolor="#E0E0E0",
            padding=15,
            elevation=4,
            shape=ft.RoundedRectangleBorder(radius=8)
        )

        self.controls = [
            # Заголовок главной страницы
            ft.Container(
                content=ft.Column([
                    ft.Text("ГЛАВНАЯ ПАНЕЛЬ", size=24, weight=ft.FontWeight.BOLD, color="#1976D2"),
                    ft.Text("Управление финансами и автомобильными сделками", size=16, color="#757575"),
                ]),
                padding=ft.padding.only(bottom=30)
            ),

            # Основные действия
            ft.Container(
                content=ft.Column([
                    ft.Text("ОСНОВНЫЕ ДЕЙСТВИЯ", size=18, weight=ft.FontWeight.BOLD, color="#424242"),
                    ft.Container(height=20),
                    ft.Row([
                        ft.ElevatedButton(
                            "ВВОД КАПИТАЛА",
                            on_click=lambda _: print("Ввод капитала"),
                            style=main_button_style,
                            width=200
                        ),
                        ft.ElevatedButton(
                            "НАСТРОЙКИ",
                            on_click=lambda _: print("Настройки"),
                            style=main_button_style,
                            width=200
                        ),
                    ], spacing=30, alignment=ft.MainAxisAlignment.CENTER),
                ]),
                padding=20,
                border=ft.border.all(1, "#BDBDBD"),
                border_radius=10,
                bgcolor="#FAFAFA"
            ),

            ft.Container(height=30),

            # Быстрый доступ
            ft.Container(
                content=ft.Column([
                    ft.Text("БЫСТРЫЙ ДОСТУП", size=18, weight=ft.FontWeight.BOLD, color="#424242"),
                    ft.Container(height=20),
                    ft.Row([
                        ft.ElevatedButton(
                            "КАПИТАЛ",
                            on_click=lambda _: print("Капитал"),
                            style=secondary_button_style
                        ),
                        ft.ElevatedButton(
                            "НАСТРОЙКИ",
                            on_click=lambda _: print("Настройки"),
                            style=secondary_button_style
                        ),
                        ft.ElevatedButton(
                            "ЭКСПОРТ",
                            on_click=lambda _: print("Экспорт"),
                            style=secondary_button_style
                        ),
                        ft.ElevatedButton(
                            "ИМПОРТ",
                            on_click=lambda _: print("Импорт"),
                            style=secondary_button_style
                        ),
                    ], spacing=20, alignment=ft.MainAxisAlignment.CENTER),
                ]),
                padding=20,
                border=ft.border.all(1, "#BDBDBD"),
                border_radius=10,
                bgcolor="#FAFAFA"
            ),

            ft.Container(height=30),

            # Статистика и быстрые данные
            ft.Container(
                content=ft.Column([
                    ft.Text("КРАТКАЯ СТАТИСТИКА", size=18, weight=ft.FontWeight.BOLD, color="#424242"),
                    ft.Container(height=20),
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                ft.Text("ТЕКУЩИЙ КАПИТАЛ", size=12, weight=ft.FontWeight.BOLD, color="#757575"),
                                ft.Text("1 250 000 ₽", size=20, weight=ft.FontWeight.BOLD, color="#4CAF50"),
                            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            padding=15,
                            border=ft.border.all(1, "#C8E6C9"),
                            border_radius=8,
                            bgcolor="#E8F5E8",
                            width=150
                        ),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("АКТИВНЫЕ СДЕЛКИ", size=12, weight=ft.FontWeight.BOLD, color="#757575"),
                                ft.Text("3", size=20, weight=ft.FontWeight.BOLD, color="#2196F3"),
                            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            padding=15,
                            border=ft.border.all(1, "#BBDEFB"),
                            border_radius=8,
                            bgcolor="#E3F2FD",
                            width=150
                        ),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("ОБЩАЯ ПРИБЫЛЬ", size=12, weight=ft.FontWeight.BOLD, color="#757575"),
                                ft.Text("350 000 ₽", size=20, weight=ft.FontWeight.BOLD, color="#FF9800"),
                            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            padding=15,
                            border=ft.border.all(1, "#FFE0B2"),
                            border_radius=8,
                            bgcolor="#FFF3E0",
                            width=150
                        ),
                    ], spacing=20, alignment=ft.MainAxisAlignment.CENTER),
                ]),
                padding=20,
                border=ft.border.all(1, "#BDBDBD"),
                border_radius=10,
                bgcolor="#FAFAFA"
            )
        ]
        self.spacing = 0
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER