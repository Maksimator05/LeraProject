import flet as ft


class FinancialReportPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        # Функция для создания стиля заголовка
        def create_header(text, width=120):
            return ft.Container(
                content=ft.Text(text, size=12, weight=ft.FontWeight.BOLD),
                padding=10,
                border=ft.border.all(1, "#BDBDBD"),
                bgcolor="#E3F2FD",
                height=40,
                width=width,
                alignment=ft.alignment.center
            )

        # Функция для создания стиля обычной ячейки
        def create_cell(text, width=120):
            return ft.Container(
                content=ft.Text(text, size=12),
                padding=10,
                border=ft.border.all(1, "#BDBDBD"),
                bgcolor="#FFFFFF",
                height=40,
                width=width,
                alignment=ft.alignment.center
            )

        self.controls = [
            ft.Container(
                content=ft.Column([
                    # Заголовок страницы
                    ft.Container(
                        content=ft.Text("ФИНАНСОВЫЙ ОТЧЕТ", size=18, weight=ft.FontWeight.BOLD),
                        padding=ft.padding.only(bottom=20)
                    ),

                    # Первая таблица - операции
                    ft.Container(
                        content=ft.Column([
                            ft.Text("ОПЕРАЦИИ", size=14, weight=ft.FontWeight.BOLD),
                            ft.Container(height=10),
                            ft.Row([
                                create_header("ДЕЛО", 80),
                                create_header("ТИП СПЕРВЫЙ", 100),
                                create_header("СЬЯМЯ", 80),
                                create_header("ОПИСАНИЕ", 120),
                                create_header("КОТОРЫЙ", 80),
                                create_header("ТИП ОПИСЫ", 100),
                            ]),
                        ]),
                        padding=10
                    ),

                    ft.Container(height=20),

                    # Вторая таблица - авто-сделки
                    ft.Container(
                        content=ft.Column([
                            ft.Text("АВТО-СДЕЛКИ", size=14, weight=ft.FontWeight.BOLD),
                            ft.Container(height=10),
                            ft.Row([
                                create_header("МАРКА", 80),
                                create_header("ГОД", 60),
                                create_header("VIN", 100),
                                create_header("ЦЕНА ПРОДАЖИ", 100),
                                create_header("ЗАКУП. СТОИМОСТЬ", 120),
                                create_header("РАСХОДЫ", 80),
                                create_header("ПРИБЫЛЬ", 80),
                                create_header("КОММЕНТАРИЙ", 120),
                            ]),
                        ]),
                        padding=10
                    ),

                    ft.Container(height=20),

                    # Третья таблица - итоги
                    ft.Container(
                        content=ft.Column([
                            ft.Text("ИТОГИ", size=14, weight=ft.FontWeight.BOLD),
                            ft.Container(height=10),
                            ft.Row([
                                create_header("КАПИТАЛ", 100),
                                create_header("ОБЩИЙ ПРИХОД", 120),
                                create_header("ОБЩИЙ РАСХОД", 120),
                                create_header("ДОП.ВЛОЖЕНИЯ", 120),
                                create_header("ПРИБЫЛЬ С АВТО", 120),
                                create_header("ОБЩАЯ ПРИБЫЛЬ", 120),
                            ]),
                        ]),
                        padding=10
                    ),

                    # Кнопки управления
                    ft.Container(
                        content=ft.Row([
                            ft.ElevatedButton(
                                "СФОРМИРОВАТЬ ОТЧЕТ",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#2196F3",
                                    padding=15
                                )
                            ),
                            ft.ElevatedButton(
                                "ЭКСПОРТ В EXCEL",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#4CAF50",
                                    padding=15
                                )
                            ),
                            ft.ElevatedButton(
                                "ОЧИСТИТЬ",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#F44336",
                                    padding=15
                                )
                            ),
                        ], spacing=20),
                        padding=ft.padding.only(top=30)
                    )
                ]),
                padding=30,
                border=ft.border.all(1, "#BDBDBD"),
                border_radius=10,
                bgcolor="#FAFAFA"
            )
        ]
        self.spacing = 10
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.START