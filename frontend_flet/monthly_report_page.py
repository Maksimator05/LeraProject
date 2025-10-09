import flet as ft


class MonthlyReportPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        # Функция для создания стиля заголовка
        def create_header(text):
            return ft.Container(
                content=ft.Text(text, size=12, weight=ft.FontWeight.BOLD),
                padding=10,
                border=ft.border.all(1, "#BDBDBD"),
                bgcolor="#E8F5E8",
                width=120,
                height=40,
                alignment=ft.alignment.center
            )

        # Функция для создания стиля обычной ячейки
        def create_cell(text):
            return ft.Container(
                content=ft.Text(text, size=12),
                padding=10,
                border=ft.border.all(1, "#BDBDBD"),
                bgcolor="#FFFFFF",
                width=120,
                height=40,
                alignment=ft.alignment.center
            )

        self.controls = [
            ft.Container(
                content=ft.Column([
                    # Заголовок страницы
                    ft.Container(
                        content=ft.Text("ОТЧЕТ ЗА МЕСЯЦ", size=18, weight=ft.FontWeight.BOLD),
                        padding=ft.padding.only(bottom=20)
                    ),

                    # Таблица отчета
                    ft.Column([
                        # Первая строка
                        ft.Row([
                            create_header("АРЕНДА"),
                            create_header("ПРИХОД"),
                            create_header("РАСХОД"),
                            create_header("ИТОГ В КАССЕ"),
                            create_header("ОПЕРАЦИЯ"),
                        ]),

                        # Вторая строка
                        ft.Row([
                            create_cell("ДАТА"),
                            create_cell("ТИП"),
                            create_cell("ОПИСАНИЕ"),
                            create_cell("КАТЕГОРИЯ"),
                            create_cell("ОПЛАТА"),
                        ]),

                        # Третья строка
                        ft.Row([
                            create_cell("АРЕНДА"),
                            create_cell("ДИЛЕРСТВО"),
                            create_cell("НАЛИЧНЫЕ"),
                            create_cell("БЕЗНАЛ"),
                            create_cell("КЦ"),
                        ]),

                        # Четвертая строка
                        ft.Row([
                            create_cell("ЗП ОКЛАД"),
                            create_cell("ЗП ПРОЦЕНТЫ"),
                            create_cell("РЕКЛАМА"),
                            create_cell("ВЕД. РЕКЛАМЫ"),
                            create_cell("КОМИССИЯ"),
                        ]),
                    ]),

                    # Кнопки управления отчетом
                    ft.Container(
                        content=ft.Row([
                            ft.ElevatedButton(
                                "ЭКСПОРТ ОТЧЕТА",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#2196F3",
                                    padding=15
                                )
                            ),
                            ft.ElevatedButton(
                                "ПЕЧАТЬ",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#4CAF50",
                                    padding=15
                                )
                            ),
                            ft.ElevatedButton(
                                "ОБНОВИТЬ",
                                style=ft.ButtonStyle(
                                    color="white",
                                    bgcolor="#FF9800",
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
        self.spacing = 0
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.START