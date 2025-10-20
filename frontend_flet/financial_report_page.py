import flet as ft


class FinancialReportPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        # Функция для создания стиля заголовка
        def create_header(text, width=120):
            return ft.Container(
                content=ft.Text(
                    text,
                    size=12,
                    weight=ft.FontWeight.BOLD,
                    color="white",
                    text_align=ft.TextAlign.CENTER
                ),
                padding=10,
                border=ft.border.all(1, "#835DA3"),
                bgcolor="#835DA3",
                height=40,
                width=width,
                border_radius=25,
                alignment=ft.alignment.center
            )

        # Функция для создания строки данных
        def create_data_row(*texts, widths, bg_color="#FFFFFF"):
            return ft.Container(
                content=ft.Row([
                    ft.Container(
                        content=ft.Text(
                            text,
                            size=14,
                            color="#FFFFFF",
                            text_align=ft.TextAlign.CENTER
                        ),
                        width=width,
                        padding=10,
                        alignment=ft.alignment.center,
                        border=ft.border.all(0.5, "#424242"),
                    ) for text, width in zip(texts, widths)
                ], spacing=50, alignment=ft.MainAxisAlignment.CENTER),
                bgcolor=bg_color,
            )

        operation_widths = [100, 120, 100, 150, 120, 120]
        auto_widths = [100, 80, 150, 120, 120, 100, 100, 150]

        self.controls = [
            ft.Container(
                content=ft.Column([
                    # Первая таблица - операции
                    ft.Container(
                        content=ft.Column([
                            # Заголовки
                            ft.Row([
                                create_header("ДАТА", 100),
                                create_header("ТИП ОПЕРАЦИИ", 160),
                                create_header("СУММА", 100),
                                create_header("ОПИСАНИЕ", 150),
                                create_header("КАТЕГОРИЯ", 120),
                                create_header("ТИП ОПЛАТЫ", 120),
                            ], spacing=50, alignment=ft.MainAxisAlignment.CENTER),

                            # Пустые данные для операций
                            ft.Container(
                                content=ft.ListView(
                                    controls=[],
                                    height=200,
                                    spacing=0,
                                ),
                                border=ft.border.all(1, "#FFFFFF"),
                                border_radius=25,
                            )
                        ], spacing=0),
                        padding=0,
                        border_radius=25,
                        bgcolor="#BFBDBD",
                        border=ft.border.all(1, "#BDBDBD"),
                    ),

                    ft.Container(height=5),

                    # Вторая таблица - авто-сделки
                    ft.Container(
                        content=ft.Column([
                            # Заголовки
                            ft.Row([
                                create_header("МАРКА", 100),
                                create_header("ГОД", 80),
                                create_header("VIN", 150),
                                create_header("ЦЕНА ПРОДАЖИ", 120),
                                create_header("ЗАКУП. СТОИМОСТЬ", 120),
                                create_header("РАСХОДЫ", 100),
                                create_header("ПРИБЫЛЬ", 100),
                                create_header("КОММЕНТАРИЙ", 150),
                            ], spacing=50, alignment=ft.MainAxisAlignment.CENTER),

                            # Пустые данные для авто-сделок
                            ft.Container(
                                content=ft.ListView(
                                    controls=[],
                                    height=200,
                                    spacing=0,
                                ),
                                border=ft.border.all(1, "#FFFFFF"),
                                border_radius=25,
                            )
                        ], spacing=0),
                        padding=0,
                        border_radius=25,
                        bgcolor="#BFBDBD",
                        border=ft.border.all(1, "#BDBDBD"),
                    ),

                    ft.Container(height=5),

                    # Третья таблица - итоги в виде карточек
                    ft.Container(
                        content=ft.Column([
                            ft.Container(height=5),

                            ft.Row([
                                # Карточка КАПИТАЛ
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "КАПИТАЛ",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),

                                # Карточка ОБЩИЙ ПРИХОД
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "ОБЩИЙ ПРИХОД",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),

                                # Карточка ОБЩИЙ РАСХОД
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "ОБЩИЙ РАСХОД",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),

                                # Карточка ДОП.ВЛОЖЕНИЯ
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "ДОП.ВЛОЖЕНИЯ",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),

                                # Карточка ПРИБЫЛЬ С АВТО
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "ПРИБЫЛЬ С АВТО",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),

                                # Карточка ОБЩАЯ ПРИБЫЛЬ
                                ft.Container(
                                    content=ft.Column([
                                        ft.Text(
                                            "ОБЩАЯ ПРИБЫЛЬ",
                                            size=12,
                                            weight=ft.FontWeight.BOLD,
                                            color="white",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                        ft.Container(height=8),
                                        ft.Text(
                                            "0 ₽",
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            color="#FFFFFF",
                                            text_align=ft.TextAlign.CENTER
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.CENTER,
                                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                    spacing=0
                                    ),
                                    width=150,
                                    height=80,
                                    padding=15,
                                    bgcolor="#835DA3",
                                    border_radius=45,
                                    alignment=ft.alignment.center,
                                ),
                            ], spacing=20, alignment=ft.MainAxisAlignment.CENTER, scroll=ft.ScrollMode.ADAPTIVE),
                        ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        padding=5,
                        bgcolor="#2C2C2C",
                        border_radius=25,
                    ),
                ]),
            ),
        ]
        self.spacing = 10
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER