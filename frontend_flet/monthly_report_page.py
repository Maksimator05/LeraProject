import flet as ft


class MonthlyReportPage(ft.Column):
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
                ], spacing=70, alignment=ft.MainAxisAlignment.CENTER),
                bgcolor=bg_color,
            )

        # Ширины столбцов для разных таблиц
        summary_widths = [150, 150, 150, 150, 150]
        details_widths = [120, 120, 150, 120, 120]

        self.controls = [
            ft.Container(
                content=ft.Column([
                    # Первая таблица - сводка по категориям
                    ft.Container(
                        content=ft.Column([
                            # Заголовки
                            ft.Row([
                                create_header("АРЕНДА", 150),
                                create_header("ПРИХОД", 150),
                                create_header("РАСХОД", 150),
                                create_header("ИТОГ В КАССЕ", 150),
                                create_header("ОПЕРАЦИЯ", 150),
                            ], spacing=70, alignment=ft.MainAxisAlignment.CENTER),

                            # Пустые данные
                            ft.Container(
                                content=ft.ListView(
                                    controls=[],
                                    height=150,
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

                    # Вторая таблица - детали операций
                    ft.Container(
                        content=ft.Column([
                            ft.Container(height=5),

                            # Заголовки
                            ft.Row([
                                create_header("ДАТА", 120),
                                create_header("ТИП", 120),
                                create_header("ОПИСАНИЕ", 150),
                                create_header("КАТЕГОРИЯ", 120),
                                create_header("ОПЛАТА", 120),
                            ], spacing=70, alignment=ft.MainAxisAlignment.CENTER),

                            # Пустые данные
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

                    # Карточки
                    ft.Container(
                        content=ft.Column([
                            ft.Container(height=0),
                            # Все карточки в два ряда
                            ft.Column([
                                # Первый ряд карточек
                                ft.Row([
                                    # Карточка АРЕНДА
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "АРЕНДА",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка ДИЛЕРСТВО
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "ДИЛЕРСТВО",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка НАЛИЧНЫЕ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "НАЛИЧНЫЕ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка БЕЗНАЛ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "БЕЗНАЛ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка КЦ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "КЦ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),
                                ], spacing=100, alignment=ft.MainAxisAlignment.CENTER, scroll=ft.ScrollMode.ADAPTIVE),

                                ft.Container(height=0),

                                # Второй ряд карточек
                                ft.Row([
                                    # Карточка ЗП ОКЛАД
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "ЗП ОКЛАД",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка ЗП ПРОЦЕНТЫ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "ЗП ПРОЦЕНТЫ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка РЕКЛАМА
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "РЕКЛАМА",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка ВЕД. РЕКЛАМЫ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "ВЕД. РЕКЛАМЫ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),

                                    # Карточка КОМИССИЯ
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Text(
                                                "КОМИССИЯ",
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
                                        width=180,
                                        height=60,
                                        padding=5,
                                        bgcolor="#835DA3",
                                        border_radius=45,
                                        alignment=ft.alignment.center,
                                    ),
                                ], spacing=100, alignment=ft.MainAxisAlignment.CENTER, scroll=ft.ScrollMode.ADAPTIVE),
                            ], alignment=ft.MainAxisAlignment.CENTER,
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        padding=5,
                        bgcolor="#424242",
                        border_radius=25,
                    ),
                ]),
            )
        ]
        self.spacing = 0
        self.alignment = ft.MainAxisAlignment.START
        self.horizontal_alignment = ft.CrossAxisAlignment.CENTER