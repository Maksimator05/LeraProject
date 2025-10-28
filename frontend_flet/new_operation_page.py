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

        # Словарь для хранения состояний dropdown'ов
        self.dropdown_states = {}

        # Функция для создания кастомного dropdown
        def create_custom_dropdown(label, options, dropdown_id):
            if dropdown_id not in self.dropdown_states:
                self.dropdown_states[dropdown_id] = {
                    "is_open": False,
                    "selected": options[0] if options else ""
                }

            state = self.dropdown_states[dropdown_id]

            # Контейнер для отображения выбранного значения
            selected_container = ft.Container(
                content=ft.Row([
                    ft.Text(
                        state["selected"],
                        color="white",
                        size=16,
                        expand=True
                    ),
                    ft.Text("▼", color="white", size=16),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                padding=20,
                border=ft.border.all(1, "#424242"),
                border_radius=25,
                bgcolor="#424242",
                width=300,
                height=60,  # ФИКСИРОВАННАЯ ВЫСОТА для расчета позиции
                on_click=lambda e: self.toggle_dropdown(dropdown_id),
            )

            # Контейнер с опциями - СДЕЛАЕМ ЕГО ПОВЕРХ ДРУГИХ ЭЛЕМЕНТОВ
            options_container = ft.Container(
                content=ft.Column([
                    ft.Container(
                        content=ft.Text(
                            option,
                            color="white",
                            size=16,
                        ),
                        padding=15,
                        width=270,
                        bgcolor="#424242",
                        border_radius=10,
                        on_click=lambda e, opt=option: self.select_option(dropdown_id, opt),
                    ) for option in options
                ], spacing=2, scroll=ft.ScrollMode.ADAPTIVE),  # ← ДОБАВЛЕН СКРОЛЛ
                bgcolor="#2C2C2C",
                padding=5,
                border_radius=15,
                width=300,  # ТА ЖЕ ШИРИНА КАК У ОСНОВНОГО ПОЛЯ
                visible=False,
                animate_opacity=200,
                # Позиционируем поверх других элементов
                top=85,  # Фиксированная позиция под основным полем (80 + 5)
                left=0,
                height=100,  # ← ФИКСИРОВАННАЯ ВЫСОТА ДЛЯ СКРОЛЛА
            )

            # Обертка с Stack для позиционирования поверх других элементов
            dropdown_stack = ft.Stack(
                [
                    ft.Column([
                        ft.Text(label, size=12, color="white", weight=ft.FontWeight.BOLD),
                        selected_container,
                    ], spacing=5),
                    options_container,
                ],
                height=200,  # ← УВЕЛИЧЕНА ВЫСОТА ДЛЯ ВМЕЩЕНИЯ ВЫПАДАЮЩЕГО СПИСКА
            )

            # Сохраняем ссылки на элементы
            state["selected_container"] = selected_container
            state["options_container"] = options_container
            state["dropdown_stack"] = dropdown_stack

            return dropdown_stack

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                create_custom_dropdown(
                                    "ОПЕРАЦИЯ",
                                    ["ПРИХОД", "РАСХОД"],
                                    "operation"
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
                                            ["Наличные", "Безнал", "Другое"],
                                            "payment"
                                        ),
                                        expand=1,
                                        padding=5
                                    ),
                                    ft.Container(
                                        content=create_custom_dropdown(
                                            "КАТЕГОРИЯ",
                                            ["КЦ", "АРЕНДА", "РЕКЛАМА", "ЗП ОКЛАДНИКИ", "ЗП ПРОЦЕНТ",
                                             "КОМИССИЯ", "ДИЛЕРСТВО", "ДРУГОЕ"],
                                            "category"
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

    def toggle_dropdown(self, dropdown_id):
        state = self.dropdown_states[dropdown_id]
        state["is_open"] = not state["is_open"]
        state["options_container"].visible = state["is_open"]

        # Закрываем другие открытые dropdown'ы
        for other_id, other_state in self.dropdown_states.items():
            if other_id != dropdown_id and other_state["is_open"]:
                other_state["is_open"] = False
                other_state["options_container"].visible = False
                other_state["options_container"].update()

        state["options_container"].update()

    def select_option(self, dropdown_id, option):
        state = self.dropdown_states[dropdown_id]
        state["selected"] = option
        # Обновляем текст в контейнере
        state["selected_container"].content.controls[0].value = option
        state["selected_container"].bgcolor = "#835DA3"
        state["selected_container"].border = ft.border.all(1, "#835DA3")
        state["is_open"] = False
        state["options_container"].visible = False

        state["selected_container"].update()
        state["options_container"].update()