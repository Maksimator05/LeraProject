import flet as ft


class HomePage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.TextButton("ввод капитала", on_click=lambda _: print("Ввод капитала")),
                        ft.TextButton("Настройки", on_click=lambda _: print("Настройки")),
                    ]),
                    ft.Divider(),
                    ft.Row([
                        ft.TextButton("Капитал", on_click=lambda _: print("Капитал")),
                        ft.TextButton("Настройки", on_click=lambda _: print("Настройки")),
                        ft.TextButton("Экспорт", on_click=lambda _: print("Экспорт")),
                        ft.TextButton("Импорт", on_click=lambda _: print("Импорт")),
                    ], spacing=20)
                ]),
                padding=20
            )
        ]