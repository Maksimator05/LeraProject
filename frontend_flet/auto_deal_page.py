import flet as ft


class AutoDealPage(ft.Column):
    def __init__(self, navigate_to):
        super().__init__()
        self.navigate_to = navigate_to

        self.controls = [
            ft.Container(
                content=ft.Column([
                    ft.Text("МАРКА", size=14),
                    ft.TextField(border_color=ft.colors.BLACK),

                    ft.Row([
                        ft.Column([
                            ft.Text("Год", size=14),
                            ft.TextField(border_color=ft.colors.BLACK, width=150),
                        ]),
                        ft.Column([
                            ft.Text("VIN", size=14),
                            ft.TextField(border_color=ft.colors.BLACK, width=200),
                        ]),
                    ], spacing=20),

                    ft.Text("ЦЕНА С ОПЦИЯМИ", size=14),
                    ft.TextField(border_color=ft.colors.BLACK),

                    ft.Text("КОММЕНТАРИЙ", size=14),
                    ft.TextField(border_color=ft.colors.BLACK, multiline=True),

                    ft.Text("ЗАКУПОЧНАЯ СТОИМОСТЬ", size=14),
                    ft.TextField(border_color=ft.colors.BLACK),

                    ft.Text("РАСХОДЫ", size=14),
                    ft.TextField(border_color=ft.colors.BLACK),

                    ft.Container(
                        content=ft.TextButton(
                            "ДОБАВИТЬ СДЕЛКУ",
                            style=ft.ButtonStyle(
                                color=ft.colors.WHITE,
                                bgcolor=ft.colors.BLUE
                            )
                        ),
                        padding=10
                    )
                ], spacing=10),
                padding=20
            )
        ]