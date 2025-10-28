import flet as ft
from home_page import HomePage
from new_operation_page import NewOperationPage
from auto_deal_page import AutoDealPage
from financial_report_page import FinancialReportPage
from monthly_report_page import MonthlyReportPage


def main(page: ft.Page):
    # Настройки окна
    page.window.width = 1400
    page.window.height = 800
    page.window.title = "Finance App"
    page.scroll = ft.ScrollMode.ADAPTIVE  # ← включаем прокрутку для всей страницы

    # Устанавливаем темный фон для страницы (будет виден в "скруглениях")
    page.bgcolor = "#2D2D2D"  # темно-серый фон вокруг контента

    # Текущая активная страница
    current_page = "new_operation"  # ← основная страница
    previous_page = "new_operation"  # ← основная страница

    # Навигация между страницами
    def navigate_to(page_name):
        nonlocal current_page, previous_page
        # Запоминаем предыдущую страницу только когда переходим в настройки
        if page_name == "home" and current_page != "home":
            previous_page = current_page
        current_page = page_name

        # Обновляем стили кнопок
        for btn in nav_buttons:
            if btn.data == page_name:
                btn.style = active_button_style
            else:
                btn.style = inactive_button_style

        content_area.content = pages[page_name]
        content_area.update()
        nav_bar.update()

    # Функция для переключения меню/назад
    def toggle_menu(e):
        if current_page == "home":  # если уже в настройках
            navigate_to(previous_page)  # возвращаемся на ту страницу, откуда открыли
        else:
            navigate_to("home")  # переходим в настройки

    # Стили для кнопок навигации
    active_button_style = ft.ButtonStyle(
        color="white",
        bgcolor="#63437D",
        padding=ft.padding.symmetric(horizontal=30, vertical=25),
        shape=ft.RoundedRectangleBorder(radius=25),
        overlay_color="#63437D",
        elevation=8
    )

    inactive_button_style = ft.ButtonStyle(
        color="white",
        bgcolor="#835DA3",
        padding=ft.padding.symmetric(horizontal=30, vertical=25),
        shape=ft.RoundedRectangleBorder(radius=25),
        overlay_color="#63437D",
        elevation=8
    )

    # Создаем кнопки навигации
    nav_buttons = [
        ft.ElevatedButton(
            "НОВАЯ ОПЕРАЦИЯ",
            on_click=lambda _: navigate_to("new_operation"),
            style=inactive_button_style,
            data="new_operation"
        ),
        ft.ElevatedButton(
            "АВТО-СДЕЛКА",
            on_click=lambda _: navigate_to("auto_deal"),
            style=inactive_button_style,
            data="auto_deal"
        ),
        ft.ElevatedButton(
            "ФИНАНСОВЫЙ ОТЧЕТ",
            on_click=lambda _: navigate_to("financial_report"),
            style=inactive_button_style,
            data="financial_report"
        ),
        ft.ElevatedButton(
            "ОТЧЕТ ЗА МЕСЯЦ",
            on_click=lambda _: navigate_to("monthly_report"),
            style=inactive_button_style,
            data="monthly_report"
        ),
    ]

    # Создаем страницы
    pages = {
        "home": HomePage(navigate_to),
        "new_operation": NewOperationPage(navigate_to),
        "auto_deal": AutoDealPage(navigate_to),
        "financial_report": FinancialReportPage(navigate_to),
        "monthly_report": MonthlyReportPage(navigate_to)

    }

    # Область контента
    content_area = ft.Container(
        content=pages["home"],
        expand=True,
        padding=25
    )

    # Верхняя панель навигации
    nav_bar = ft.Container(
        content=ft.Row([
            # Основные кнопки в контейнере с отступом
            ft.Container(
                content=ft.Row(
                    nav_buttons,
                    spacing=50,
                    alignment=ft.MainAxisAlignment.CENTER
                ),
                expand=True,
                padding=ft.padding.only(left=40)  # ← отступ слева
            ),

            # SVG иконка меню справа
            ft.Container(
                content=ft.Image(
                    src="assets/menu_icon.svg",
                    width=40,
                    height=40,
                    color="#2C2C2C",
                ),
                on_click=toggle_menu,
                tooltip="Меню",
                padding=5,
                border_radius=25,

            )
        ]),
        padding=50,
        bgcolor="#8768B8",
        border_radius=ft.border_radius.only(
            bottom_left=25,
            bottom_right=25
        ),
        border=ft.border.only(bottom=ft.border.BorderSide(2, "#8768B8"))
    )


    # Главный контейнер со скругленными углами
    main_container = ft.Container(
        content=ft.Column([
            nav_bar,
            ft.Container(
                content=ft.Column([
                    content_area
                ], expand=True),
                expand=True
            )
        ]),
        margin=5,  # отступ от краев окна
        border_radius=25,  # скругление углов основного контейнера
        bgcolor="2C2C2C",  # белый фон контента
        expand=True,
        clip_behavior=ft.ClipBehavior.HARD_EDGE  # обрезает контент по скруглениям
    )

    page.add(main_container)

    # Устанавливаем домашнюю страницу как активную
    navigate_to("new_operation")
    page.update()


if __name__ == "__main__":
    ft.app(target=main)