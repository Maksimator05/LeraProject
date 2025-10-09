import flet as ft
from home_page import HomePage
from new_operation_page import NewOperationPage
from auto_deal_page import AutoDealPage
from financial_report_page import FinancialReportPage
from monthly_report_page import MonthlyReportPage


def main(page: ft.Page):
    # Текущая активная страница
    current_page = "home"

    # Навигация между страницами
    def navigate_to(page_name):
        nonlocal current_page
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

    # Стили для кнопок навигации
    active_button_style = ft.ButtonStyle(
        color="white",
        bgcolor="#835DA3",
        padding=ft.padding.symmetric(horizontal=20, vertical=12),
        shape=ft.RoundedRectangleBorder(radius=25)
    )

    inactive_button_style = ft.ButtonStyle(
        color="white",
        bgcolor="#835DA3",
        padding=ft.padding.symmetric(horizontal=20, vertical=12),
        shape=ft.RoundedRectangleBorder(radius=25)
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

    # Верхняя панель навигации
    nav_bar = ft.Container(
        content=ft.Column([
            # Кнопки навигации
            ft.Row(nav_buttons, spacing=15, alignment=ft.MainAxisAlignment.CENTER),
        ]),
        padding=20,
        bgcolor="#8768B8",
        border=ft.border.only(bottom=ft.border.BorderSide(2, "#8768B8"))
    )

    # Область контента
    content_area = ft.Container(
        content=pages["home"],
        expand=True,
        padding=25
    )

    # Добавляем кнопку для возврата на главную
    home_button = ft.FloatingActionButton(
        content=ft.Text("🏠"),
        bgcolor="#8768B8",
        on_click=lambda _: navigate_to("home"),
        tooltip="На главную",
        mini=True
    )

    page.add(
        nav_bar,
        ft.Container(
            content=ft.Column([
                content_area,
                ft.Container(
                    content=home_button,
                    alignment=ft.alignment.bottom_right,
                    padding=25
                )
            ], expand=True),
            expand=True
        )
    )

    # Устанавливаем домашнюю страницу как активную
    navigate_to("new_operation")
    page.update()


if __name__ == "__main__":
    ft.app(target=main)