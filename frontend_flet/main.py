import flet as ft
from datetime import datetime
import pandas as pd
import sys
import os
from typing import Dict, List, Any
import asyncio

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.database import DatabaseManager


class FletMoneyTrackerApp:
    def __init__(self):
        self.db = DatabaseManager()
        self.page = None
        self.data = {
            'transactions': [],
            'car_deals': [],
            'initial_capital': 0
        }

        # Переменные состояния
        self.selected_year = datetime.now().year
        self.selected_month = datetime.now().month
        self.selected_day = None

        # Таблицы (инициализируем позже)
        self.transactions_table = None
        self.car_deals_table = None
        self.daily_summary_table = None
        self.daily_details_table = None

        # Статистика
        self.summary_labels = {}
        self.category_labels = {}

        self.load_data()

    def load_data(self):
        """Загрузка данных из базы"""
        self.data['transactions'] = self.db.get_all_transactions()
        self.data['car_deals'] = self.db.get_all_car_deals()
        self.data['initial_capital'] = self.db.get_initial_capital()

    def refresh_data(self):
        """Обновление данных"""
        self.load_data()
        self.update_report()
        self.update_monthly_report()
        self.update_ui()

    def update_ui(self):
        """Обновление UI"""
        if self.page:
            self.page.update()

    def show_snack_bar(self, message: str, color: str = ft.Colors.GREEN):
        """Показать уведомление"""
        if self.page:
            self.page.snack_bar = ft.SnackBar(
                content=ft.Text(message, color=ft.Colors.WHITE),
                bgcolor=color,
                duration=3000
            )
            self.page.snack_bar.open = True
            self.page.update()

    # Бизнес-логика (остается без изменений)
    def add_transaction(self, operation: str, amount: float, description: str,
                        category: str, payment_type: str = "Наличные",
                        exclude_from_total: bool = False) -> bool:
        """Добавление транзакции"""
        try:
            if not description or amount <= 0:
                return False

            transaction = {
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                "type": operation,
                "amount": amount if operation == "Приход" else -amount,
                "description": description.strip(),
                "category": category,
                "payment_type": payment_type,
                "exclude_from_total": exclude_from_total
            }

            self.db.add_transaction(transaction)
            self.refresh_data()
            return True
        except Exception as e:
            print(f"Ошибка добавления транзакции: {e}")
            return False

    def add_car_deal(self, brand: str, year: str, vin: str, price: float = 0,
                     cost: float = 0, expenses: float = 0, comment: str = "") -> bool:
        """Добавление авто-сделки"""
        try:
            if not brand:
                return False

            profit = price - cost - expenses
            car_deal = {
                "brand": brand.strip(),
                "year": year.strip(),
                "vin": vin.strip(),
                "price": price,
                "cost": cost,
                "expenses": expenses,
                "header": profit,
                "comment": comment.strip()
            }

            self.db.add_car_deal(car_deal)
            self.refresh_data()
            return True
        except Exception as e:
            print(f"Ошибка добавления авто-сделки: {e}")
            return False

    def get_financial_summary(self) -> Dict[str, float]:
        """Получение финансовой сводки"""
        transactions = self.data['transactions']
        car_deals = self.data['car_deals']
        initial_capital = self.data['initial_capital']

        excluded_categories = ["ЗП окладники", "ЗП проценты", "Комиссия брок"]

        filtered_transactions = [
            t for t in transactions
            if t["category"] not in excluded_categories and not t.get("exclude_from_total", False)
        ]

        total_income = sum(t["amount"] for t in filtered_transactions if t["type"] == "Приход")
        total_expense = abs(sum(t["amount"] for t in filtered_transactions if t["type"] == "Расход"))
        additional_investment = max(0, total_expense - initial_capital)
        car_profit = sum(deal.get("header", 0) for deal in car_deals)
        total_profit = car_profit + total_income - additional_investment

        return {
            "initial_capital": initial_capital,
            "total_income": total_income,
            "total_expense": total_expense,
            "additional_investment": additional_investment,
            "car_profit": car_profit,
            "total_profit": total_profit
        }

    def get_monthly_report_data(self, year: int, month: int) -> Dict[str, Any]:
        """Данные месячного отчета"""
        transactions = self.data['transactions']

        daily_data = {}
        categories = ["Аренда", "Дилерство", "Наличные", "Безнал", "КЦ",
                      "ЗП окладники", "ЗП проценты", "Реклама", "Вед.рекламы", "Комиссия брок"]
        category_stats = {category: 0 for category in categories}
        total_income = 0
        total_expense = 0

        for transaction in transactions:
            try:
                date_str = transaction["date"].split()[0]
                day, trans_month, trans_year = map(int, date_str.split('.'))

                if trans_year == year and trans_month == month:
                    if date_str not in daily_data:
                        daily_data[date_str] = {
                            'all_income': 0,
                            'all_expense': 0,
                            'transactions': []
                        }

                    amount = abs(transaction["amount"])
                    daily_data[date_str]['transactions'].append(transaction)

                    if transaction["type"] == "Приход":
                        daily_data[date_str]['all_income'] += amount
                        total_income += amount
                    else:
                        daily_data[date_str]['all_expense'] += amount
                        total_expense += amount

                    # Статистика по категориям
                    category = transaction["category"]
                    if category in category_stats:
                        if transaction["type"] == "Расход":
                            category_stats[category] += amount
                        else:
                            category_stats[category] -= amount

            except (ValueError, IndexError):
                continue

        # Формируем ежедневную сводку
        daily_summary = []
        for date_str in sorted(daily_data.keys()):
            data = daily_data[date_str]
            balance = data['all_income'] - data['all_expense']
            transactions_count = len(data['transactions'])

            daily_summary.append({
                'Дата': date_str,
                'Приход': data['all_income'],
                'Расход': data['all_expense'],
                'Баланс': balance,
                'Количество_операций': transactions_count
            })

        return {
            'daily_summary': daily_summary,
            'category_stats': category_stats,
            'total_income': total_income,
            'total_expense': total_expense
        }

    def update_report(self):
        """Обновление финансового отчета"""
        if not self.transactions_table or not self.car_deals_table:
            return

        # Обновление таблицы транзакций
        transactions_rows = []
        for i, tr in enumerate(self.data['transactions']):
            transactions_rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(tr["date"])),
                    ft.DataCell(ft.Text(tr["type"])),
                    ft.DataCell(ft.Text(f"{abs(tr['amount']):,.2f}")),
                    ft.DataCell(ft.Text(tr["description"])),
                    ft.DataCell(ft.Text(tr["category"])),
                    ft.DataCell(ft.Text(tr.get("payment_type", "Наличные"))),
                ]
            ))

        self.transactions_table.rows = transactions_rows

        # Обновление таблицы авто-сделок
        car_rows = []
        for i, deal in enumerate(self.data['car_deals']):
            car_rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(deal.get("brand", ""))),
                    ft.DataCell(ft.Text(deal.get("year", ""))),
                    ft.DataCell(ft.Text(deal.get("vin", ""))),
                    ft.DataCell(ft.Text(f"{deal.get('price', 0):,.2f}")),
                    ft.DataCell(ft.Text(f"{deal.get('cost', 0):,.2f}")),
                    ft.DataCell(ft.Text(f"{deal.get('expenses', 0):,.2f}")),
                    ft.DataCell(ft.Text(f"{deal.get('header', 0):,.2f}")),
                    ft.DataCell(ft.Text(deal.get("comment", ""))),
                ]
            ))

        self.car_deals_table.rows = car_rows

        # Обновление сводки
        summary = self.get_financial_summary()
        if hasattr(self, 'summary_labels'):
            for key, value in summary.items():
                if key in self.summary_labels:
                    self.summary_labels[key].value = f"{value:,.2f} ₽"

    def update_monthly_report(self):
        """Обновление месячного отчета"""
        if not self.daily_summary_table:
            return

        monthly_data = self.get_monthly_report_data(self.selected_year, self.selected_month)

        # Ежедневная сводка
        daily_rows = []
        for day_data in monthly_data['daily_summary']:
            daily_rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(day_data['Дата'])),
                    ft.DataCell(ft.Text(f"{day_data['Приход']:,.2f}")),
                    ft.DataCell(ft.Text(f"{day_data['Расход']:,.2f}")),
                    ft.DataCell(ft.Text(f"{day_data['Баланс']:,.2f}")),
                    ft.DataCell(ft.Text(str(day_data['Количество_операций']))),
                ]
            ))

        self.daily_summary_table.rows = daily_rows

        # Статистика по категориям
        if hasattr(self, 'category_labels'):
            for category, amount in monthly_data['category_stats'].items():
                if category in self.category_labels:
                    self.category_labels[category].value = f"{abs(amount):,.2f} ₽"

    async def main(self, page: ft.Page):
        self.page = page
        page.title = "💰 Авто-Трекер Финансов v3.0 (Flet)"
        page.theme_mode = ft.ThemeMode.DARK
        page.padding = 0

        # Установка иконки приложения, если файл существует
        if os.path.exists("app_icon.png"):
            page.window.icon = "app_icon.png"

        # Инициализация UI компонентов
        await self.init_ui_components()

        # Навигация
        nav_bar = ft.NavigationBar(
            destinations=[
                ft.NavigationBarDestination(icon=ft.Icons.ADD, label="Добавить"),
                ft.NavigationBarDestination(icon=ft.Icons.DIRECTIONS_CAR, label="Авто"),
                ft.NavigationBarDestination(icon=ft.Icons.ANALYTICS, label="Отчет"),
                ft.NavigationBarDestination(icon=ft.Icons.CALENDAR_MONTH, label="Месяц"),
                ft.NavigationBarDestination(icon=ft.Icons.SETTINGS, label="Настройки"),
            ],
            on_change=self.on_navigation_change
        )

        # Контентная область
        self.content_area = ft.Container(
            content=await self.build_add_page(),
            expand=True,
            padding=20
        )

        page.add(
            ft.Column([
                # Заголовок
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.ACCOUNT_BALANCE_WALLET, size=30, color=ft.Colors.BLUE),
                        ft.Text("Авто-Трекер Финансов", size=24, weight="bold", color=ft.Colors.WHITE),
                    ]),
                    padding=20,
                    bgcolor=ft.Colors.BLUE_GREY_900
                ),

                # Основной контент
                self.content_area,

                # Навигация
                nav_bar
            ], expand=True)
        )

        # Первоначальная загрузка данных
        self.refresh_data()

    async def init_ui_components(self):
        """Инициализация UI компонентов"""
        # Поля для добавления транзакций
        self.transaction_type = ft.Dropdown(
            label="Тип операции",
            options=[
                ft.dropdown.Option("Приход"),
                ft.dropdown.Option("Расход")
            ],
            value="Приход",
            width=300
        )

        self.amount_input = ft.TextField(
            label="Сумма",
            value="0.00",
            width=300
        )

        self.description_input = ft.TextField(
            label="Описание",
            width=300
        )

        self.payment_type = ft.Dropdown(
            label="Тип оплаты",
            options=[
                ft.dropdown.Option("Наличные"),
                ft.dropdown.Option("Безнал"),
                ft.dropdown.Option("Другое")
            ],
            value="Наличные",
            width=300
        )

        self.category = ft.Dropdown(
            label="Категория",
            options=[
                ft.dropdown.Option("КЦ"),
                ft.dropdown.Option("Реклама"),
                ft.dropdown.Option("Вед.рекламы"),
                ft.dropdown.Option("Комиссия брок"),
                ft.dropdown.Option("Дилерство"),
                ft.dropdown.Option("Аренда"),
                ft.dropdown.Option("ЗП окладники"),
                ft.dropdown.Option("ЗП проценты"),
                ft.dropdown.Option("Другое")
            ],
            value="Другое",
            width=300
        )

        self.exclude_checkbox = ft.Checkbox(
            label="Исключить из общего расхода",
            value=False
        )

        # Поля для авто-сделок
        self.car_brand = ft.TextField(label="Марка", width=300)
        self.car_year = ft.TextField(label="Год", width=300)
        self.car_vin = ft.TextField(label="VIN", width=300)
        self.car_price = ft.TextField(label="Цена продажи", value="0", width=300)
        self.car_cost = ft.TextField(label="Закупочная стоимость", value="0", width=300)
        self.car_expenses = ft.TextField(label="Расходы", value="0", width=300)
        self.car_comment = ft.TextField(label="Комментарий", width=300)

        # Настройки
        self.capital_input = ft.TextField(
            label="Стартовый капитал",
            value=str(self.data['initial_capital']),
            width=300
        )

        # Инициализация таблиц
        self.transactions_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Дата")),
                ft.DataColumn(ft.Text("Тип")),
                ft.DataColumn(ft.Text("Сумма")),
                ft.DataColumn(ft.Text("Описание")),
                ft.DataColumn(ft.Text("Категория")),
                ft.DataColumn(ft.Text("Тип оплаты")),
            ],
            rows=[]
        )

        self.car_deals_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Марка")),
                ft.DataColumn(ft.Text("Год")),
                ft.DataColumn(ft.Text("VIN")),
                ft.DataColumn(ft.Text("Цена")),
                ft.DataColumn(ft.Text("Стоимость")),
                ft.DataColumn(ft.Text("Расходы")),
                ft.DataColumn(ft.Text("Прибыль")),
                ft.DataColumn(ft.Text("Комментарий")),
            ],
            rows=[]
        )

        self.daily_summary_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Дата")),
                ft.DataColumn(ft.Text("Приход")),
                ft.DataColumn(ft.Text("Расход")),
                ft.DataColumn(ft.Text("Баланс")),
                ft.DataColumn(ft.Text("Операций")),
            ],
            rows=[]
        )

    async def build_add_page(self):
        """Страница добавления операций"""
        return ft.Column([
            ft.Text("➕ Добавить операцию", size=24, weight="bold"),
            ft.Container(height=20),

            self.transaction_type,
            self.amount_input,
            self.description_input,
            self.payment_type,
            self.category,
            self.exclude_checkbox,

            ft.Container(height=20),
            ft.ElevatedButton(
                "Добавить операцию",
                on_click=self.on_add_transaction,
                style=ft.ButtonStyle(
                    color=ft.Colors.WHITE,
                    bgcolor=ft.Colors.GREEN,
                    padding=20
                ),
                width=300
            )
        ])

    async def build_car_page(self):
        """Страница авто-сделок"""
        return ft.Column([
            ft.Text("🚗 Авто-сделки", size=24, weight="bold"),
            ft.Container(height=20),

            self.car_brand,
            self.car_year,
            self.car_vin,
            self.car_price,
            self.car_cost,
            self.car_expenses,
            self.car_comment,

            ft.Container(height=20),
            ft.ElevatedButton(
                "Добавить авто-сделку",
                on_click=self.on_add_car_deal,
                style=ft.ButtonStyle(
                    color=ft.Colors.WHITE,
                    bgcolor=ft.Colors.PURPLE,
                    padding=20
                ),
                width=300
            ),

            ft.Container(height=30),
            ft.Text("Список авто-сделок", size=18, weight="bold"),
            ft.Container(
                content=ft.Column([self.car_deals_table], scroll=ft.ScrollMode.ADAPTIVE),
                height=400,
                border=ft.border.all(1, ft.Colors.GREY_700)
            )
        ])

    async def build_report_page(self):
        """Страница финансового отчета"""
        # Инициализация сводки
        summary_data = self.get_financial_summary()
        self.summary_labels = {
            "initial_capital": ft.Text(f"{summary_data['initial_capital']:,.2f} ₽", size=18, weight="bold"),
            "total_income": ft.Text(f"{summary_data['total_income']:,.2f} ₽", size=18, weight="bold"),
            "total_expense": ft.Text(f"{summary_data['total_expense']:,.2f} ₽", size=18, weight="bold"),
            "additional_investment": ft.Text(f"{summary_data['additional_investment']:,.2f} ₽", size=18, weight="bold"),
            "car_profit": ft.Text(f"{summary_data['car_profit']:,.2f} ₽", size=20, weight="bold",
                                  color=ft.Colors.GREEN),
            "total_profit": ft.Text(f"{summary_data['total_profit']:,.2f} ₽", size=20, weight="bold",
                                    color=ft.Colors.BLUE),
        }

        return ft.Column([
            ft.Text("📊 Финансовый отчет", size=24, weight="bold"),
            ft.Container(height=20),

            # Финансовая сводка
            ft.Card(
                content=ft.Container(
                    content=ft.Column([
                        ft.Text("Финансовая сводка", size=20, weight="bold"),
                        ft.Container(height=10),
                        ft.Row([
                            ft.Column([
                                ft.Text("Стартовый капитал:", size=16),
                                ft.Text("Общий приход:", size=16),
                                ft.Text("Общий расход:", size=16),
                            ]),
                            ft.Column([
                                self.summary_labels["initial_capital"],
                                self.summary_labels["total_income"],
                                self.summary_labels["total_expense"],
                            ])
                        ]),
                        ft.Container(height=10),
                        ft.Row([
                            ft.Column([
                                ft.Text("Доп. вложения:", size=16),
                                ft.Text("Прибыль с авто:", size=16),
                                ft.Text("Общая прибыль:", size=16),
                            ]),
                            ft.Column([
                                self.summary_labels["additional_investment"],
                                self.summary_labels["car_profit"],
                                self.summary_labels["total_profit"],
                            ])
                        ])
                    ]),
                    padding=20
                )
            ),

            ft.Container(height=30),
            ft.Text("История операций", size=18, weight="bold"),
            ft.Container(
                content=ft.Column([self.transactions_table], scroll=ft.ScrollMode.ADAPTIVE),
                height=400,
                border=ft.border.all(1, ft.Colors.GREY_700)
            )
        ])

    async def build_monthly_page(self):
        """Страница месячного отчета"""
        # Выбор года и месяца
        current_year = datetime.now().year
        years = [str(year) for year in range(current_year - 2, current_year + 1)]
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

        year_dropdown = ft.Dropdown(
            label="Год",
            options=[ft.dropdown.Option(year) for year in years],
            value=str(current_year),
            width=150,
            on_change=self.on_monthly_filter_change
        )

        month_dropdown = ft.Dropdown(
            label="Месяц",
            options=[ft.dropdown.Option(month) for month in months],
            value=months[datetime.now().month - 1],
            width=150,
            on_change=self.on_monthly_filter_change
        )

        # Статистика по категориям
        categories = ["Аренда", "Дилерство", "Наличные", "Безнал", "КЦ",
                      "ЗП окладники", "ЗП проценты", "Реклама", "Вед.рекламы", "Комиссия брок"]

        self.category_labels = {}
        category_grid = ft.GridView(
            runs_count=5,
            max_extent=150,
            child_aspect_ratio=1.0,
            spacing=10,
            run_spacing=10,
        )

        for category in categories:
            container = ft.Container(
                content=ft.Column([
                    ft.Text(category, size=12, text_align="center"),
                    ft.Text("0.00 ₽", size=14, weight="bold", text_align="center")
                ], alignment="center", horizontal_alignment="center"),
                width=140,
                height=140,
                padding=10,
                border_radius=10,
                bgcolor=ft.Colors.BLUE_GREY_900,
            )
            self.category_labels[category] = container.content.controls[1]
            category_grid.controls.append(container)

        return ft.Column([
            ft.Text("📅 Расходы за месяц", size=24, weight="bold"),
            ft.Container(height=20),

            # Фильтры
            ft.Row([year_dropdown, month_dropdown]),
            ft.Container(height=20),

            # Ежедневная сводка
            ft.Text("Ежедневная сводка", size=18, weight="bold"),
            ft.Container(
                content=ft.Column([self.daily_summary_table], scroll=ft.ScrollMode.ADAPTIVE),
                height=300,
                border=ft.border.all(1, ft.Colors.GREY_700)
            ),

            ft.Container(height=20),
            ft.Text("Статистика по категориям", size=18, weight="bold"),
            category_grid
        ])

    async def build_settings_page(self):
        """Страница настроек"""
        return ft.Column([
            ft.Text("⚙️ Настройки", size=24, weight="bold"),
            ft.Container(height=20),

            self.capital_input,
            ft.Container(height=10),

            ft.Row([
                ft.ElevatedButton(
                    "💾 Сохранить капитал",
                    on_click=self.on_save_capital
                ),
                ft.ElevatedButton(
                    "📥 Импорт из Excel",
                    on_click=self.on_import_excel
                ),
                ft.ElevatedButton(
                    "📤 Экспорт в Excel",
                    on_click=self.on_export_excel
                ),
            ])
        ])

    async def on_navigation_change(self, e):
        """Смена страницы"""
        index = e.control.selected_index
        pages = [
            await self.build_add_page(),
            await self.build_car_page(),
            await self.build_report_page(),
            await self.build_monthly_page(),
            await self.build_settings_page()
        ]
        self.content_area.content = pages[index]
        await self.update_ui()

    async def on_add_transaction(self, e):
        """Добавление транзакции"""
        try:
            operation = self.transaction_type.value
            amount = float(self.amount_input.value)
            description = self.description_input.value.strip()
            payment_type = self.payment_type.value
            category = self.category.value
            exclude_from_total = self.exclude_checkbox.value

            if not description:
                self.show_snack_bar("❌ Введите описание операции!", ft.Colors.RED)
                return
            if amount <= 0:
                self.show_snack_bar("❌ Сумма должна быть положительной!", ft.Colors.RED)
                return

            success = self.add_transaction(operation, amount, description, category, payment_type, exclude_from_total)

            if success:
                # Сброс формы
                self.amount_input.value = "0.00"
                self.description_input.value = ""
                self.exclude_checkbox.value = False

                message = "✅ Операция добавлена"
                if exclude_from_total:
                    message += " (исключена из расхода)"
                self.show_snack_bar(message)
                await self.update_ui()
            else:
                self.show_snack_bar("❌ Не удалось добавить операцию", ft.Colors.RED)

        except ValueError:
            self.show_snack_bar("❌ Введите корректную сумму!", ft.Colors.RED)

    async def on_add_car_deal(self, e):
        """Добавление авто-сделки"""
        try:
            brand = self.car_brand.value.strip()
            year = self.car_year.value.strip()
            vin = self.car_vin.value.strip()
            price = float(self.car_price.value or 0)
            cost = float(self.car_cost.value or 0)
            expenses = float(self.car_expenses.value or 0)
            comment = self.car_comment.value.strip()

            if not brand:
                self.show_snack_bar("❌ Введите марку авто!", ft.Colors.RED)
                return

            success = self.add_car_deal(brand, year, vin, price, cost, expenses, comment)

            if success:
                # Сброс формы
                self.car_brand.value = ""
                self.car_year.value = ""
                self.car_vin.value = ""
                self.car_price.value = "0"
                self.car_cost.value = "0"
                self.car_expenses.value = "0"
                self.car_comment.value = ""

                self.show_snack_bar("🚗 Авто-сделка добавлена")
                await self.update_ui()
            else:
                self.show_snack_bar("❌ Не удалось добавить авто-сделку", ft.Colors.RED)

        except Exception as ex:
            self.show_snack_bar(f"❌ Ошибка при добавлении: {str(ex)}", ft.Colors.RED)

    async def on_save_capital(self, e):
        """Сохранение капитала"""
        try:
            new_capital = float(self.capital_input.value)
            self.db.update_initial_capital(new_capital)
            self.data['initial_capital'] = new_capital
            self.show_snack_bar("💾 Капитал обновлен")
            self.refresh_data()
        except ValueError:
            self.show_snack_bar("❌ Введите число!", ft.Colors.RED)

    async def on_import_excel(self, e):
        """Импорт из Excel"""
        self.show_snack_bar("📥 Функция импорта в разработке")

    async def on_export_excel(self, e):
        """Экспорт в Excel"""
        self.show_snack_bar("📤 Функция экспорта в разработке")

    async def on_monthly_filter_change(self, e):
        """Изменение фильтров месячного отчета"""
        try:
            self.selected_year = int(e.control.value) if e.control.label == "Год" else self.selected_year
            if e.control.label == "Месяц":
                months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
                self.selected_month = months.index(e.control.value) + 1

            self.update_monthly_report()
            await self.update_ui()
        except Exception as ex:
            print(f"Ошибка обновления фильтров: {ex}")

    async def update_ui(self):
        """Асинхронное обновление UI"""
        if self.page:
            self.page.update()


def main():
    ft.app(target=FletMoneyTrackerApp().main)


if __name__ == "__main__":
    main()