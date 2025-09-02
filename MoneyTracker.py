import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from calendar import monthrange

# Настройка тем - принудительно темная
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

import sqlite3
import pandas as pd
from typing import List, Dict


class DatabaseManager:
    def __init__(self, db_file="money_tracker.db"):
        self.db_name = db_file
        self.conn = sqlite3.connect(self.db_name, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                type TEXT NOT NULL,
                amount REAL NOT NULL,
                description TEXT NOT NULL,
                category TEXT NOT NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS car_deals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                brand TEXT NOT NULL,
                year TEXT NOT NULL,
                vin TEXT NOT NULL,
                comment TEXT,
                price REAL DEFAULT 0,
                cost REAL DEFAULT 0,
                expenses REAL DEFAULT 0,
                header REAL DEFAULT 0
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                initial_capital REAL NOT NULL DEFAULT 0
            )
        """)

        cursor.execute("SELECT COUNT(*) FROM settings")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO settings (initial_capital) VALUES (0)")

        self.conn.commit()

    # ---------------- Транзакции ----------------
    def add_transaction(self, transaction: Dict) -> int:
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO transactions (date, type, amount, description, category)
            VALUES (?, ?, ?, ?, ?)
        """, (
            transaction["date"],
            transaction["type"],
            transaction["amount"],
            transaction["description"],
            transaction["category"]
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_all_transactions(self) -> List[Dict]:
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM transactions ORDER BY date DESC")
        return [dict(row) for row in cursor.fetchall()]

    def update_transaction(self, transaction_id: int, updates: Dict) -> bool:
        if not updates:
            return False
        cursor = self.conn.cursor()
        set_clause = ", ".join(f"{key} = ?" for key in updates.keys())
        values = list(updates.values()) + [transaction_id]
        cursor.execute(f"UPDATE transactions SET {set_clause} WHERE id = ?", values)
        self.conn.commit()
        return cursor.rowcount > 0

    # ---------------- Авто-сделки ----------------
    def add_car_deal(self, car_deal: Dict) -> int:
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO car_deals (
                brand, year, vin, comment, price, cost, expenses, header
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            car_deal["brand"],
            car_deal["year"],
            car_deal["vin"],
            car_deal.get("comment", ""),
            car_deal.get("price", 0),
            car_deal.get("cost", 0),
            car_deal.get("expenses", 0),  # 👈 теперь расходы сохраняются
            car_deal.get("header", 0)
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_all_car_deals(self) -> List[Dict]:
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM car_deals ORDER BY year DESC")
        return [dict(row) for row in cursor.fetchall()]

    def update_car_deal(self, deal_id: int, updates: Dict) -> bool:
        if not updates:
            return False
        cursor = self.conn.cursor()
        set_clause = ", ".join(f"{key} = ?" for key in updates.keys())
        values = list(updates.values()) + [deal_id]
        cursor.execute(f"UPDATE car_deals SET {set_clause} WHERE id = ?", values)
        self.conn.commit()
        return cursor.rowcount > 0

    def exists_transaction(self, transaction: Dict) -> bool:
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) FROM transactions
            WHERE date = ? AND type = ? AND amount = ? AND description = ? AND category = ?
        """, (
            transaction["date"],
            transaction["type"],
            transaction["amount"],
            transaction["description"],
            transaction["category"]
        ))
        return cursor.fetchone()[0] > 0

    def exists_car_deal(self, car_deal: Dict) -> bool:
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) FROM car_deals
            WHERE brand = ? AND year = ? AND vin = ? AND price = ? AND cost = ? AND comment = ?
        """, (
            car_deal["brand"],
            car_deal["year"],
            car_deal["vin"],
            car_deal["price"],
            car_deal["cost"],
            car_deal.get("comment", "")
        ))
        return cursor.fetchone()[0] > 0


    # ---------------- Настройки ----------------
    def get_initial_capital(self) -> float:
        cursor = self.conn.cursor()
        cursor.execute("SELECT initial_capital FROM settings LIMIT 1")
        result = cursor.fetchone()
        return result[0] if result else 0.0

    def update_initial_capital(self, amount: float) -> bool:
        cursor = self.conn.cursor()
        cursor.execute("UPDATE settings SET initial_capital = ?", (amount,))
        self.conn.commit()
        return cursor.rowcount > 0

    # ---------------- Экспорт / импорт ----------------
    def export_to_excel(self, file_path: str) -> bool:
        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                # Экспорт транзакций
                transactions = self.get_all_transactions()
                if transactions:
                    pd.DataFrame(transactions).to_excel(
                        writer,
                        sheet_name="Транзакции",
                        index=False,
                        columns=["date", "type", "amount", "description", "category"]
                    )

                # Экспорт авто-сделок
                car_deals = self.get_all_car_deals()
                if car_deals:
                    # Создаем DataFrame с нужными колонками
                    df_car_deals = pd.DataFrame(car_deals)
                    # Переименовываем колонки для удобства
                    df_car_deals = df_car_deals.rename(columns={
                        "year": "Год",
                        "header": "Прибыль"
                    })
                    # Выбираем нужные колонки в правильном порядке
                    df_car_deals.to_excel(
                        writer,
                        sheet_name="Авто-сделки",
                        index=False,
                        columns=["brand", "Год", "vin", "price", "cost", "Прибыль", "comment"]
                    )

                # Экспорт настроек
                settings_data = {
                    "initial_capital": [self.get_initial_capital()]
                }
                pd.DataFrame(settings_data).to_excel(
                    writer, sheet_name="Настройки", index=False
                )
            return True
        except Exception as e:
            print(f"Ошибка при экспорте: {e}")
            return False

    def import_from_excel(self, file_path: str) -> bool:
        try:
            with pd.ExcelFile(file_path) as xls:
                if "Transactions" in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name="Transactions")
                    for _, row in df.iterrows():
                        self.add_transaction(row.to_dict())

                if "CarDeals" in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name="CarDeals")
                    for _, row in df.iterrows():
                        self.add_car_deal(row.to_dict())

                if "Config" in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name="Config")
                    if "initial_capital" in df.columns:
                        self.update_initial_capital(float(df.iloc[0]["initial_capital"]))
            return True
        except Exception as e:
            print(f"Ошибка при импорте: {e}")
            return False

    # ---------------- Закрытие соединения ----------------
    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None

class MoneyTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 Авто-Трекер Финансов v2.5")
        self.root.geometry("1300x900")

        self.large_font = ("Arial", 14)
        self.xlarge_font = ("Arial", 16, "bold")
        self.xxlarge_font = ("Arial", 18, "bold")

        # Инициализация менеджера базы данных
        self.db = DatabaseManager()

        # Загрузка данных
        self.transactions = self.db.get_all_transactions()
        self.car_deals = self.db.get_all_car_deals()
        self.initial_capital = self.db.get_initial_capital()

        self.setup_ui()

    def setup_ui(self):
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        style = ttk.Style()
        style.theme_create("custom_style", parent="alt", settings={
            "TNotebook": {"configure": {"background": "#f0f0f0", "tabmargins": [10, 5, 0, 0]}},
            "TNotebook.Tab": {
                "configure": {
                    "font": self.xlarge_font,
                    "padding": [30, 15],
                    "foreground": "#333333",
                    "background": "#e0e0e0",
                    "borderwidth": 1,
                },
                "map": {
                    "background": [("selected", "#f5f5f5")],
                    "expand": [("selected", [1, 1, 1, 0])]
                }
            }
        })
        style.theme_use("custom_style")

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True)

        self.add_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.add_frame, text="➕ Добавить операцию")
        self.setup_add_frame()

        self.car_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.car_frame, text="🚗 Авто-сделки")
        self.setup_car_frame()

        self.report_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.report_frame, text="📊 Финансовый отчет")
        self.setup_report_frame()

        self.monthly_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.monthly_frame, text="📅 Расходы за месяц")
        self.setup_monthly_frame()

        self.settings_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.settings_frame, text="⚙️ Настройки")
        self.setup_settings_frame()

        self.setup_context_menus()

    def setup_add_frame(self):
        self.add_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.add_frame, text="Новая операция", font=self.large_font).grid(row=0, column=0, columnspan=2,
                                                                                       pady=(0, 20))
        fields = [
            ("Тип операции:", "combobox", ["Приход", "Расход"], "Приход"),
            ("Сумма:", "entry", None, "0.00"),
            ("Описание:", "entry", None, ""),
            ("Категория:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные")
        ]

        self.entries = {}
        for row, (label, field_type, options, default) in enumerate(fields, start=1):
            lbl = ctk.CTkLabel(self.add_frame, text=label, font=self.large_font)
            lbl.grid(row=row, column=0, sticky="e", pady=5, padx=10)

            if field_type == "entry":
                entry = ctk.CTkEntry(self.add_frame)
                entry.insert(0, default)
            elif field_type == "combobox":
                entry = ctk.CTkComboBox(self.add_frame, values=options)
                entry.set(default)

            entry.grid(row=row, column=1, sticky="ew", pady=5, padx=10)
            self.entries[label] = entry

        ctk.CTkButton(
            self.add_frame,
            text="Добавить операцию",
            command=self.add_transaction,
            fg_color="#4CAF50",
            hover_color="#45a049",
            height=40
        ).grid(row=len(fields) + 1, column=0, columnspan=2, pady=20, sticky="we")

    def setup_car_frame(self):
        self.car_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.car_frame, text="Новая авто-сделка", font=self.large_font).grid(
            row=0, column=0, columnspan=2, pady=(0, 20))

        car_fields = [
            ("Марка:", "entry", None, ""),
            ("Год:", "entry", None, ""),
            ("VIN:", "entry", None, ""),
            ("Цена продажи с учетом опций:", "entry", None, "0"),
            ("Закупочная стоимость:", "entry", None, "0"),
            ("Расходы:", "entry", None, "0"),
            ("Комментарий:", "entry", None, "")
        ]

        self.car_entries = {}
        for row, (label, field_type, options, default) in enumerate(car_fields, start=1):
            ctk.CTkLabel(self.car_frame, text=label).grid(
                row=row, column=0, padx=10, pady=5, sticky="e")

            entry = ctk.CTkEntry(self.car_frame)
            entry.insert(0, default)
            entry.grid(row=row, column=1, padx=10, pady=5, sticky="we")
            self.car_entries[label.replace(":", "")] = entry

        ctk.CTkButton(
            self.car_frame,
            text="Добавить авто-сделку",
            command=self.add_car_deal,
            fg_color="#2196F3",
            hover_color="#0b7dda",
            height=40
        ).grid(row=len(car_fields) + 1, column=0, columnspan=2, pady=20, sticky="we")

    def setup_monthly_frame(self):
        self.monthly_frame.grid_columnconfigure(0, weight=1)
        self.monthly_frame.grid_rowconfigure(2, weight=1)

        # Заголовок
        ctk.CTkLabel(self.monthly_frame, text="Анализ расходов по месяцам",
                     font=self.xxlarge_font).grid(row=0, column=0, pady=(10, 20))

        # Фрейм для выбора года и месяца
        control_frame = ctk.CTkFrame(self.monthly_frame, fg_color="#363636", height=60)
        control_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_propagate(False)

        # Внутренний фрейм для элементов управления
        inner_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0, sticky="w", padx=20, pady=10)

        # Выбор года
        ctk.CTkLabel(inner_frame, text="Год:", font=self.xlarge_font).grid(
            row=0, column=0, padx=(0, 10), pady=5, sticky="w")

        current_year = datetime.now().year
        years = [str(year) for year in range(current_year - 2, current_year + 1)]
        self.year_combo = ctk.CTkComboBox(inner_frame, values=years, width=100, height=40,
                                          dropdown_font=self.large_font, font=self.xlarge_font)
        self.year_combo.set(str(current_year))
        self.year_combo.grid(row=0, column=1, padx=(0, 30), pady=5, sticky="w")
        self.year_combo.configure(command=lambda event: self.update_monthly_report())

        # Выбор месяца
        ctk.CTkLabel(inner_frame, text="Месяц:", font=self.xlarge_font).grid(
            row=0, column=2, padx=(0, 10), pady=5, sticky="w")

        months = [
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ]
        current_month = datetime.now().month - 1
        self.month_combo = ctk.CTkComboBox(inner_frame, values=months, width=150, height=40,
                                           dropdown_font=self.large_font, font=self.xlarge_font)
        self.month_combo.set(months[current_month])
        self.month_combo.grid(row=0, column=3, padx=(0, 20), pady=5, sticky="w")
        self.month_combo.configure(command=lambda event: self.update_monthly_report())

        # Таблица с ежедневной сводкой
        daily_columns = {
            "#1": {"name": "date", "text": "Дата", "width": 100, "anchor": "center"},
            "#2": {"name": "income", "text": "Приход", "width": 120, "anchor": "e"},
            "#3": {"name": "expense", "text": "Расход", "width": 120, "anchor": "e"},
            "#4": {"name": "balance", "text": "Итог в кассе", "width": 140, "anchor": "e"},
            "#5": {"name": "transactions", "text": "Операций", "width": 80, "anchor": "center"}
        }

        self.daily_tree = ttk.Treeview(self.monthly_frame, columns=list(daily_columns.keys()), show="headings",
                                       height=8)
        for col, params in daily_columns.items():
            self.daily_tree.heading(col, text=params["text"])
            self.daily_tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

        # Детальная таблица операций
        detail_columns = {
            "#1": {"name": "date", "text": "Дата", "width": 150, "anchor": "center"},
            "#2": {"name": "type", "text": "Тип", "width": 80, "anchor": "center"},
            "#3": {"name": "description", "text": "Описание", "width": 250, "anchor": "w"},
            "#4": {"name": "category", "text": "Категория", "width": 120, "anchor": "center"},
            "#5": {"name": "amount", "text": "Сумма", "width": 120, "anchor": "e"}
        }

        self.detail_tree = ttk.Treeview(self.monthly_frame, columns=list(detail_columns.keys()), show="headings",
                                        height=12)
        for col, params in detail_columns.items():
            self.detail_tree.heading(col, text=params["text"])
            self.detail_tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

        # Полосы прокрутки
        scrollbar_y1 = ttk.Scrollbar(self.monthly_frame, orient="vertical")
        scrollbar_y2 = ttk.Scrollbar(self.monthly_frame, orient="vertical")

        self.daily_tree.configure(yscrollcommand=scrollbar_y1.set)
        self.detail_tree.configure(yscrollcommand=scrollbar_y2.set)
        scrollbar_y1.configure(command=self.daily_tree.yview)
        scrollbar_y2.configure(command=self.detail_tree.yview)

        # Размещение элементов
        self.daily_tree.grid(row=2, column=0, sticky="nsew", padx=(10, 0), pady=(0, 5))
        scrollbar_y1.grid(row=2, column=1, sticky="ns", pady=(0, 5))

        # Заголовок детальной таблицы
        ctk.CTkLabel(self.monthly_frame, text="Детализация операций по дням:",
                     font=self.xlarge_font).grid(row=3, column=0, sticky="w", padx=10, pady=(10, 5))

        self.detail_tree.grid(row=4, column=0, sticky="nsew", padx=(10, 0), pady=(0, 10))
        scrollbar_y2.grid(row=4, column=1, sticky="ns", pady=(0, 10))

        # Панель статистики месяца
        stats_frame = ctk.CTkFrame(self.monthly_frame)
        stats_frame.grid(row=5, column=0, columnspan=2, sticky="we", padx=10, pady=10)

        self.stats_labels = {
            "total_income": ctk.CTkLabel(stats_frame, text="Общий приход: 0.00 ₽", font=self.xlarge_font),
            "total_expense": ctk.CTkLabel(stats_frame, text="Общий расход: 0.00 ₽", font=self.xlarge_font),
            "month_balance": ctk.CTkLabel(stats_frame, text="Итог за месяц: 0.00 ₽", font=self.xxlarge_font),
            "days_count": ctk.CTkLabel(stats_frame, text="Дней с операциями: 0", font=self.large_font)
        }

        for i, (key, label) in enumerate(self.stats_labels.items()):
            label.grid(row=0, column=i, padx=15, pady=5)

        # Настройка весов для растягивания
        self.monthly_frame.grid_rowconfigure(2, weight=1)
        self.monthly_frame.grid_rowconfigure(4, weight=2)

        # Привязка события выбора дня
        self.daily_tree.bind("<<TreeviewSelect>>", self.on_day_selected)

        # Первоначальное обновление отчета
        self.update_monthly_report()

    def update_monthly_report(self, event=None):
        """Обновляет отчет о расходах за месяц"""
        # Очищаем таблицы
        for item in self.daily_tree.get_children():
            self.daily_tree.delete(item)
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        # Получаем выбранные год и месяц
        try:
            selected_year = int(self.year_combo.get())
            selected_month = self.month_combo.get()
            month_number = self.get_month_number(selected_month)
        except (ValueError, AttributeError):
            return

        # Группируем операции по дням
        daily_data = {}
        total_income = 0
        total_expense = 0

        for transaction in self.transactions:
            try:
                # Парсим дату
                date_str = transaction["date"].split()[0]
                day, month, year = map(int, date_str.split('.'))

                if year == selected_year and month == month_number:
                    if date_str not in daily_data:
                        daily_data[date_str] = {
                            'income': 0,
                            'expense': 0,
                            'transactions': []
                        }

                    amount = abs(transaction["amount"])
                    if transaction["type"] == "Приход":
                        daily_data[date_str]['income'] += amount
                        total_income += amount
                    else:
                        daily_data[date_str]['expense'] += amount
                        total_expense += amount

                    daily_data[date_str]['transactions'].append(transaction)

            except (ValueError, IndexError):
                continue

        # Заполняем таблицу ежедневной сводки
        for date_str in sorted(daily_data.keys(), reverse=True):
            data = daily_data[date_str]
            balance = data['income'] - data['expense']
            transactions_count = len(data['transactions'])

            self.daily_tree.insert(
                "",
                "end",
                values=(
                    date_str,
                    f"{data['income']:,.2f}",
                    f"{data['expense']:,.2f}",
                    f"{balance:,.2f}",
                    transactions_count
                )
            )

        # Обновляем статистику месяца
        month_balance = total_income - total_expense
        days_with_operations = len(daily_data)

        self.stats_labels["total_income"].configure(text=f"Общий приход: {total_income:,.2f} ₽")
        self.stats_labels["total_expense"].configure(text=f"Общий расход: {total_expense:,.2f} ₽")
        self.stats_labels["month_balance"].configure(text=f"Итог за месяц: {month_balance:,.2f} ₽")
        self.stats_labels["days_count"].configure(text=f"Дней с операциями: {days_with_operations}")

    def on_day_selected(self, event):
        """Обработчик выбора дня в таблице"""
        selected_items = self.daily_tree.selection()
        if not selected_items:
            return

        selected_item = selected_items[0]
        date_str = self.daily_tree.item(selected_item, "values")[0]

        # Очищаем детальную таблицу
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        # Получаем выбранные год и месяц
        try:
            selected_year = int(self.year_combo.get())
            selected_month = self.month_combo.get()
            month_number = self.get_month_number(selected_month)
        except (ValueError, AttributeError):
            return

        # Заполняем детальную таблицу операциями за выбранный день
        for transaction in self.transactions:
            try:
                transaction_date = transaction["date"].split()[0]
                day, month, year = map(int, transaction_date.split('.'))

                if (year == selected_year and month == month_number and
                        transaction_date == date_str):
                    self.detail_tree.insert(
                        "",
                        "end",
                        values=(
                            transaction["date"],
                            transaction["type"],
                            transaction["description"],
                            transaction["category"],
                            f"{abs(transaction['amount']):,.2f} ₽"
                        )
                    )

            except (ValueError, IndexError):
                continue

    # Добавляем метод для обновления данных при изменении транзакций
    def refresh_data(self):
        """Обновляет все данные из базы и перерисовывает отчеты"""
        # Загружаем свежие данные из базы
        self.transactions = self.db.get_all_transactions()
        self.car_deals = self.db.get_all_car_deals()
        self.initial_capital = self.db.get_initial_capital()

        # Обновляем все отчеты
        self.update_report()
        self.update_monthly_report()

    def get_month_number(self, month_name):
        months = {
            "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4,
            "Май": 5, "Июнь": 6, "Июль": 7, "Август": 8,
            "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12
        }
        return months.get(month_name, datetime.now().month)

    def parse_date(self, date_str):
        try:
            # Парсим дату из формата "dd.mm.yyyy HH:MM"
            date_part = date_str.split()[0]
            return datetime.strptime(date_part, "%d.%m.%Y")
        except:
            return datetime.now()

    def update_monthly_chart(self, expenses, year, month):
        """Обновляет график расходов по дням месяца"""
        self.ax.clear()

        # Группируем расходы по дням
        daily_totals = {}
        for expense in expenses:
            try:
                date_str = expense["date"].split()[0]
                day = int(date_str.split('.')[0])
                amount = abs(expense["amount"])

                if day not in daily_totals:
                    daily_totals[day] = 0
                daily_totals[day] += amount
            except (ValueError, IndexError):
                continue

        # Создаем данные для графика
        days = list(range(1, monthrange(year, month)[1] + 1))
        amounts = [daily_totals.get(day, 0) for day in days]

        # Строим график
        self.ax.bar(days, amounts, alpha=0.7, color='red')
        self.ax.set_xlabel('День месяца', fontsize=12)
        self.ax.set_ylabel('Сумма расходов (₽)', fontsize=12)
        self.ax.set_title(f'Расходы за {self.month_combo.get()} {year}', fontsize=14)
        self.ax.grid(True, alpha=0.3)

        # Форматирование осей
        self.ax.set_xlim(0.5, len(days) + 0.5)
        self.ax.set_xticks(days[::2])  # Каждый второй день

        self.fig.tight_layout()
        self.canvas.draw()

    def setup_report_frame(self):
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_rowconfigure(1, weight=1)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=self.large_font)
        style.configure("Treeview", font=self.large_font, rowheight=35)

        columns = {
            "#1": {"name": "date", "text": "Дата", "width": 180, "anchor": "center"},
            "#2": {"name": "type", "text": "Тип", "width": 120, "anchor": "center"},
            "#3": {"name": "amount", "text": "Сумма", "width": 150, "anchor": "e"},
            "#4": {"name": "description", "text": "Описание", "width": 300, "anchor": "center"},
            "#5": {"name": "category", "text": "Категория", "width": 150, "anchor": "center"}
        }

        self.tree = ttk.Treeview(self.report_frame, columns=list(columns.keys()), show="headings")
        for col, params in columns.items():
            self.tree.heading(col, text=params["text"])
            self.tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

        car_columns = {
            "#1": {"name": "brand", "text": "Марка", "width": 120, "anchor": "center"},
            "#2": {"name": "model_year", "text": "Год", "width": 80, "anchor": "center"},
            "#3": {"name": "vin", "text": "VIN", "width": 150, "anchor": "center"},
            "#4": {"name": "price", "text": "Цена продажи", "width": 120, "anchor": "e"},
            "#5": {"name": "cost", "text": "Закупочная стоимость", "width": 120, "anchor": "e"},
            "#6": {"name": "expenses", "text": "Расходы", "width": 120, "anchor": "e"},  # Добавлен столбец расходов
            "#7": {"name": "profit", "text": "Прибыль", "width": 120, "anchor": "e"},
            "#8": {"name": "comment", "text": "Комментарий", "width": 200, "anchor": "center"}
        }

        self.car_tree = ttk.Treeview(self.report_frame, columns=list(car_columns.keys()), show="headings")
        for col, params in car_columns.items():
            self.car_tree.heading(col, text=params["text"])
            self.car_tree.column(col, width=params["width"], anchor=params.get("anchor", "center"))
        scrollbar = ttk.Scrollbar(self.report_frame, orient="vertical")
        car_scrollbar = ttk.Scrollbar(self.report_frame, orient="vertical")

        self.tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.car_tree.grid(row=3, column=0, sticky="nsew", padx=10, pady=(20, 10))
        scrollbar.grid(row=1, column=1, sticky="ns")
        car_scrollbar.grid(row=3, column=1, sticky="ns")

        self.tree.configure(yscrollcommand=scrollbar.set)
        self.car_tree.configure(yscrollcommand=car_scrollbar.set)
        scrollbar.configure(command=self.tree.yview)
        car_scrollbar.configure(command=self.car_tree.yview)

        self.summary_panel = ctk.CTkFrame(self.report_frame)
        self.summary_panel.grid(row=4, column=0, columnspan=2, sticky="we", padx=10, pady=10)

        summary_items = [
            ("initial_capital", "Стартовый капитал:", f"{self.initial_capital:,.2f} ₽", self.xlarge_font),
            ("total_income", "Общий приход:", "0.00 ₽", self.xlarge_font),
            ("total_expense", "Общий расход:", "0.00 ₽", self.xlarge_font),
            ("additional_investment", "Доп. вложения:", "0.00 ₽", self.xlarge_font),
            ("car_profit", "Прибыль с авто:", "0.00 ₽", self.xxlarge_font),
            ("total_profit", "Общая прибыль:", "0.00 ₽", self.xxlarge_font)
        ]

        self.summary_labels = {}
        for i, (name, text, value, font) in enumerate(summary_items):
            frame = ctk.CTkFrame(self.summary_panel, width=220)
            frame.grid(row=0, column=i, padx=5)
            ctk.CTkLabel(frame, text=text, font=self.large_font).pack()
            self.summary_labels[name] = ctk.CTkLabel(frame, text=value, font=font)
            self.summary_labels[name].pack()

        self.update_report()

        key_order = ["date", "type", "amount", "description", "category"]
        car_key_order = ["brand", "year", "vin", "price", "cost", "expenses", "header", "comment"]


        self.tree.bind(
            "<Double-1>",
            lambda event: self._handle_treeview_double_click(
                event,
                self.tree,
                lambda tree, item, col: self.edit_cell(tree, item, col, self.transactions, key_order)
            )
        )

        self.car_tree.bind(
            "<Double-1>",
            lambda event: self._handle_treeview_double_click(
                event,
                self.car_tree,
                lambda tree, item, col: self.edit_cell(tree, item, col, self.car_deals, car_key_order)
            )
        )

    def setup_settings_frame(self):
        ctk.CTkLabel(self.settings_frame, text="Стартовый капитал:", font=self.large_font).pack(pady=(20, 5))
        self.capital_entry = ctk.CTkEntry(self.settings_frame)
        self.capital_entry.insert(0, str(self.initial_capital))
        self.capital_entry.pack()

        def save_capital():
            try:
                self.initial_capital = float(self.capital_entry.get())
                self.db.update_initial_capital(self.initial_capital)
                messagebox.showinfo("Сохранено", "Стартовый капитал обновлён.")
                self.update_report()
            except ValueError:
                messagebox.showerror("Ошибка", "Введите число.")

        ctk.CTkButton(self.settings_frame, text="💾 Сохранить капитал", command=save_capital).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📥 Импорт из Excel", command=self.import_from_excel).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📤 Экспорт в Excel", command=self.export_to_excel).pack(pady=10)

    def get_data_list_by_iid(self, iid):
        if iid.startswith("tr_"):
            return self.transactions
        elif iid.startswith("car_"):
            return self.car_deals
        else:
            return None

    def edit_cell(self, tree, item, col_index, data_list, key_order):
        x, y, width, height = tree.bbox(item, f"#{col_index + 1}")
        value = tree.item(item, "values")[col_index]

        entry = ctk.CTkEntry(tree, width=width, height=height)
        entry.insert(0, value)
        entry.place(x=x, y=y)

        def save_edit(event=None):
            new_val = entry.get()
            values = list(tree.item(item, "values"))
            values[col_index] = new_val
            tree.item(item, values=values)
            entry.destroy()

            iid = item
            data_list = self.get_data_list_by_iid(iid)
            if data_list is None:
                return

            if iid.startswith("tr_") or iid.startswith("car_"):
                index = int(iid.split("_")[1])
            else:
                index = int(iid)

            if index < len(data_list) and col_index < len(key_order):
                key = key_order[col_index]
                cleaned = new_val.replace(",", "")
                try:
                    if key in ["amount", "price", "cost", "header"]:
                        cleaned = float(cleaned)
                except ValueError:
                    pass

                data_list[index][key] = cleaned

            # Приход/расход
            if iid.startswith("tr_"):
                self.db.update_transaction(data_list[index]["id"], {key: cleaned})

            # Авто-сделка
            elif iid.startswith("car_"):
                updates = {}
                if key == "price":
                    # При изменении цены пересчитываем прибыль
                    price = float(cleaned)
                    cost = float(data_list[index].get("cost", 0))
                    header = price - cost
                    updates = {
                        "price": price,
                        "header": header
                    }
                elif key == "cost":
                    # При изменении стоимости пересчитываем прибыль
                    cost = float(cleaned)
                    price = float(data_list[index].get("price", 0))
                    header = price - cost
                    updates = {
                        "cost": cost,
                        "header": header
                    }
                elif key == "expenses":
                    expenses = float(cleaned)
                    price = float(data_list[index].get("price", 0))
                    cost = float(data_list[index].get("cost", 0))
                    header = price - cost - expenses  # Правильный расчет прибыли с учетом расходов
                    updates = {
                        "expenses": expenses,
                        "header": header
                    }

                else:
                    updates = {key: cleaned}

                # Обновляем данные в списке
                data_list[index].update(updates)

                # Обновляем в базе данных
                self.db.update_car_deal(data_list[index]["id"], updates)

                # Обновляем таблицу
                self.car_tree.delete(*self.car_tree.get_children())
                for i, deal in enumerate(self.car_deals):
                    self.car_tree.insert(
                        "", "end", iid=f"car_{i}",
                        values=(
                            deal.get("brand", ""),
                            deal.get("year", ""),
                            deal.get("vin", ""),
                            f"{deal.get('price', 0):,.2f}",
                            f"{deal.get('cost', 0):,.2f}",
                            f"{deal.get('expenses', 0):,.2f}"  # Этой строки не было!
                            f"{deal.get('header', 0):,.2f}",
                            deal.get("comment", "")
                        )
                    )

            self.update_report()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.focus_set()

    def _handle_treeview_double_click(self, event, tree, edit_callback):
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item or not column:
            return
        col_index = int(column[1:]) - 1
        edit_callback(tree, item, col_index)

    def add_transaction(self):
        try:
            operation = self.entries["Тип операции:"].get()
            amount = float(self.entries["Сумма:"].get())
            description = self.entries["Описание:"].get().strip()
            category = self.entries["Категория:"].get()

            if not description:
                messagebox.showerror("Ошибка", "Введите описание операции!")
                return
            if amount <= 0:
                messagebox.showerror("Ошибка", "Сумма должна быть положительной!")
                return

            transaction = {
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                "type": operation,
                "amount": amount if operation == "Приход" else -amount,
                "description": description,
                "category": category
            }

            self.db.add_transaction(transaction)

            # Обновляем данные и все отчеты
            self.refresh_data()

            self.entries["Сумма:"].delete(0, tk.END)
            self.entries["Сумма:"].insert(0, "0.00")
            self.entries["Описание:"].delete(0, tk.END)

            messagebox.showinfo("Успех", "Операция успешно добавлена!")

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму!")

    def add_car_deal(self):
        try:
            brand = self.car_entries["Марка"].get().strip()
            year = self.car_entries["Год"].get().strip()
            vin = self.car_entries["VIN"].get().strip()
            price = float(self.car_entries["Цена продажи с учетом опций"].get() or 0)
            cost = float(self.car_entries["Закупочная стоимость"].get() or 0)
            expenses = float(self.car_entries["Расходы"].get() or 0)
            comment = self.car_entries["Комментарий"].get().strip()

            profit = price - cost - expenses

            if not brand:
                messagebox.showerror("Ошибка", "Введите марку авто!")
                return

            car_deal = {
                "brand": brand,
                "year": year,
                "vin": vin,
                "price": price,
                "cost": cost,
                "expenses": expenses,
                "header": profit,
                "comment": comment
            }

            self.db.add_car_deal(car_deal)

            # Обновляем данные и все отчеты
            self.refresh_data()

            messagebox.showinfo("Успех", "Авто-сделка успешно добавлена!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении: {str(e)}")

    def update_report(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.car_tree.get_children():
            self.car_tree.delete(item)

        for i, tr in enumerate(self.transactions):
            self.tree.insert(
                "",
                "end",
                iid=f"tr_{i}",
                values=(
                    tr["date"],
                    tr["type"],
                    f"{abs(tr['amount']):,.2f}",
                    tr["description"],
                    tr["category"]
                )
            )

        for i, deal in enumerate(self.car_deals):
            self.car_tree.insert(
                "", "end", iid=f"car_{i}",
                values=(
                    deal.get("brand", ""),
                    deal.get("year", ""),
                    deal.get("vin", ""),
                    f"{deal.get('price', 0):,.2f}",
                    f"{deal.get('cost', 0):,.2f}",
                    f"{deal.get('expenses', 0):,.2f}",  # Исправлено: добавлены расходы
                    f"{deal.get('header', 0):,.2f}",
                    deal.get("comment", "")
                )
            )

        self.update_summary()

    def update_summary(self):
        total_income = sum(t["amount"] for t in self.transactions if t["type"] == "Приход")
        total_expense = abs(sum(t["amount"] for t in self.transactions if t["type"] == "Расход"))

        additional_investment = max(0, total_expense - self.initial_capital)
        car_profit = sum(deal.get("header", 0) for deal in self.car_deals)
        total_profit = car_profit + total_income - additional_investment

        self.summary_labels["initial_capital"].configure(text=f"{self.initial_capital:,.2f} ₽")
        self.summary_labels["total_income"].configure(text=f"{total_income:,.2f} ₽")
        self.summary_labels["total_expense"].configure(text=f"{total_expense:,.2f} ₽")
        self.summary_labels["additional_investment"].configure(text=f"{additional_investment:,.2f} ₽")
        self.summary_labels["car_profit"].configure(text=f"{car_profit:,.2f} ₽")
        self.summary_labels["total_profit"].configure(text=f"{total_profit:,.2f} ₽")

    def import_from_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        try:
            with pd.ExcelFile(path) as xls:
                # Получаем список всех листов в файле
                sheet_names = xls.sheet_names

                # Импорт транзакций (проверяем разные варианты названий)
                transaction_sheet = None
                for name in ['Транзакции', 'Transactions']:
                    if name in sheet_names:
                        transaction_sheet = name
                        break

                if transaction_sheet:
                    df = pd.read_excel(xls, sheet_name=transaction_sheet)
                    for _, row in df.iterrows():
                        try:
                            transaction = {
                                "date": row.get("date", datetime.now().strftime("%d.%m.%Y %H:%M")),
                                "type": row.get("type", "Приход"),
                                "amount": float(row.get("amount", 0)),
                                "description": str(row.get("description", "")),
                                "category": row.get("category", "Наличные")
                            }
                            if not self.db.exists_transaction(transaction):
                                self.db.add_transaction(transaction)

                        except Exception as e:
                            print(f"Ошибка при импорте транзакции: {e}")

                # Импорт авто-сделок (проверяем разные варианты названий)
                car_sheet = None
                for name in ['Площадка', 'Авто-сделки', 'CarDeals']:
                    if name in sheet_names:
                        car_sheet = name
                        break

                if car_sheet:
                    df = pd.read_excel(xls, sheet_name=car_sheet)
                    for _, row in df.iterrows():
                        try:
                            # Проверяем обязательное поле brand (разные варианты названий)
                            brand = str(row.get("brand", row.get("Марка", row.get("Модель", ""))).strip())
                            if not brand:
                                continue  # Пропускаем записи без марки авто

                            year = str(row.get("year", row.get("Год", ""))).strip()
                            vin = str(row.get("vin", row.get("VIN", ""))).strip()
                            price = float(row.get("price", row.get("Цена", 0)))
                            cost = float(row.get("cost", row.get("Стоимость", 0)))
                            profit = float(row.get("profit", row.get("Прибыль", row.get("header", price - cost))))
                            comment = str(row.get("comment", row.get("Комментарий", "")))

                            car_deal = {
                                "brand": brand,
                                "year": year,
                                "vin": vin,
                                "price": price,
                                "cost": cost,
                                "header": profit,
                                "comment": comment
                            }
                            if not self.db.exists_car_deal(car_deal):
                                self.db.add_car_deal(car_deal)

                        except Exception as e:
                            print(f"Ошибка при импорте авто-сделки: {e}")

                # Импорт настроек (проверяем разные варианты названий)
                settings_sheet = None
                for name in ['Настройки', 'Settings', 'Config']:
                    if name in sheet_names:
                        settings_sheet = name
                        break

                if settings_sheet:
                    df = pd.read_excel(xls, sheet_name=settings_sheet)
                    if "initial_capital" in df.columns:
                        try:
                            capital = float(df.iloc[0]["initial_capital"])
                            self.db.update_initial_capital(capital)
                            self.initial_capital = capital
                        except:
                            pass

            # Обновляем данные из базы
            self.transactions = self.db.get_all_transactions()
            self.car_deals = self.db.get_all_car_deals()
            self.initial_capital = self.db.get_initial_capital()

            # Обновление поля капитала
            self.capital_entry.delete(0, tk.END)
            self.capital_entry.insert(0, str(self.initial_capital))

            # Полностью обновляем отчёт
            self.update_report()

            messagebox.showinfo("Успех", "Данные успешно импортированы из Excel.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при импорте: {str(e)}")

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить отчет как"
        )
        if not path:
            return

        try:
            if self.db.export_to_excel(path):
                messagebox.showinfo("Успех", "Данные успешно экспортированы в Excel.")
            else:
                messagebox.showerror("Ошибка", "Не удалось экспортировать данные")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {str(e)}")

    def setup_context_menus(self):
        # Контекстное меню для таблицы транзакций
        self.transaction_menu = tk.Menu(self.root, tearoff=0)
        self.transaction_menu.add_command(label="Удалить", command=self.delete_selected_transaction)

        # Контекстное меню для таблицы авто-сделок
        self.car_menu = tk.Menu(self.root, tearoff=0)
        self.car_menu.add_command(label="Удалить", command=self.delete_selected_car_deal)

        # Привязка меню к таблицам
        self.tree.bind("<Button-3>", self.show_transaction_menu)
        self.car_tree.bind("<Button-3>", self.show_car_menu)

    def show_transaction_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.transaction_menu.post(event.x_root, event.y_root)

    def show_car_menu(self, event):
        item = self.car_tree.identify_row(event.y)
        if item:
            self.car_tree.selection_set(item)
            self.car_menu.post(event.x_root, event.y_root)

    def delete_selected_transaction(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item = selected_item[0]
        if not item.startswith("tr_"):
            return

        if not messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить эту транзакцию?"):
            return

        index = int(item.split("_")[1])
        transaction_id = self.transactions[index]["id"]

        # Удаляем из базы данных
        cursor = self.db.conn.cursor()
        cursor.execute("DELETE FROM transactions WHERE id = ?", (transaction_id,))
        self.db.conn.commit()

        # Обновляем данные и все отчеты
        self.refresh_data()

        messagebox.showinfo("Успех", "Транзакция успешно удалена")

    def delete_selected_car_deal(self):
        selected_item = self.car_tree.selection()
        if not selected_item:
            return

        item = selected_item[0]
        if not item.startswith("car_"):
            return

        if not messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить эту авто-сделку?"):
            return

        index = int(item.split("_")[1])
        deal_id = self.car_deals[index]["id"]

        # Удаляем из базы данных
        cursor = self.db.conn.cursor()
        cursor.execute("DELETE FROM car_deals WHERE id = ?", (deal_id,))
        self.db.conn.commit()

        # Обновляем данные и все отчеты
        self.refresh_data()

        messagebox.showinfo("Успех", "Авто-сделка успешно удалена")

if __name__ == "__main__":
    root = ctk.CTk()
    app = MoneyTrackerApp(root)
    root.mainloop()