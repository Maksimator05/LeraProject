import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import sys
import os
from typing import Dict
import pandas as pd
# Добавляем родительскую директорию в путь для импорта
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.database import DatabaseManager

# Настройка тем - принудительно темная
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

class Toast(ctk.CTkToplevel):
    def __init__(self, parent, message, duration=2500):
        super().__init__(parent)
        self.duration = duration

        # Настраиваем окно
        self.overrideredirect(True) # Убираем заголовок
        self.attributes("-alpha", 0.0) # Полупрозрачность
        self.attributes("-topmost", True) # Поверх всех окон

        # Создаем сообщение
        self.label = ctk.CTkLabel(
            self,
            text=message,
            corner_radius=10,
            text_color=("#000000", "#FFFFFF"),
            font=("Arial", 14, "bold"),
            padx=20,
            pady=10
        )
        self.label.pack()

        # Вычисляем позицию (центр снизу)
        self.update_idletasks()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        toast_width = self.winfo_width()
        x = parent.winfo_rootx() + (parent_width - toast_width) // 2
        y = parent.winfo_rooty() + parent_height - 100  # 100px от нижнего края

        self.geometry(f"+{x}+{y}")

        # Анимация появления
        self.fade_in()

        # Планируем закрытие
        self.after(self.duration, self.fade_out)

    def fade_in(self):
        """Плавное появление"""
        current_alpha = self.attributes("-alpha")
        if current_alpha < 0.9:
            self.attributes("-alpha", current_alpha + 0.1)
            self.after(20, self.fade_in)
        else:
            self.attributes("-alpha", 0.9)

    def fade_out(self):
        """Плавное исчезновение"""
        current_alpha = self.attributes("-alpha")
        if current_alpha > 0:
            self.attributes("-alpha", current_alpha - 0.1)
            self.after(20, self.fade_out)
        else:
            self.destroy()

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

    def ask_confirmation(self, title, message):
        """Стильное окно подтверждения вместо стандартного messagebox"""
        dialog = ctk.CTkToplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set()  # Блокирует главное окно

        # Центрируем диалог
        dialog.transient(self.root)
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        # Содержимое диалога
        ctk.CTkLabel(dialog, text=message, font=self.large_font,
                     wraplength=350).pack(pady=20, padx=20)

        result = [False]  # Используем список для передачи по ссылке

        def on_yes():
            result[0] = True
            dialog.destroy()

        def on_no():
            result[0] = False
            dialog.destroy()

        # Фрейм для кнопок
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=10)

        ctk.CTkButton(button_frame, text="✅ Да", command=on_yes,
                      fg_color="#4CAF50", hover_color="#45a049",
                      width=100).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="❌ Нет", command=on_no,
                      fg_color="#f44336", hover_color="#da190b",
                      width=100).pack(side="right", padx=10)

        # Ждем закрытия диалога
        self.root.wait_window(dialog)
        return result[0]

    def show_toast(self, message, duration=2500):
        """Показывает временное уведомление в правом нижнем углу"""
        toast = Toast(self.root, message, duration)

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

        # Сначала создаем все фреймы
        self.add_frame = ctk.CTkFrame(self.notebook)
        self.car_frame = ctk.CTkFrame(self.notebook)
        self.report_frame = ctk.CTkFrame(self.notebook)
        self.monthly_frame = ctk.CTkFrame(self.notebook)
        self.settings_frame = ctk.CTkFrame(self.notebook)

        # Затем настраиваем их (теперь monthly_frame будет создан до report_frame)
        self.setup_add_frame()
        self.setup_car_frame()
        self.setup_monthly_frame()  # Сначала создаем monthly_frame
        self.setup_report_frame()  # Затем report_frame
        self.setup_settings_frame()

        # Добавляем вкладки после настройки
        self.notebook.add(self.add_frame, text="➕ Добавить операцию")
        self.notebook.add(self.car_frame, text="🚗 Авто-сделки")
        self.notebook.add(self.report_frame, text="📊 Финансовый отчет")
        self.notebook.add(self.monthly_frame, text="📅 Расходы за месяц")
        self.notebook.add(self.settings_frame, text="⚙️ Настройки")

        self.setup_context_menus()

    def setup_add_frame(self):
        self.add_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.add_frame, text="Новая операция", font=self.large_font).grid(row=0, column=0, columnspan=2,
                                                                                       pady=(0, 20))

        # Тип операции
        fields = [
            ("Тип операции:", "combobox", ["Приход", "Расход"], "Приход"),
            ("Сумма:", "entry", None, "0.00"),
            ("Описание:", "entry", None, ""),
            ("Тип оплаты:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные"),
            ("Категория:", "combobox", [
                "КЦ",
                "Реклама",
                "Вед.рекламы",
                "Комиссия брок",
                "Дилерство",
                "Аренда",
                "ЗП окладники",
                "ЗП проценты",
                "Другое"
            ], "Другое")
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

        # Переменная для хранения состояния "исключить из расхода"
        self.exclude_from_total = tk.BooleanVar(value=False)

        # Чекбокс для исключения из общего расхода
        exclude_checkbox = ctk.CTkCheckBox(
            self.add_frame,
            text="Исключить из общего расхода",
            variable=self.exclude_from_total,
            font=self.large_font
        )
        exclude_checkbox.grid(row=len(fields) + 1, column=0, columnspan=2, pady=10, padx=10, sticky="w")

        # ОДНА кнопка добавления операции
        ctk.CTkButton(
            self.add_frame,
            text="Добавить операцию",
            command=self.add_transaction,
            fg_color="#4CAF50",
            hover_color="#45a049",
            height=40
        ).grid(row=len(fields) + 2, column=0, columnspan=2, pady=20, sticky="we")

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
            fg_color="#b674ec",
            hover_color="#c78df7",
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

        # Добавьте после выбора месяца
        info_label = ctk.CTkLabel(
            inner_frame,
            text="💡 Данные месяца включаются в полный экспорт (кнопка в Настройках)",
            text_color="#FFD700",
            font=("Arial", 12)
        )
        info_label.grid(row=0, column=4, padx=(20, 0), pady=5, sticky="w")

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
            "#3": {"name": "description", "text": "Описание", "width": 250, "anchor": "center"},
            "#4": {"name": "category", "text": "Категория", "width": 120, "anchor": "center"},
            "#5": {"name": "amount", "text": "Суммa", "width": 120, "anchor": "e"}
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

        # Новая панель статистики по категориям (распределяем по горизонтали)
        categories_frame = ctk.CTkFrame(self.monthly_frame)
        categories_frame.grid(row=5, column=0, columnspan=2, sticky="we", padx=10, pady=10)

        # Настраиваем grid для равномерного распределения
        for i in range(5):  # 5 колонок
            categories_frame.grid_columnconfigure(i, weight=1)

        # Категории для статистики
        self.categories = [
            "Аренда", "Дилерство", "Наличные", "Безнал", "КЦ",
            "ЗП окладники", "ЗП проценты", "Реклама", "Вед.рекламы", "Комиссия брок"
        ]

        self.category_labels = {}
        for i, category in enumerate(self.categories):
            frame = ctk.CTkFrame(categories_frame, height=60)
            frame.grid(row=i // 5, column=i % 5, padx=5, pady=5, sticky="nsew")
            frame.grid_propagate(False)

            ctk.CTkLabel(frame, text=category, font=self.large_font,
                         wraplength=100).pack(pady=(5, 0))  # Перенос текста
            self.category_labels[category] = ctk.CTkLabel(frame, text="0.00 ₽",
                                                          font=self.large_font)
            self.category_labels[category].pack()

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

        # Словарь для статистики по категориям
        category_stats = {category: 0 for category in self.categories}

        # Собираем ВСЕ операции для отображения
        for transaction in self.transactions:
            try:
                # Парсим дату
                date_str = transaction["date"].split()[0]
                day, month, year = map(int, date_str.split('.'))

                if year == selected_year and month == month_number:
                    if date_str not in daily_data:
                        daily_data[date_str] = {
                            'all_income': 0,  # Все приходы
                            'all_expense': 0,  # Все расходы
                            'transactions': []  # Все транзакции
                        }

                    amount = abs(transaction["amount"])

                    # Всегда добавляем транзакцию в список для детализации
                    daily_data[date_str]['transactions'].append(transaction)

                    # Учитываем ВСЕ операции для отображения
                    if transaction["type"] == "Приход":
                        daily_data[date_str]['all_income'] += amount
                    else:
                        daily_data[date_str]['all_expense'] += amount

                    # Собираем статистику по всем категориям (и приход и расход)
                    category = transaction["category"]
                    if category in category_stats:
                        if transaction["type"] == "Расход":
                            category_stats[category] += amount
                        else:
                            category_stats[category] -= amount  # Приход уменьшает расход по категории

            except (ValueError, IndexError):
                continue

        # Заполняем таблицу ежедневной сводки (показываем ВСЕ операции)
        for date_str in sorted(daily_data.keys(), reverse=True):
            data = daily_data[date_str]
            # Для отображения в таблице используем все операции
            balance = data['all_income'] - data['all_expense']  # Баланс по всем операциям
            transactions_count = len(data['transactions'])  # Все операции включая исключенные

            self.daily_tree.insert(
                "",
                "end",
                values=(
                    date_str,
                    f"{data['all_income']:,.2f}",  # Все приходы
                    f"{data['all_expense']:,.2f}",  # Все расходы
                    f"{balance:,.2f}",  # Баланс по всем операциям
                    transactions_count
                )
            )

        # Обновляем статистику по категориям
        for category, amount in category_stats.items():
            self.category_labels[category].configure(text=f"{abs(amount):,.2f} ₽")

    def get_monthly_report_data(self) -> Dict:
        """Собирает данные для месячного отчета"""
        try:
            selected_year = int(self.year_combo.get())
            selected_month = self.month_combo.get()
            month_number = self.get_month_number(selected_month)
        except (ValueError, AttributeError):
            return {}

        daily_data = {}
        category_stats = {category: 0 for category in self.categories}
        all_daily_details = []

        total_income = 0
        total_expense = 0

        # Собираем данные
        for transaction in self.transactions:
            try:
                date_str = transaction["date"].split()[0]
                day, month, year = map(int, date_str.split('.'))

                if year == selected_year and month == month_number:
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

                    # Добавляем в детализацию
                    all_daily_details.append({
                        'Дата': transaction["date"],
                        'День': date_str,
                        'Тип': transaction["type"],
                        'Описание': transaction["description"],
                        'Категория': transaction["category"],
                        'Тип_оплаты': transaction.get("payment_type", "Наличные"),
                        'Сумма': abs(transaction['amount']),
                        'Сумма_руб': f"{abs(transaction['amount']):,.2f} ₽"
                    })

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
                'Количество_операций': transactions_count,
                'Приход_руб': f"{data['all_income']:,.2f} ₽",
                'Расход_руб': f"{data['all_expense']:,.2f} ₽",
                'Баланс_руб': f"{balance:,.2f} ₽"
            })

        # Информация о месяце
        month_info = {
            'Год': selected_year,
            'Месяц': selected_month,
            'Всего_дней_с_операциями': len(daily_summary),
            'Общий_приход': total_income,
            'Общий_расход': total_expense,
            'Итоговый_баланс': total_income - total_expense,
            'Общий_приход_руб': f"{total_income:,.2f} ₽",
            'Общий_расход_руб': f"{total_expense:,.2f} ₽",
            'Итоговый_баланс_руб': f"{total_income - total_expense:,.2f} ₽"
        }

        # Преобразуем статистику по категориям в удобный формат
        formatted_category_stats = {}
        for category, amount in category_stats.items():
            formatted_category_stats[category] = {
                'Сумма': abs(amount),
                'Сумма_руб': f"{abs(amount):,.2f} ₽",
                'Тип': 'Расход' if amount > 0 else 'Приход'
            }

        return {
            'daily_summary': daily_summary,
            'daily_details': all_daily_details,
            'category_stats': formatted_category_stats,
            'month_info': month_info
        }

    def export_monthly_report(self):
        """Экспорт только месячного отчета"""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить месячный отчет как"
        )
        if not path:
            return

        try:
            monthly_data = self.get_monthly_report_data()

            if monthly_data:
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    # Ежедневная сводка
                    if 'daily_summary' in monthly_data and monthly_data['daily_summary']:
                        pd.DataFrame(monthly_data['daily_summary']).to_excel(
                            writer,
                            sheet_name="Ежедневная_сводка",
                            index=False
                        )

                    # Детализация операций
                    if 'daily_details' in monthly_data and monthly_data['daily_details']:
                        pd.DataFrame(monthly_data['daily_details']).to_excel(
                            writer,
                            sheet_name="Детализация_операций",
                            index=False
                        )

                    # Статистика по категориям
                    if 'category_stats' in monthly_data and monthly_data['category_stats']:
                        stats_df = pd.DataFrame(list(monthly_data['category_stats'].items()),
                                                columns=['Категория', 'Сумма'])
                        stats_df.to_excel(
                            writer,
                            sheet_name="Статистика_по_категориям",
                            index=False
                        )

                    # Общая информация
                    if 'month_info' in monthly_data:
                        pd.DataFrame([monthly_data['month_info']]).to_excel(
                            writer,
                            sheet_name="Общая_информация",
                            index=False
                        )

                messagebox.showinfo("Успех", "Месячный отчет успешно экспортирован в Excel")
            else:
                messagebox.showwarning("Предупреждение", "Нет данных для экспорта за выбранный месяц")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте месячного отчета: {str(e)}")

    def on_day_selected(self, event):
        """Обработчик выбора дня в таблице - показывает ВСЕ операции выбранного дня"""
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

        # Заполняем детальную таблицу ВСЕМИ операциями за выбранный день
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

    def refresh_data(self):
        """Обновляет все данные из базы и перерисовывает отчеты"""
        try:
            # Загружаем свежие данные из базы
            self.transactions = self.db.get_all_transactions()
            self.car_deals = self.db.get_all_car_deals()
            self.initial_capital = self.db.get_initial_capital()

            # Обновляем все отчеты
            self.update_report()
            if hasattr(self, 'daily_tree'):
                self.update_monthly_report()

            # Обновляем поле капитала в настройках
            if hasattr(self, 'capital_entry'):
                self.capital_entry.delete(0, tk.END)
                self.capital_entry.insert(0, str(self.initial_capital))

        except Exception as e:
            print(f"Ошибка при обновлении данных: {e}")

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

    def setup_report_frame(self):
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_rowconfigure(1, weight=1)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=self.large_font)
        style.configure("Treeview", font=self.large_font, rowheight=35)

        # Обновляем колонки - добавляем "Тип оплаты"
        columns = {
            "#1": {"name": "date", "text": "Дата", "width": 180, "anchor": "center"},
            "#2": {"name": "type", "text": "Тип", "width": 120, "anchor": "center"},
            "#3": {"name": "amount", "text": "Сумма", "width": 150, "anchor": "e"},
            "#4": {"name": "description", "text": "Описание", "width": 300, "anchor": "center"},
            "#5": {"name": "category", "text": "Категория", "width": 150, "anchor": "center"},
            "#6": {"name": "payment_type", "text": "Тип оплаты", "width": 120, "anchor": "center"}  # Новая колонка
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
            "#6": {"name": "expenses", "text": "Расходы", "width": 120, "anchor": "e"},
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

        key_order = ["date", "type", "amount", "description", "category", "payment_type"]
        car_key_order = ["brand", "year", "vin", "price", "cost", "expenses", "header", "comment"]

        # Исправляем привязку событий
        self.tree.bind("<Double-1>",
                       lambda event: self.on_tree_double_click(event, self.tree, self.transactions, key_order))
        self.car_tree.bind("<Double-1>",
                           lambda event: self.on_tree_double_click(event, self.car_tree, self.car_deals, car_key_order))

    def setup_settings_frame(self):
        ctk.CTkLabel(self.settings_frame, text="Стартовый капитал:", font=self.large_font).pack(pady=(20, 5))
        self.capital_entry = ctk.CTkEntry(self.settings_frame)
        self.capital_entry.insert(0, str(self.initial_capital))
        self.capital_entry.pack()

        def save_capital():
            try:
                self.initial_capital = float(self.capital_entry.get())
                self.db.update_initial_capital(self.initial_capital)
                self.show_toast("💾 Капитал обновлен")
                self.update_report()
            except ValueError:
                messagebox.showerror("Ошибка", "Введите число.")

        ctk.CTkButton(self.settings_frame, text="💾 Сохранить капитал", command=save_capital).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📥 Импорт из Excel", command=self.import_from_excel).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📤 Экспорт в Excel", command=self.export_to_excel).pack(pady=10)

    def on_tree_double_click(self, event, tree, data_list, key_order):
        """Обработчик двойного клика по дереву"""
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)

        if not item or not column:
            return

        col_index = int(column[1:]) - 1
        self.edit_cell(tree, item, col_index, data_list, key_order)

    def get_data_list_by_iid(self, iid):
        if iid.startswith("tr_"):
            return self.transactions
        elif iid.startswith("car_"):
            return self.car_deals
        else:
            return None

    def edit_cell(self, tree, item, col_index, data_list, key_order):
        try:
            x, y, width, height = tree.bbox(item, f"#{col_index + 1}")
            if not all([x, y, width, height]):
                return

            value = tree.item(item, "values")[col_index]

            entry = ctk.CTkEntry(tree, width=width, height=height)
            entry.insert(0, str(value))
            entry.place(x=x, y=y)

            def save_edit(event=None):
                try:
                    new_val = entry.get()
                    values = list(tree.item(item, "values"))
                    values[col_index] = new_val
                    tree.item(item, values=values)
                    entry.destroy()

                    # Получаем iid элемента
                    iid = tree.item(item, "text")

                    if not iid:
                        return

                    # Определяем тип данных (транзакции или авто-сделки)
                    if iid.startswith("tr_"):
                        index = int(iid.split("_")[1])
                        data_type = "transaction"
                    elif iid.startswith("car_"):
                        index = int(iid.split("_")[1])
                        data_type = "car_deal"
                    else:
                        return

                    if index < len(data_list) and col_index < len(key_order):
                        key = key_order[col_index]
                        cleaned = new_val.replace(",", "").replace(" ", "").strip()

                        try:
                            # Преобразуем числовые поля
                            if key in ["amount", "price", "cost", "header", "expenses"]:
                                cleaned = float(cleaned)
                        except ValueError:
                            # Если не число, оставляем как строку
                            pass

                        # Сохраняем изменения в базе данных
                        if data_type == "transaction":
                            transaction_id = data_list[index]["id"]
                            updates = {key: cleaned}

                            # Особенная обработка для суммы
                            if key == "amount":
                                current_type = data_list[index]["type"]
                                if current_type == "Расход":
                                    updates["amount"] = -abs(float(cleaned))
                                else:
                                    updates["amount"] = abs(float(cleaned))

                            # Обновляем в базе данных
                            success = self.db.update_transaction(transaction_id, updates)
                            if success:
                                print(f"Транзакция {transaction_id} обновлена: {updates}")

                        elif data_type == "car_deal":
                            deal_id = data_list[index]["id"]
                            updates = {}

                            # Особенная обработка для авто-сделок
                            if key == "price":
                                price = float(cleaned)
                                cost = float(data_list[index].get("cost", 0))
                                expenses = float(data_list[index].get("expenses", 0))
                                header = price - cost - expenses
                                updates = {"price": price, "header": header}

                            elif key == "cost":
                                cost = float(cleaned)
                                price = float(data_list[index].get("price", 0))
                                expenses = float(data_list[index].get("expenses", 0))
                                header = price - cost - expenses
                                updates = {"cost": cost, "header": header}

                            elif key == "expenses":
                                expenses = float(cleaned)
                                price = float(data_list[index].get("price", 0))
                                cost = float(data_list[index].get("cost", 0))
                                header = price - cost - expenses
                                updates = {"expenses": expenses, "header": header}

                            else:
                                updates = {key: cleaned}

                            # Обновляем в базе данных
                            success = self.db.update_car_deal(deal_id, updates)
                            if success:
                                print(f"Авто-сделка {deal_id} обновлена: {updates}")

                    # Полностью перезагружаем все данные из базы
                    self.refresh_data()

                except Exception as e:
                    print(f"Ошибка при сохранении редактирования: {e}")
                    entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)
            entry.focus_set()

        except Exception as e:
            print(f"Ошибка при редактировании ячейки: {e}")

    def _handle_treeview_double_click(self, event, tree, edit_callback):
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item or not column:
            return

        # Получаем правильный iid
        iid = tree.item(item, "text") if hasattr(tree.item(item), "text") else item

        col_index = int(column[1:]) - 1

        # Определяем тип данных для редактирования
        if iid.startswith("tr_"):
            edit_callback(tree, item, col_index, self.transactions,
                          ["date", "type", "amount", "description", "category", "payment_type"])
        elif iid.startswith("car_"):
            edit_callback(tree, item, col_index, self.car_deals,
                          ["brand", "year", "vin", "price", "cost", "expenses", "header", "comment"])

    def add_transaction(self):
        try:
            operation = self.entries["Тип операции:"].get()
            amount = float(self.entries["Сумма:"].get())
            description = self.entries["Описание:"].get().strip()
            payment_type = self.entries["Тип оплаты:"].get()
            category = self.entries["Категория:"].get()

            # Используем значение чекбокса
            exclude_from_total = self.exclude_from_total.get()

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
                "category": category,
                "payment_type": payment_type,
                "exclude_from_total": exclude_from_total
            }

            self.db.add_transaction(transaction)

            # Обновляем данные и все отчеты
            self.refresh_data()

            # Сбрасываем форму
            self.entries["Сумма:"].delete(0, tk.END)
            self.entries["Сумма:"].insert(0, "0.00")
            self.entries["Описание:"].delete(0, tk.END)
            self.exclude_from_total.set(False)  # Сбрасываем чекбокс

            if exclude_from_total:
                self.show_toast("✅ Операция добавлена (исключена из расхода)")
            else:
                self.show_toast("✅ Операция добавлена")

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

            self.show_toast("🚗 Авто-сделка добавлена")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении: {str(e)}")

    def update_report(self):
        # Очищаем таблицы
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.car_tree.get_children():
            self.car_tree.delete(item)

        # Заполняем таблицу транзакций
        for i, tr in enumerate(self.transactions):
            item_id = f"tr_{i}"
            self.tree.insert(
                "",
                "end",
                iid=item_id,
                text=item_id,  # Важно: устанавливаем text такой же как iid
                values=(
                    tr["date"],
                    tr["type"],
                    f"{abs(tr['amount']):,.2f}",
                    tr["description"],
                    tr["category"],
                    tr.get("payment_type", "Наличные")
                )
            )

        # Заполняем таблицу авто-сделок
        for i, deal in enumerate(self.car_deals):
            item_id = f"car_{i}"
            self.car_tree.insert(
                "", "end",
                iid=item_id,
                text=item_id,  # Важно: устанавливаем text такой же как iid
                values=(
                    deal.get("brand", ""),
                    deal.get("year", ""),
                    deal.get("vin", ""),
                    f"{deal.get('price', 0):,.2f}",
                    f"{deal.get('cost', 0):,.2f}",
                    f"{deal.get('expenses', 0):,.2f}",
                    f"{deal.get('header', 0):,.2f}",
                    deal.get("comment", "")
                )
            )

        self.update_summary()

    def update_summary(self):
        # Исключаемые категории (не учитываются в общем приходе/расходе)
        excluded_categories = ["ЗП окладники", "ЗП проценты", "Комиссия брок"]

        # Фильтруем транзакции, исключая указанные категории И операции с флагом exclude_from_total
        filtered_transactions = [
            t for t in self.transactions
            if t["category"] not in excluded_categories and not t.get("exclude_from_total", False)
        ]

        # ВАЖНО: используем abs() для суммы, но сохраняем знак для расчета
        total_income = sum(t["amount"] for t in filtered_transactions if t["type"] == "Приход")
        total_expense = abs(sum(t["amount"] for t in filtered_transactions if t["type"] == "Расход"))

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
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            imported_count = {
                'transactions': 0,
                'car_deals': 0
            }

            with pd.ExcelFile(path) as xls:
                sheet_names = xls.sheet_names

                # Импорт транзакций
                transaction_sheets = ['Транзакции', 'Transactions']
                for sheet_name in transaction_sheets:
                    if sheet_name in sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        print(f"Найдено {len(df)} транзакций в листе {sheet_name}")

                        for index, row in df.iterrows():
                            try:
                                # Отладочная информация
                                print(f"Обрабатываем строку {index}: {dict(row)}")

                                # Обрабатываем разные форматы названий колонок
                                date_value = row.get("date", row.get("Дата", ""))
                                if pd.isna(date_value):
                                    date_value = datetime.now().strftime("%d.%m.%Y %H:%M")
                                else:
                                    # Конвертируем дату из Excel в правильный формат
                                    if isinstance(date_value, datetime):
                                        date_value = date_value.strftime("%d.%m.%Y %H:%M")
                                    elif isinstance(date_value, str):
                                        # Если это строка, оставляем как есть
                                        pass
                                    else:
                                        # Для других типов (например, timestamp)
                                        date_value = datetime.now().strftime("%d.%m.%Y %H:%M")

                                trans_type = row.get("type", row.get("Тип", "Приход"))
                                if pd.isna(trans_type):
                                    trans_type = "Приход"

                                amount = float(row.get("amount", row.get("Сумма", 0)))
                                if pd.isna(amount):
                                    amount = 0

                                description = str(row.get("description", row.get("Описание", "")))
                                if pd.isna(description):
                                    description = ""

                                category = row.get("category", row.get("Категория", "Другое"))
                                if pd.isna(category):
                                    category = "Другое"

                                payment_type = row.get("payment_type", row.get("Тип_оплаты", "Наличные"))
                                if pd.isna(payment_type):
                                    payment_type = "Наличные"

                                # Корректируем amount в зависимости от типа
                                if trans_type == "Расход":
                                    amount = -abs(amount)
                                else:
                                    amount = abs(amount)

                                transaction = {
                                    "date": str(date_value),
                                    "type": str(trans_type),
                                    "amount": amount,
                                    "description": str(description).strip(),
                                    "category": str(category),
                                    "payment_type": str(payment_type)
                                }

                                print(f"Создана транзакция: {transaction}")

                                if not self.db.exists_transaction(transaction):
                                    self.db.add_transaction(transaction)
                                    imported_count['transactions'] += 1
                                    print(f"Транзакция добавлена: {imported_count['transactions']}")
                                else:
                                    print("Транзакция уже существует, пропускаем")

                            except Exception as e:
                                print(f"Ошибка при импорте транзакции в строке {index}: {e}")
                                import traceback
                                traceback.print_exc()

                # Импорт авто-сделок
                car_sheets = ['Авто-сделки', 'CarDeals']
                for sheet_name in car_sheets:
                    if sheet_name in sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        print(f"Найдено {len(df)} авто-сделок в листе {sheet_name}")

                        for index, row in df.iterrows():
                            try:
                                brand = str(row.get("brand", row.get("Марка", ""))).strip()
                                if pd.isna(brand) or not brand:
                                    continue

                                year = str(row.get("year", row.get("Год", ""))).strip()
                                if pd.isna(year):
                                    year = ""

                                vin = str(row.get("vin", row.get("VIN", ""))).strip()
                                if pd.isna(vin):
                                    vin = ""

                                price = float(row.get("price", row.get("Цена_продажи", row.get("Цена", 0))))
                                if pd.isna(price):
                                    price = 0

                                cost = float(row.get("cost", row.get("Закупочная_стоимость", row.get("Стоимость", 0))))
                                if pd.isna(cost):
                                    cost = 0

                                expenses = float(row.get("expenses", row.get("Расходы", 0)))
                                if pd.isna(expenses):
                                    expenses = 0

                                profit = float(
                                    row.get("profit", row.get("Прибыль", row.get("header", price - cost - expenses))))
                                if pd.isna(profit):
                                    profit = price - cost - expenses

                                comment = str(row.get("comment", row.get("Комментарий", "")))
                                if pd.isna(comment):
                                    comment = ""

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

                                print(f"Создана авто-сделка: {car_deal}")

                                if not self.db.exists_car_deal(car_deal):
                                    self.db.add_car_deal(car_deal)
                                    imported_count['car_deals'] += 1
                                    print(f"Авто-сделка добавлена: {imported_count['car_deals']}")
                                else:
                                    print("Авто-сделка уже существует, пропускаем")

                            except Exception as e:
                                print(f"Ошибка при импорте авто-сделки в строке {index}: {e}")
                                import traceback
                                traceback.print_exc()

                # Импорт настроек
                settings_sheets = ['Настройки', 'Settings']
                for sheet_name in settings_sheets:
                    if sheet_name in sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        capital_columns = ['initial_capital', 'Стартовый_капитал']

                        for col in capital_columns:
                            if col in df.columns:
                                try:
                                    capital_value = df.iloc[0][col]
                                    if not pd.isna(capital_value):
                                        capital = float(capital_value)
                                        self.db.update_initial_capital(capital)
                                        self.initial_capital = capital
                                        print(f"Установлен стартовый капитал: {capital}")
                                        break
                                except Exception as e:
                                    print(f"Ошибка при импорте настроек: {e}")

            # Обновляем данные из базы
            self.transactions = self.db.get_all_transactions()
            self.car_deals = self.db.get_all_car_deals()
            self.initial_capital = self.db.get_initial_capital()

            # Обновление поля капитала
            self.capital_entry.delete(0, tk.END)
            self.capital_entry.insert(0, str(self.initial_capital))

            # Полностью обновляем отчёты
            self.update_report()
            self.update_monthly_report()

            # Формируем многострочное сообщение для тоста
            message_lines = [
                "📥 Импорт завершен",
                f"Транзакций: +{imported_count['transactions']}",
                f"Авто-сделок: +{imported_count['car_deals']}",
                f"Капитал: {self.initial_capital:,.2f}₽"
            ]
            self.show_toast("\n".join(message_lines), 4000)  # Показываем чуть дольше

        except Exception as e:
            print(f"Общая ошибка импорта: {e}")
            import traceback
            traceback.print_exc()
            self.show_toast(f"❌ Ошибка импорта: {str(e)}")

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить полный отчет"
        )
        if not path:
            return

        try:
            # Получаем данные месячного отчета
            monthly_data = self.get_monthly_report_data()

            if self.db.export_to_excel(path, monthly_data):
                sheets = ["Транзакции", "Авто-сделки", "Настройки"]
                if monthly_data:
                    sheets.extend(["Месяц_Ежедневно", "Месяц_Операции", "Месяц_Категории", "Месяц_Инфо"])

                self.show_toast("📊 Полный отчет экспортирован в Excel", 3500)
            else:
                messagebox.showerror("Ошибка", "Не удалось экспортировать данные")
        except Exception as e:
            self.show_toast(f"❌ Ошибка экспорта: {str(e)}")

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

        if not self.ask_confirmation("Подтверждение", "Вы уверены, что хотите удалить эту транзакцию?"):
            return

        index = int(item.split("_")[1])
        transaction_id = self.transactions[index]["id"]

        # Удаляем из базы данных
        cursor = self.db.conn.cursor()
        cursor.execute("DELETE FROM transactions WHERE id = ?", (transaction_id,))
        self.db.conn.commit()

        # Обновляем данные и все отчеты
        self.refresh_data()

        self.show_toast("🗑️ Транзакция удалена")

    def delete_selected_car_deal(self):
        selected_item = self.car_tree.selection()
        if not selected_item:
            return

        item = selected_item[0]
        if not item.startswith("car_"):
            return

        if not self.ask_confirmation("Подтверждение", "Вы уверены, что хотите удалить эту авто-сделку?"):
            return

        index = int(item.split("_")[1])
        deal_id = self.car_deals[index]["id"]

        # Удаляем из базы данных
        cursor = self.db.conn.cursor()
        cursor.execute("DELETE FROM car_deals WHERE id = ?", (deal_id,))
        self.db.conn.commit()

        # Обновляем данные и все отчеты
        self.refresh_data()

        self.show_toast("🚗 Авто-сделка удалена")

if __name__ == "__main__":
    root = ctk.CTk()
    app = MoneyTrackerApp(root)
    root.mainloop()