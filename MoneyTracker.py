import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import csv
from datetime import datetime
import pandas as pd

# Настройка тем - принудительно темная
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")


class MoneyTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 Авто-Трекер Финансов v2.2")
        self.root.geometry("1300x900")

        self.large_font = ("Arial", 14)
        self.xlarge_font = ("Arial", 16, "bold")
        self.xxlarge_font = ("Arial", 18, "bold")

        self.transactions = []
        self.car_deals = []
        self.initial_capital = 0
        self.load_data()

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

        self.settings_frame = ctk.CTkFrame(self.notebook)
        self.notebook.add(self.settings_frame, text="⚙️ Настройки")
        self.setup_settings_frame()

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

        # 👇 Добавляем кнопку после всех полей
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
            ("Модель авто:", "entry", None, ""),
            ("Дата покупки:", "entry", None, datetime.now().strftime("%d.%m.%Y")),
            ("Цена покупки:", "entry", None, "0.00"),
            ("Тип оплаты покупки:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные"),
            ("ФИО продавца:", "entry", None, ""),  # Новое необязательное поле
            ("Дата продажи:", "entry", None, datetime.now().strftime("%d.%m.%Y")),
            ("Цена продажи:", "entry", None, "0.00"),
            ("Тип оплаты продажи:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные"),
            ("ФИО покупателя:", "entry", None, ""),  # Новое необязательное поле
            ("На комиссии:", "combobox", ["Да", "Нет"], "Нет"),  # Новое поле
            ("Доп. расходы:", "entry", None, "0.00"),
            ("Тип оплаты расходов:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные"),
            ("Описание расходов:", "entry", None, "")
        ]

        self.car_entries = {}
        for row, (label, field_type, options, default) in enumerate(car_fields, start=1):
            ctk.CTkLabel(self.car_frame, text=label).grid(
                row=row, column=0, padx=10, pady=5, sticky="e")

            if field_type == "combobox":
                entry = ctk.CTkComboBox(self.car_frame, values=options)
                entry.set(default)
            else:
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

    def setup_settings_frame(self):
        ctk.CTkLabel(self.settings_frame, text="Стартовый капитал:", font=self.large_font).pack(pady=(20, 5))
        self.capital_entry = ctk.CTkEntry(self.settings_frame)
        self.capital_entry.insert(0, str(self.initial_capital))
        self.capital_entry.pack()

        def save_capital():
            try:
                self.initial_capital = float(self.capital_entry.get())
                self.save_data()
                messagebox.showinfo("Сохранено", "Стартовый капитал обновлён.")
            except ValueError:
                messagebox.showerror("Ошибка", "Введите число.")

        ctk.CTkButton(self.settings_frame, text="💾 Сохранить капитал", command=save_capital).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📥 Импорт из Excel", command=self.import_from_excel).pack(pady=10)
        ctk.CTkButton(self.settings_frame, text="📤 Экспорт в Excel", command=self.export_to_excel).pack(pady=10)

    def import_from_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            df = pd.read_excel(path, sheet_name=None)
            if "Transactions" in df:
                self.transactions = df["Transactions"].to_dict("records")
            if "CarDeals" in df:
                self.car_deals = df["CarDeals"].to_dict("records")
            if "Config" in df and "initial_capital" in df["Config"].columns:
                self.initial_capital = df["Config"].iloc[0]["initial_capital"]
                self.capital_entry.delete(0, tk.END)
                self.capital_entry.insert(0, str(self.initial_capital))
            self.save_data()
            self.update_report()
            messagebox.showinfo("Успех", "Данные успешно импортированы из Excel.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при импорте: {e}")

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                pd.DataFrame(self.transactions).to_excel(writer, sheet_name="Transactions", index=False)
                pd.DataFrame(self.car_deals).to_excel(writer, sheet_name="CarDeals", index=False)
                pd.DataFrame([{"initial_capital": self.initial_capital}]).to_excel(writer, sheet_name="Config", index=False)
            messagebox.showinfo("Успех", "Данные успешно экспортированы в Excel.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")

    def setup_report_frame(self):
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_rowconfigure(1, weight=1)

        # Стиль
        style = ttk.Style()
        style.configure("Treeview.Heading", font=self.large_font)
        style.configure("Treeview", font=self.large_font, rowheight=35)

        # Колонки операций
        columns = {
            "#1": {"name": "date", "text": "Дата", "width": 180, "anchor": "center"},
            "#2": {"name": "type", "text": "Тип", "width": 120, "anchor": "center"},
            "#3": {"name": "amount", "text": "Сумма", "width": 150, "anchor": "e"},
            "#4": {"name": "description", "text": "Описание", "width": 300},
            "#5": {"name": "category", "text": "Категория", "width": 150, "anchor": "center"}
        }

        self.tree = ttk.Treeview(self.report_frame, columns=list(columns.keys()), show="headings")
        for col, params in columns.items():
            self.tree.heading(col, text=params["text"])
            self.tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

        # Колонки авто-сделок
        car_columns = {
            "#1": {"name": "model", "text": "Модель", "width": 150, "anchor": "center"},
            "#2": {"name": "buy_date", "text": "Дата покупки", "width": 120, "anchor": "center"},
            "#3": {"name": "buy_price", "text": "Цена покупки", "width": 120, "anchor": "e"},
            "#4": {"name": "buy_type", "text": "Оплата покупки", "width": 120, "anchor": "center"},
            "#5": {"name": "seller_name", "text": "Продавец", "width": 150, "anchor": "center"},
            "#6": {"name": "sell_date", "text": "Дата продажи", "width": 120, "anchor": "center"},
            "#7": {"name": "sell_price", "text": "Цена продажи", "width": 120, "anchor": "e"},
            "#8": {"name": "sell_type", "text": "Оплата продажи", "width": 120, "anchor": "center"},
            "#9": {"name": "buyer_name", "text": "Покупатель", "width": 150, "anchor": "center"},
            "#10": {"name": "on_commission", "text": "Комиссия", "width": 100, "anchor": "center"},
            "#11": {"name": "expenses", "text": "Доп. расходы", "width": 120, "anchor": "e"},
            "#12": {"name": "expenses_type", "text": "Оплата расходов", "width": 120, "anchor": "center"},
            "#13": {"name": "profit", "text": "Прибыль", "width": 120, "anchor": "e"},
            "#14": {"name": "expenses_desc", "text": "Описание расходов", "width": 200}
        }

        self.car_tree = ttk.Treeview(self.report_frame, columns=list(car_columns.keys()), show="headings")
        for col, params in car_columns.items():
            self.car_tree.heading(col, text=params["text"])
            self.car_tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

        # Скроллбары
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

        # Панель итогов
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

        # Редактирование
        key_order = ["date", "type", "amount", "description", "category"]
        car_key_order = [
            "model", "buy_date", "buy_price", "buy_type", "seller_name",
            "sell_date", "sell_price", "sell_type", "buyer_name", "on_commission",
            "expenses", "expenses_type", "profit", "expenses_desc"
        ]

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

            # Сохраняем в data_list
            index = tree.index(item)

            # Защита от выхода за границы
            if index < len(data_list) and col_index < len(key_order):
                key = key_order[col_index]

                cleaned = new_val.replace(",", "")
                try:
                    if key in ["amount", "buy_price", "sell_price", "expenses", "profit"]:
                        cleaned = float(cleaned)
                except ValueError:
                    pass

                # Убедимся, что это словарь и нужный ключ есть
                if isinstance(data_list[index], dict):
                    data_list[index][key] = cleaned

            self.save_data()
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

            self.transactions.append(transaction)
            self.save_data()

            self.entries["Сумма:"].delete(0, tk.END)
            self.entries["Сумма:"].insert(0, "0.00")
            self.entries["Описание:"].delete(0, tk.END)

            self.update_report()
            messagebox.showinfo("Успех", "Операция успешно добавлена!")

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму!")

    def add_car_deal(self):
        try:
            model = self.car_entries["Модель авто"].get().strip()
            buy_date = self.car_entries["Дата покупки"].get().strip()
            buy_price = float(self.car_entries["Цена покупки"].get())
            buy_type = self.car_entries["Тип оплаты покупки"].get()
            seller_name = self.car_entries["ФИО продавца"].get().strip()  # Новое поле
            sell_date = self.car_entries["Дата продажи"].get().strip()
            sell_price = float(self.car_entries["Цена продажи"].get())
            sell_type = self.car_entries["Тип оплаты продажи"].get()
            buyer_name = self.car_entries["ФИО покупателя"].get().strip()  # Новое поле
            on_commission = self.car_entries["На комиссии"].get()  # Новое поле
            expenses = float(self.car_entries["Доп. расходы"].get())
            expenses_type = self.car_entries["Тип оплаты расходов"].get()
            expenses_desc = self.car_entries["Описание расходов"].get().strip()

            if not model:
                messagebox.showerror("Ошибка", "Введите модель авто!")
                return

            profit = sell_price - buy_price - expenses

            car_deal = {
                "model": model,
                "buy_date": buy_date,
                "buy_price": buy_price,
                "buy_type": buy_type,
                "seller_name": seller_name,  # Новое поле
                "sell_date": sell_date,
                "sell_price": sell_price,
                "sell_type": sell_type,
                "buyer_name": buyer_name,  # Новое поле
                "on_commission": on_commission,  # Новое поле
                "expenses": expenses,
                "expenses_type": expenses_type,
                "profit": profit,
                "expenses_desc": expenses_desc
            }

            self.car_deals.append(car_deal)
            self.save_data()

            # Очищаем поля
            for entry in self.car_entries.values():
                if isinstance(entry, ctk.CTkEntry):
                    entry.delete(0, tk.END)
                elif isinstance(entry, ctk.CTkComboBox):
                    entry.set("Наличные")  # Сбрасываем на значение по умолчанию

            self.update_report()
            messagebox.showinfo("Успех", "Авто-сделка успешно добавлена!")

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения!")

    def save_settings(self):
        try:
            self.initial_capital = float(self.initial_capital_entry.get().replace(",", ""))
            messagebox.showinfo("Успех", "Настройки успешно сохранены!")
            self.update_report()
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму начального капитала!")

    def update_report(self):
        # Очищаем таблицы
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.car_tree.get_children():
            self.car_tree.delete(item)

        # Заполняем таблицу операций
        for tr in sorted(self.transactions, key=lambda x: x["date"], reverse=True):
            self.tree.insert("", "end", values=(
                tr["date"],
                tr["type"],
                f"{abs(tr['amount']):,.2f}",
                tr["description"],
                tr["category"]
            ))

        # Заполняем таблицу авто-сделок
        for deal in self.car_deals:
            self.car_tree.insert("", "end", values=(
                deal["model"],
                deal["buy_date"],
                f"{deal['buy_price']:,.2f}",
                deal["buy_type"],
                deal.get("seller_name", ""),  # Новое поле
                deal["sell_date"],
                f"{deal['sell_price']:,.2f}",
                deal["sell_type"],
                deal.get("buyer_name", ""),  # Новое поле
                deal.get("on_commission", "Нет"),  # Новое поле
                f"{deal['expenses']:,.2f}",
                deal["expenses_type"],
                f"{deal['profit']:,.2f}",
                deal["expenses_desc"]
            ))

        # Обновляем итоги
        self.update_summary()

    def update_summary(self):
        total_income = sum(t["amount"] for t in self.transactions if t["type"] == "Приход")
        total_expense = abs(sum(t["amount"] for t in self.transactions if t["type"] == "Расход"))

        # Дополнительные вложения (расходы сверх начального капитала)
        additional_investment = max(0, total_expense - self.initial_capital)

        # Прибыль с авто
        car_profit = sum(deal["profit"] for deal in self.car_deals)

        # Общая прибыль
        total_profit = car_profit + total_income - max(0, total_expense - self.initial_capital)

        # Обновляем значения
        self.summary_labels["initial_capital"].configure(text=f"{self.initial_capital:,.2f} ₽")
        self.summary_labels["total_income"].configure(text=f"{total_income:,.2f} ₽")
        self.summary_labels["total_expense"].configure(text=f"{total_expense:,.2f} ₽")
        self.summary_labels["additional_investment"].configure(text=f"{additional_investment:,.2f} ₽")
        self.summary_labels["car_profit"].configure(text=f"{car_profit:,.2f} ₽")
        self.summary_labels["total_profit"].configure(text=f"{total_profit:,.2f} ₽")

    def load_data(self):
        try:
            # Загрузка обычных операций
            with open("transactions.csv", "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                self.transactions = [{
                    "date": row["date"],
                    "type": row["type"],
                    "amount": float(row["amount"]),
                    "description": row["description"],
                    "category": row["category"]
                } for row in reader]

            # Загрузка авто-сделок
            with open("car_deals.csv", "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                self.car_deals = [{
                    "model": row["model"],
                    "buy_date": row["buy_date"],
                    "buy_price": float(row["buy_price"]),
                    "buy_type": row.get("buy_type", "Наличные"),
                    "seller_name": row.get("seller_name", ""),  # Новое поле
                    "sell_date": row["sell_date"],
                    "sell_price": float(row["sell_price"]),
                    "sell_type": row.get("sell_type", "Наличные"),
                    "buyer_name": row.get("buyer_name", ""),  # Новое поле
                    "on_commission": row.get("on_commission", "Нет"),  # Новое поле
                    "expenses": float(row["expenses"]),
                    "expenses_type": row.get("expenses_type", "Наличные"),
                    "profit": float(row["profit"]),
                    "expenses_desc": row["expenses_desc"]
                } for row in reader]

            # Загрузка начального капитала
            with open("settings.csv", "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self.initial_capital = float(row.get("initial_capital", 0))

        except FileNotFoundError:
            self.transactions = []
            self.car_deals = []
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")
            self.transactions = []
            self.car_deals = []

    def save_data(self):
        try:
            # Сохранение обычных операций
            with open("transactions.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["date", "type", "amount", "description", "category"])
                writer.writeheader()
                writer.writerows(self.transactions)

            # Сохранение авто-сделок
            with open("car_deals.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=[
                    "model", "buy_date", "buy_price", "buy_type", "seller_name",
                    "sell_date", "sell_price", "sell_type", "buyer_name", "on_commission",
                    "expenses", "expenses_type", "profit", "expenses_desc"
                ])
                writer.writeheader()
                writer.writerows(self.car_deals)

            # Сохранение настроек
            with open("settings.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["initial_capital"])
                writer.writeheader()
                writer.writerow({"initial_capital": self.initial_capital})

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {str(e)}")


if __name__ == "__main__":
    root = ctk.CTk()
    app = MoneyTrackerApp(root)
    root.mainloop()
