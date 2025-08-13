import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from openpyxl import Workbook, load_workbook

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
                model TEXT NOT NULL,
                buy_date TEXT NOT NULL,
                buy_price REAL NOT NULL,
                buy_type TEXT NOT NULL,
                seller_name TEXT,
                sell_date TEXT NOT NULL,
                sell_price REAL NOT NULL,
                sell_type TEXT NOT NULL,
                buyer_name TEXT,
                on_commission TEXT NOT NULL,
                expenses REAL NOT NULL,
                expenses_type TEXT NOT NULL,
                expenses_desc TEXT,
                profit REAL NOT NULL
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
                model, buy_date, buy_price, buy_type, seller_name,
                sell_date, sell_price, sell_type, buyer_name, on_commission,
                expenses, expenses_type, expenses_desc, profit
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            car_deal["model"],
            car_deal["buy_date"],
            car_deal["buy_price"],
            car_deal["buy_type"],
            car_deal.get("seller_name", ""),
            car_deal["sell_date"],
            car_deal["sell_price"],
            car_deal["sell_type"],
            car_deal.get("buyer_name", ""),
            car_deal["on_commission"],
            car_deal["expenses"],
            car_deal["expenses_type"],
            car_deal.get("expenses_desc", ""),
            car_deal["profit"]
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_all_car_deals(self) -> List[Dict]:
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM car_deals ORDER BY buy_date DESC")
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
                transactions = self.get_all_transactions()
                if transactions:
                    pd.DataFrame(transactions).to_excel(writer, sheet_name="Transactions", index=False)

                car_deals = self.get_all_car_deals()
                if car_deals:
                    pd.DataFrame(car_deals).to_excel(writer, sheet_name="CarDeals", index=False)

                pd.DataFrame([{"initial_capital": self.get_initial_capital()}]).to_excel(
                    writer, sheet_name="Config", index=False
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
        self.root.title("💰 Авто-Трекер Финансов v2.2")
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
            ("ФИО продавца:", "entry", None, ""),
            ("Дата продажи:", "entry", None, datetime.now().strftime("%d.%m.%Y")),
            ("Цена продажи:", "entry", None, "0.00"),
            ("Тип оплаты продажи:", "combobox", ["Наличные", "Безнал", "Другое"], "Наличные"),
            ("ФИО покупателя:", "entry", None, ""),
            ("На комиссии:", "combobox", ["Да", "Нет"], "Нет"),
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
            "#4": {"name": "description", "text": "Описание", "width": 300},
            "#5": {"name": "category", "text": "Категория", "width": 150, "anchor": "center"}
        }

        self.tree = ttk.Treeview(self.report_frame, columns=list(columns.keys()), show="headings")
        for col, params in columns.items():
            self.tree.heading(col, text=params["text"])
            self.tree.column(col, width=params["width"], anchor=params.get("anchor", "w"))

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
                    if key in ["amount", "buy_price", "sell_price", "expenses", "profit"]:
                        cleaned = float(cleaned)
                except ValueError:
                    pass

                data_list[index][key] = cleaned

            if iid.startswith("tr_"):
                self.db.update_transaction(data_list[index]["id"], {key: cleaned})
            elif iid.startswith("car_"):
                self.db.update_car_deal(data_list[index]["id"], {key: cleaned})

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
            self.transactions = self.db.get_all_transactions()

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
            seller_name = self.car_entries["ФИО продавца"].get().strip()
            sell_date = self.car_entries["Дата продажи"].get().strip()
            sell_price = float(self.car_entries["Цена продажи"].get())
            sell_type = self.car_entries["Тип оплаты продажи"].get()
            buyer_name = self.car_entries["ФИО покупателя"].get().strip()
            on_commission = self.car_entries["На комиссии"].get()
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
                "seller_name": seller_name,
                "sell_date": sell_date,
                "sell_price": sell_price,
                "sell_type": sell_type,
                "buyer_name": buyer_name,
                "on_commission": on_commission,
                "expenses": expenses,
                "expenses_type": expenses_type,
                "expenses_desc": expenses_desc,
                "profit": profit
            }

            self.db.add_car_deal(car_deal)
            self.car_deals = self.db.get_all_car_deals()

            for entry in self.car_entries.values():
                if isinstance(entry, ctk.CTkEntry):
                    entry.delete(0, tk.END)
                elif isinstance(entry, ctk.CTkComboBox):
                    entry.set("Наличные")

            self.update_report()
            messagebox.showinfo("Успех", "Авто-сделка успешно добавлена!")

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения!")

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
                "",
                "end",
                iid=f"car_{i}",
                values=(
                    deal["model"],
                    deal["buy_date"],
                    f"{deal['buy_price']:,.2f}",
                    deal["buy_type"],
                    deal.get("seller_name", ""),
                    deal["sell_date"],
                    f"{deal['sell_price']:,.2f}",
                    deal["sell_type"],
                    deal.get("buyer_name", ""),
                    deal.get("on_commission", "Нет"),
                    f"{deal['expenses']:,.2f}",
                    deal["expenses_type"],
                    f"{deal['profit']:,.2f}",
                    deal["expenses_desc"]
                )
            )

        self.update_summary()

    def update_summary(self):
        total_income = sum(t["amount"] for t in self.transactions if t["type"] == "Приход")
        total_expense = abs(sum(t["amount"] for t in self.transactions if t["type"] == "Расход"))

        additional_investment = max(0, total_expense - self.initial_capital)
        car_profit = sum(deal["profit"] for deal in self.car_deals)
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
            if self.db.import_from_excel(path):
                self.transactions = self.db.get_all_transactions()
                self.car_deals = self.db.get_all_car_deals()
                self.initial_capital = self.db.get_initial_capital()
                self.capital_entry.delete(0, tk.END)
                self.capital_entry.insert(0, str(self.initial_capital))
                self.update_report()
                messagebox.showinfo("Успех", "Данные успешно импортированы из Excel.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при импорте: {e}")

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            if self.db.export_to_excel(path):
                messagebox.showinfo("Успех", "Данные успешно экспортированы в Excel.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")


if __name__ == "__main__":
    root = ctk.CTk()
    app = MoneyTrackerApp(root)
    root.mainloop()