import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

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
                brand, year, vin, comment, price, cost, header
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            car_deal["brand"],
            car_deal["year"],
            car_deal["vin"],
            car_deal.get("comment", ""),
            car_deal.get("price", 0),
            car_deal.get("cost", 0),
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
        self.root.title("💰 Авто-Трекер Финансов v2.4")
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
            ("Цена:", "entry", None, "0"),
            ("Стоимость:", "entry", None, "0"),
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
            "#1": {"name": "brand", "text": "Марка", "width": 120, "anchor": "center"},
            "#2": {"name": "model_year", "text": "Год", "width": 80, "anchor": "center"},
            "#3": {"name": "vin", "text": "VIN", "width": 150, "anchor": "center"},
            "#4": {"name": "price", "text": "Цена", "width": 120, "anchor": "e"},
            "#5": {"name": "comment", "text": "Комментарий", "width": 200},
            "#6": {"name": "cost", "text": "Стоимость", "width": 120, "anchor": "e"},
            "#7": {"name": "profit", "text": "Прибыль", "width": 120, "anchor": "e"}
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
        car_key_order = ["brand", "year", "vin", "price", "comment", "cost", "header"]

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
                            deal.get("comment", ""),
                            f"{deal.get('cost', 0):,.2f}",
                            f"{deal.get('header', 0):,.2f}"
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
            brand = self.car_entries["Марка"].get().strip()
            year = self.car_entries["Год"].get().strip()
            vin = self.car_entries["VIN"].get().strip()
            price = float(self.car_entries["Цена"].get() or 0)
            cost = float(self.car_entries["Стоимость"].get() or 0)
            comment = self.car_entries["Комментарий"].get().strip()

            profit = price - cost  # Рассчитываем прибыль

            if not brand:
                messagebox.showerror("Ошибка", "Введите марку авто!")
                return

            car_deal = {
                "brand": brand,
                "year": year,
                "vin": vin,
                "price": price,
                "cost": cost,
                "header": profit,
                "comment": comment
            }

            self.db.add_car_deal(car_deal)
            self.car_deals = self.db.get_all_car_deals()
            self.update_report()
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
                    deal.get("comment", ""),
                    f"{deal.get('cost', 0):,.2f}",
                    f"{deal.get('header', 0):,.2f}"
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

        # Удаляем из списка и обновляем таблицу
        del self.transactions[index]
        self.update_report()
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

        # Удаляем из списка и обновляем таблицу
        del self.car_deals[index]
        self.update_report()
        messagebox.showinfo("Успех", "Авто-сделка успешно удалена")


if __name__ == "__main__":
    root = ctk.CTk()
    app = MoneyTrackerApp(root)
    root.mainloop()