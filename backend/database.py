import sqlite3
import pandas as pd
from typing import List, Dict
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_file="money_tracker.db"):
        self.db_name = db_file
        self.conn = sqlite3.connect(self.db_name, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()

        # Создаем таблицу транзакций (с полем exclude_from_total)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                type TEXT NOT NULL,
                amount REAL NOT NULL,
                description TEXT NOT NULL,
                category TEXT NOT NULL,
                payment_type TEXT NOT NULL DEFAULT 'Наличные',
                exclude_from_total INTEGER DEFAULT 0
            )
        """)

        # Проверяем наличие столбца exclude_from_total и добавляем его, если нужно
        try:
            cursor.execute("PRAGMA table_info(transactions)")
            columns = [column[1] for column in cursor.fetchall()]
            if 'exclude_from_total' not in columns:
                cursor.execute("ALTER TABLE transactions ADD COLUMN exclude_from_total INTEGER DEFAULT 0")
        except sqlite3.Error as e:
            print(f"Ошибка при проверке столбцов: {e}")

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

        # Проверяем наличие столбца expenses и добавляем его, если нужно
        try:
            cursor.execute("SELECT expenses FROM car_deals LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE car_deals ADD COLUMN expenses REAL DEFAULT 0")

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
    def add_transaction(self, transaction):
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO transactions (date, type, amount, description, category, payment_type, exclude_from_total)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            transaction['date'],
            transaction['type'],
            transaction['amount'],
            transaction['description'],
            transaction['category'],
            transaction.get('payment_type', 'Наличные'),
            int(transaction.get('exclude_from_total', False))
        ))
        self.conn.commit()

    def get_all_transactions(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT id, date, type, amount, description, category, payment_type, exclude_from_total 
            FROM transactions ORDER BY date DESC
        ''')
        transactions = []
        for row in cursor.fetchall():
            transactions.append({
                'id': row[0],
                'date': row[1],
                'type': row[2],
                'amount': row[3],
                'description': row[4],
                'category': row[5],
                'payment_type': row[6],
                'exclude_from_total': bool(row[7])
            })
        return transactions

    def update_transaction(self, transaction_id: int, updates: Dict) -> bool:
        if not updates:
            return False
        try:
            cursor = self.conn.cursor()
            set_clause = ", ".join(f"{key} = ?" for key in updates.keys())
            values = list(updates.values()) + [transaction_id]
            cursor.execute(f"UPDATE transactions SET {set_clause} WHERE id = ?", values)
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка при обновлении транзакции: {e}")
            return False

    def delete_transaction(self, transaction_id: int) -> bool:
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM transactions WHERE id = ?", (transaction_id,))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка при удалении транзакции: {e}")
            return False

    def exists_transaction(self, transaction: Dict) -> bool:
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) FROM transactions
            WHERE type = ? AND amount = ? AND description = ? 
            AND category = ? AND payment_type = ?
            AND date LIKE ?
        """, (
            transaction["type"],
            transaction["amount"],
            transaction["description"],
            transaction["category"],
            transaction.get("payment_type", "Наличные"),
            f"%{transaction['date'].split()[0]}%"
        ))
        return cursor.fetchone()[0] > 0

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
            car_deal.get("expenses", 0),
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
        try:
            cursor = self.conn.cursor()
            set_clause = ", ".join(f"{key} = ?" for key in updates.keys())
            values = list(updates.values()) + [deal_id]
            cursor.execute(f"UPDATE car_deals SET {set_clause} WHERE id = ?", values)
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка при обновлении авто-сделки: {e}")
            return False

    def delete_car_deal(self, deal_id: int) -> bool:
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM car_deals WHERE id = ?", (deal_id,))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка при удалении авто-сделки: {e}")
            return False

    def exists_car_deal(self, car_deal: Dict) -> bool:
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) FROM car_deals
            WHERE brand = ? AND year = ? AND vin = ?
            AND ABS(price - ?) < 0.01 AND ABS(cost - ?) < 0.01
        """, (
            car_deal["brand"],
            car_deal["year"],
            car_deal["vin"],
            car_deal.get("price", 0),
            car_deal.get("cost", 0)
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
    def export_to_excel(self, file_path: str, monthly_data: Dict = None) -> bool:
        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                # Экспорт транзакций
                transactions = self.get_all_transactions()
                if transactions:
                    df_transactions = pd.DataFrame(transactions)
                    df_transactions = df_transactions.rename(columns={
                        "date": "Дата", "type": "Тип", "amount": "Сумма",
                        "description": "Описание", "category": "Категория",
                        "payment_type": "Тип_оплаты", "exclude_from_total": "Исключено_из_расхода"
                    })
                    df_transactions.to_excel(writer, sheet_name="Транзакции", index=False)

                # Экспорт авто-сделок
                car_deals = self.get_all_car_deals()
                if car_deals:
                    df_car_deals = pd.DataFrame(car_deals)
                    df_car_deals = df_car_deals.rename(columns={
                        "brand": "Марка", "year": "Год", "vin": "VIN",
                        "price": "Цена_продажи", "cost": "Закупочная_стоимость",
                        "expenses": "Расходы", "header": "Прибыль", "comment": "Комментарий"
                    })
                    df_car_deals.to_excel(writer, sheet_name="Авто-сделки", index=False)

                # Экспорт настроек
                settings_data = {"Стартовый_капитал": [self.get_initial_capital()]}
                pd.DataFrame(settings_data).to_excel(writer, sheet_name="Настройки", index=False)

                # Экспорт месячного отчета
                if monthly_data:
                    if 'daily_summary' in monthly_data and monthly_data['daily_summary']:
                        pd.DataFrame(monthly_data['daily_summary']).to_excel(
                            writer, sheet_name="Месяц_Ежедневно", index=False)
                    if 'daily_details' in monthly_data and monthly_data['daily_details']:
                        pd.DataFrame(monthly_data['daily_details']).to_excel(
                            writer, sheet_name="Месяц_Операции", index=False)
                    if 'category_stats' in monthly_data and monthly_data['category_stats']:
                        stats_df = pd.DataFrame(list(monthly_data['category_stats'].items()),
                                                columns=['Категория', 'Сумма'])
                        stats_df.to_excel(writer, sheet_name="Месяц_Категории", index=False)
                    if 'month_info' in monthly_data:
                        pd.DataFrame([monthly_data['month_info']]).to_excel(
                            writer, sheet_name="Месяц_Инфо", index=False)

            return True
        except Exception as e:
            print(f"Ошибка при экспорте: {e}")
            return False

    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None