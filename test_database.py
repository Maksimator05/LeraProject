import os
import pytest
from MoneyTracker import DatabaseManager  # Импортируем твой класс


@pytest.fixture
def db():
    test_db_name = "test_money_tracker.db"
    if os.path.exists(test_db_name):
        os.remove(test_db_name)
    db = DatabaseManager(test_db_name)
    yield db
    db.close()  # Закрываем соединение
    if os.path.exists(test_db_name):
        os.remove(test_db_name)


def test_add_and_get_transaction(db):
    """Проверка добавления и получения транзакции"""
    data = {
        "date": "2025-08-13",
        "type": "income",
        "amount": 1000.50,
        "description": "Test income",
        "category": "Salary"
    }
    tid = db.add_transaction(data)
    assert isinstance(tid, int)

    transactions = db.get_all_transactions()
    assert len(transactions) == 1
    assert transactions[0]["description"] == "Test income"


def test_update_transaction(db):
    """Проверка обновления транзакции"""
    tid = db.add_transaction({
        "date": "2025-08-13",
        "type": "expense",
        "amount": 500,
        "description": "Test expense",
        "category": "Food"
    })
    updated = db.update_transaction(tid, {"amount": 750})
    assert updated

    transactions = db.get_all_transactions()
    assert transactions[0]["amount"] == 750


def test_add_and_get_car_deal(db):
    """Проверка добавления и получения авто-сделки"""
    data = {
        "model": "Toyota Camry",
        "buy_date": "2025-08-01",
        "buy_price": 10000,
        "buy_type": "cash",
        "seller_name": "John Doe",
        "sell_date": "2025-08-10",
        "sell_price": 12000,
        "sell_type": "transfer",
        "buyer_name": "Jane Smith",
        "on_commission": "No",
        "expenses": 500,
        "expenses_type": "repair",
        "expenses_desc": "Engine repair",
        "profit": 1500
    }
    did = db.add_car_deal(data)
    assert isinstance(did, int)

    deals = db.get_all_car_deals()
    assert len(deals) == 1
    assert deals[0]["model"] == "Toyota Camry"


def test_update_car_deal(db):
    """Проверка обновления авто-сделки"""
    did = db.add_car_deal({
        "model": "BMW X5",
        "buy_date": "2025-08-01",
        "buy_price": 15000,
        "buy_type": "credit",
        "seller_name": "AutoDealer",
        "sell_date": "2025-08-11",
        "sell_price": 18000,
        "sell_type": "cash",
        "buyer_name": "Buyer",
        "on_commission": "Yes",
        "expenses": 800,
        "expenses_type": "service",
        "expenses_desc": "Oil change",
        "profit": 2200
    })
    updated = db.update_car_deal(did, {"profit": 3000})
    assert updated

    deals = db.get_all_car_deals()
    assert deals[0]["profit"] == 3000


def test_initial_capital(db):
    """Проверка установки и получения стартового капитала"""
    assert db.get_initial_capital() == 0
    db.update_initial_capital(5000)
    assert db.get_initial_capital() == 5000


def test_export_import_excel(db, tmp_path):
    """Проверка экспорта и импорта в Excel"""

    # --- Добавляем данные в основную БД ---
    db.add_transaction({
        "date": "2025-08-13",
        "type": "income",
        "amount": 1000,
        "description": "Test export",
        "category": "Work"
    })
    db.add_car_deal({
        "model": "Audi A4",
        "buy_date": "2025-08-01",
        "buy_price": 20000,
        "buy_type": "cash",
        "seller_name": "Seller",
        "sell_date": "2025-08-12",
        "sell_price": 23000,
        "sell_type": "cash",
        "buyer_name": "Buyer",
        "on_commission": "No",
        "expenses": 1000,
        "expenses_type": "repair",
        "expenses_desc": "Paint",
        "profit": 2000
    })
    db.update_initial_capital(7000)

    # --- Экспорт в Excel ---
    export_file = tmp_path / "export.xlsx"
    assert db.export_to_excel(export_file)

    # --- Создаём новую пустую БД для импорта ---
    import_db_path = tmp_path / "import_test.db"
    new_db = DatabaseManager(import_db_path)

    # Проверяем, что новая база изначально пустая
    assert len(new_db.get_all_transactions()) == 0
    assert len(new_db.get_all_car_deals()) == 0

    # --- Импорт из Excel ---
    assert new_db.import_from_excel(export_file)

    # Проверяем, что импорт добавил ровно одну транзакцию и одну сделку
    transactions = new_db.get_all_transactions()
    car_deals = new_db.get_all_car_deals()

    assert len(transactions) == 1
    assert len(car_deals) == 1

    # Проверка содержимого транзакции
    transaction = transactions[0]
    assert transaction["amount"] == 1000
    assert transaction["category"] == "Work"
    assert transaction["description"] == "Test export"

    # Проверка содержимого сделки
    deal = car_deals[0]
    assert deal["model"] == "Audi A4"
    assert deal["profit"] == 2000


