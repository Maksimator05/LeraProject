import os
import tempfile
import pytest
import pandas as pd
from datetime import datetime
from MoneyTracker import DatabaseManager  # <-- замени на свой путь

@pytest.fixture
def db():
    db = DatabaseManager(":memory:")  # временная БД в памяти
    yield db
    db.close()

# ---------- Тесты транзакций ----------
def test_add_and_get_transaction(db):
    transaction = {
        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "type": "Приход",
        "amount": 1000.0,
        "description": "Зарплата",
        "category": "Наличные"
    }
    tr_id = db.add_transaction(transaction)
    assert isinstance(tr_id, int)

    transactions = db.get_all_transactions()
    assert len(transactions) == 1
    assert transactions[0]["description"] == "Зарплата"

def test_update_transaction(db):
    tr_id = db.add_transaction({
        "date": "2025-01-01",
        "type": "Расход",
        "amount": -500,
        "description": "Еда",
        "category": "Карта"
    })
    updated = db.update_transaction(tr_id, {"amount": -600, "description": "Продукты"})
    assert updated is True
    tr = db.get_all_transactions()[0]
    assert tr["amount"] == -600
    assert tr["description"] == "Продукты"

# ---------- Тесты авто-сделок ----------
def test_add_and_get_car_deal(db):
    deal = {
        "brand": "BMW",
        "year": "2020",
        "vin": "123VIN",
        "price": 20000,
        "cost": 15000,
        "header": 5000,
        "comment": "Хорошее авто"
    }
    deal_id = db.add_car_deal(deal)
    assert isinstance(deal_id, int)

    deals = db.get_all_car_deals()
    assert len(deals) == 1
    assert deals[0]["brand"] == "BMW"
    assert deals[0]["header"] == 5000

def test_update_car_deal(db):
    deal_id = db.add_car_deal({
        "brand": "Audi",
        "year": "2019",
        "vin": "VIN456",
        "price": 15000,
        "cost": 10000,
        "header": 5000,
        "comment": ""
    })
    updated = db.update_car_deal(deal_id, {"price": 18000, "header": 8000})
    assert updated is True
    deal = db.get_all_car_deals()[0]
    assert deal["price"] == 18000
    assert deal["header"] == 8000

# ---------- Тесты настроек ----------
def test_initial_capital(db):
    assert db.get_initial_capital() == 0
    db.update_initial_capital(5000)
    assert db.get_initial_capital() == 5000

# ---------- Тесты экспорта/импорта ----------
def test_export_and_import_excel(db, tmp_path):
    # добавим данные
    transaction_id = db.add_transaction({
        "date": "2025-01-01 12:00",
        "type": "Приход",
        "amount": 1000.0,
        "description": "Тестовая транзакция",
        "category": "Наличные"
    })

    # Проверим, что транзакция добавилась
    initial_transactions = db.get_all_transactions()
    assert len(initial_transactions) == 1

    # Используем СТАРЫЕ ключи, которые ожидает DatabaseManager
    deal_id = db.add_car_deal({
        "brand": "Tesla Model 3",
        "year": "2021",
        "vin": "TESLA123",
        "price": 50000.0,
        "cost": 40000.0,
        "header": 10000.0,
        "comment": "Электро"
    })

    # Проверим, что авто-сделка добавилась
    initial_deals = db.get_all_car_deals()
    assert len(initial_deals) == 1

    db.update_initial_capital(777.0)
    assert db.get_initial_capital() == 777.0

    # экспорт
    export_file = tmp_path / "export.xlsx"
    success = db.export_to_excel(str(export_file))
    assert success is True
    assert export_file.exists()

    # создаем новую бд и импортируем
    db2 = DatabaseManager(":memory:")

    # Проверим, что новая БД пустая
    assert len(db2.get_all_transactions()) == 0
    assert len(db2.get_all_car_deals()) == 0
    assert db2.get_initial_capital() == 0.0

    imported = db2.import_from_excel(str(export_file))
    assert imported is True

    # проверяем, что данные перенеслись
    trs = db2.get_all_transactions()
    deals = db2.get_all_car_deals()
    capital = db2.get_initial_capital()

    print(f"Импортированные транзакции: {trs}")
    print(f"Импортированные сделки: {deals}")
    print(f"Импортированный капитал: {capital}")

    assert len(trs) == 1
    assert trs[0]['description'] == "Тестовая транзакция"

    assert len(deals) == 1
    assert deals[0]['brand'] == "Tesla Model 3"  # Используем старый ключ 'brand'

    assert capital == 777.0

# ---------- Тест закрытия ----------
def test_close_connection(db):
    db.close()
    assert db.conn is None
