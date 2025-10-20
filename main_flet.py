#!/usr/bin/env python3
"""
Главный файл для запуска Flet версии
"""

import sys
import os

# Добавляем текущую директорию в путь для импорта
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from frontend_flet.main import
import flet as ft

def main():
    """Запуск Flet приложения"""
    app = MoneyTrackerApp()
    ft.app(target=app.main)

if __name__ == "__main__":
    main()