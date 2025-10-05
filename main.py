#!/usr/bin/env python3
"""
Главный файл приложения Money Tracker
"""

import sys
import os

# Добавляем текущую директорию в путь для импорта
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from frontend.main import MoneyTrackerApp
import customtkinter as ctk

def main():
    """Запускает главное приложение"""
    try:
        root = ctk.CTk()
        app = MoneyTrackerApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка при запуске приложения: {e}")
        input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()