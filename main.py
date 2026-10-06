"""Точка входа: Агрегатор Excel-таблиц.

Запуск: python main.py  (или собранный «Агрегатор таблиц.exe»).
"""

from __future__ import annotations

import os
import sys


def _setup_console_encoding() -> None:
    """UTF-8 в консоли Windows, чтобы русский текст и рамки не ломались."""
    if sys.platform == "win32":
        try:
            os.system("chcp 65001 >nul 2>&1")
        except OSError:
            pass
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def main() -> None:
    _setup_console_encoding()
    from excel_aggregator.app import AggregatorApp, log_unexpected_error

    app = AggregatorApp()
    while True:
        try:
            app.run()
            return
        except KeyboardInterrupt:
            print("\nРабота прервана пользователем.")
            return
        except EOFError:
            print("\nВвод завершён. Работа программы остановлена.")
            return
        except Exception:  # непредвиденная ошибка: журнал + возврат в меню
            log_unexpected_error()
            print("\nПроизошла непредвиденная ошибка. "
                  "Подробности записаны в файл aggregator_error.log.")
            print("Возврат в главное меню...\n")


if __name__ == "__main__":
    main()
