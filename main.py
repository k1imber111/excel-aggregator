"""Точка входа: Агрегатор Excel-таблиц.

Запуск: python main.py  (или собранный «Агрегатор таблиц.exe»).
"""

from __future__ import annotations

import os
import sys
import traceback


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


def _error_log_path() -> str:
    """Путь к файлу журнала ошибок — рядом с программой."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "aggregator_error.log")


def main() -> None:
    _setup_console_encoding()
    from excel_aggregator.app import AggregatorApp

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
            try:
                with open(_error_log_path(), "a", encoding="utf-8") as fh:
                    traceback.print_exc(file=fh)
            except OSError:
                pass
            print("\nПроизошла непредвиденная ошибка. "
                  "Подробности записаны в файл aggregator_error.log.")
            print("Возврат в главное меню...\n")


if __name__ == "__main__":
    main()
