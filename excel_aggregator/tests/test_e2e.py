"""Сквозной e2e-тест: весь сценарий приложения через subprocess.

Гоняет main.py с заранее записанным stdin на реальных примерах из
«Примеры таблиц/»: первый файл с ручными координатами, два следующих —
через автопоиск, сборка итогового .xlsx и проверка его содержимого.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SAMPLES = PROJECT_ROOT / "Примеры таблиц"
FILE1 = SAMPLES / "Каталог Сочи 2024 МОС ФАП-262.xls"
FILE2 = SAMPLES / "Каталог Сочи 2024 ПЗ-90.11 00.00.xls"
FILE3 = SAMPLES / "Превышающие препятствия 2024.xls"

pytestmark = pytest.mark.skipif(
    not all(f.is_file() for f in (FILE1, FILE2, FILE3)),
    reason="Примеры таблиц недоступны",
)


def _run_app(stdin_script: str, cwd: Path) -> subprocess.CompletedProcess:
    env = dict(os.environ, AGGREGATOR_SCRIPTED="1", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(PROJECT_ROOT / "main.py")],
        input=stdin_script,
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=cwd,
        env=env,
        timeout=300,
    )


def test_full_scenario_end_to_end(tmp_path: Path) -> None:
    out = tmp_path / "итог.xlsx"
    script = "\n".join([
        "2",                                # главное меню -> Запуск программы
        f'"{FILE1}"',                       # путь с кавычками — должны быть сняты
        "15",                               # лист «Каталог 2024»
        "A6",                               # якорь ключей
        "1",                                # вертикально
        "2-3",                              # характеристики: столбцы C и D (A — ключ)
        # названия столбцов программа читает из заголовков сама
        "1",                                # добавить ещё файл
        str(FILE2),
        "2",                                # лист «Препятствия ПЗ-90.11_2024 00.00»
        "1",                                # автопоиск -> подтвердить лучший вариант
        "2",                                # характеристика: столбец C
        "1",                                # добавить ещё файл
        str(FILE3),                         # один лист — выбирается автоматически
        "1",                                # автопоиск -> подтвердить
        "2",                                # характеристика: столбец C
        "2",                                # «Собрать итоговый файл»
        str(out),                           # путь сохранения
        "н",                                # не открывать файл
        "н",                                # не начинать новую настройку
        "3",                                # главное меню -> Выход
        "",
    ])
    result = _run_app(script, cwd=PROJECT_ROOT)
    assert result.returncode == 0, f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    assert "Найдено значений: 48" in result.stdout
    assert "Найдено значений: 577" in result.stdout
    assert "Найдено значений: 25" in result.stdout
    assert "Найдено совпадений: 48 из 48" in result.stdout
    assert "Найдено совпадений: 25 из 48" in result.stdout
    assert "Строк: 577, столбцов: 5" in result.stdout

    import openpyxl

    wb = openpyxl.load_workbook(out)
    ws = wb.active
    assert ws.max_row == 578          # 577 ключей + шапка
    assert ws.max_column == 5         # ключ + 2 + 1 + 1 характеристик
    assert ws.freeze_panes == "A2"

    # Заголовки прочитаны из таблиц автоматически (с единицами измерения)
    headers = [c.value for c in ws[1]]
    assert headers[0] == "№ препятствия"
    assert headers[1] == "Наименование препятствия"
    assert headers[2] == "Sп, м"
    assert headers[3] == "Местоположение"
    assert headers[4] == "Sп, м (2)"

    rows = {str(r[0]): r for r in ws.iter_rows(min_row=2, values_only=True)}
    # «22» есть во всех трёх источниках — все характеристики заполнены
    row22 = rows["22"]
    assert all(v is not None for v in row22[1:])
    # «2» отсутствует в файле 3 -> его столбец пуст, остальные заполнены
    row2 = rows["2"]
    assert row2[1] is not None and row2[2] is not None and row2[3] is not None
    assert row2[4] is None


def test_instruction_screen_and_exit() -> None:
    result = _run_app("1\n\n3\n", cwd=PROJECT_ROOT)
    assert result.returncode == 0
    assert "Для чего нужна программа" in result.stdout


def test_bad_path_then_cancel() -> None:
    # Несуществующий путь -> понятная ошибка, затем отмена в главное меню
    script = "2\nнесуществующий_файл\n\nо\nд\n3\n"
    result = _run_app(script, cwd=PROJECT_ROOT)
    assert result.returncode == 0
    assert "Файл не найден" in result.stdout
