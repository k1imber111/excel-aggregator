"""Сквозной e2e-тест: весь сценарий приложения через subprocess.

Гоняет main.py с заранее записанным stdin на реальных примерах из
«Примеры таблиц/»: три файла, автоматическое распознавание таблиц и столбца
названий, выбор столбцов, сборка итогового .xlsx и проверка его содержимого.
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
        "1",                                # главное меню -> Начать работу
        f'"{FILE1}"',                       # путь с кавычками — должны быть сняты
        str(FILE2),
        str(FILE3),
        "",                                 # Enter — файлы добавлены, к проверке
        "",                                 # Enter — найденное верно
        "2,3",                              # файл 1: «Наименование» и «Sп, м»
        "2",                                # файл 2: «Местоположение»
        "2",                                # файл 3: «Sп, м»
        str(out),                           # путь сохранения
        "н",                                # не открывать файл
        "н",                                # не начинать заново
        "3",                                # главное меню -> Выход
        "",
    ])
    result = _run_app(script, cwd=PROJECT_ROOT)
    assert result.returncode == 0, f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    assert "Таблицы соединяются по названиям объектов" in result.stdout
    assert "Строк: 577, столбцов: 5" in result.stdout
    assert "Запись проверена" in result.stdout
    assert "Непредвиденная ошибка" not in result.stdout

    import openpyxl

    wb = openpyxl.load_workbook(out)
    ws = wb.active
    # строка 1 — название таблицы, 2 — источники, 3–4 — шапка, с 5 — данные
    assert ws["A1"].value == "итог"
    assert ws.max_row == 4 + 577
    assert ws.max_column == 5
    assert ws.freeze_panes == "B5"
    assert ws.auto_filter.ref == "A4:E581"

    assert ws["A2"].value == "№ препятствия"
    assert ws["B2"].value == FILE1.stem and ws["D2"].value == FILE2.stem
    assert ws["E2"].value == FILE3.stem
    assert ws["B3"].value == "Наименование препятствия"
    assert ws["C3"].value == "Полярные координаты относительно КТА"
    assert ws["C4"].value == "Sп, м"
    assert ws["D3"].value == "Местоположение"
    assert ws["E3"].value == "Полярные координаты" and ws["E4"].value == "Sп, м"

    rows = {str(r[0]): r for r in ws.iter_rows(min_row=5, values_only=True)}
    assert len(rows) == 577
    # «22» есть во всех трёх источниках — все характеристики заполнены
    assert all(v is not None for v in rows["22"][1:])
    assert rows["22"][1] == "Строение на горе"
    # «2» отсутствует в файле 3 -> его столбец пуст, остальные заполнены
    row2 = rows["2"]
    assert row2[1] == "Трубы котельной" and row2[2] is not None and row2[3] is not None
    assert row2[4] is None
    # объект только из файла 2 — в конце списка, у файла 1 пусто
    assert rows["3626"][1] is None and rows["3626"][3] is not None


def test_folder_and_all_columns(tmp_path: Path) -> None:
    # Самый короткий путь: папка -> четыре Enter -> имя файла
    out = tmp_path / "всё.xlsx"
    script = "\n".join(["1", str(SAMPLES), "", "", "", "", "", str(out), "н", "н", "3", ""])
    result = _run_app(script, cwd=PROJECT_ROOT)
    assert result.returncode == 0, f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    assert "Строк: 577, столбцов: 50" in result.stdout
    assert "28850 ячеек данных совпали" in result.stdout


def test_instruction_screen_and_exit() -> None:
    result = _run_app("2\n\n3\n", cwd=PROJECT_ROOT)
    assert result.returncode == 0
    assert "Для чего нужна программа" in result.stdout


def test_bad_path_then_cancel() -> None:
    # Несуществующий путь -> понятная ошибка, затем отмена в главное меню
    script = "1\nнесуществующий_файл\nо\nд\n3\n"
    result = _run_app(script, cwd=PROJECT_ROOT)
    assert result.returncode == 0
    assert "Файл не найден" in result.stdout
