"""Интеграционные тесты на реальных примерах из «Примеры таблиц/»."""

from pathlib import Path

import pytest

from excel_aggregator.core.aggregator import Aggregator
from excel_aggregator.core.keys import extract_keys
from excel_aggregator.core.matcher import find_key_candidates
from excel_aggregator.core.models import Orientation, SourceConfig, parse_coordinate
from excel_aggregator.core.reader import load_workbook

SAMPLES_DIR = Path(__file__).resolve().parents[2] / "Примеры таблиц"

FILE1 = SAMPLES_DIR / "Каталог Сочи 2024 МОС ФАП-262.xls"
FILE2 = SAMPLES_DIR / "Каталог Сочи 2024 ПЗ-90.11 00.00.xls"
FILE3 = SAMPLES_DIR / "Превышающие препятствия 2024.xls"

SHEET1 = "Каталог 2024"
SHEET2 = "Препятствия ПЗ-90.11_2024 00.00"
SHEET3 = "Лист1"


def _load(path, sheet_name):
    if not path.is_file():
        pytest.skip(f"Пример отсутствует: {path.name}")
    return load_workbook(path).sheet(sheet_name)


@pytest.fixture(scope="module")
def grid1():
    return _load(FILE1, SHEET1)


@pytest.fixture(scope="module")
def grid2():
    return _load(FILE2, SHEET2)


@pytest.fixture(scope="module")
def grid3():
    return _load(FILE3, SHEET3)


def test_file1_keys(grid1):
    row0, col0 = parse_coordinate("A6")
    keyset = extract_keys(grid1, row0, col0, Orientation.VERTICAL)
    assert len(keyset.keys) == 48
    assert keyset.keys[0] == "2"
    assert keyset.keys[-1] == "50"


def test_file2_matcher_finds_key_column(grid1, grid2):
    row0, col0 = parse_coordinate("A6")
    base = extract_keys(grid1, row0, col0, Orientation.VERTICAL)
    candidates = find_key_candidates(grid2, base.keys, Orientation.VERTICAL)
    assert candidates
    best = candidates[0]
    assert best.index == 0
    assert best.match_count == 48


def test_file2_keys(grid2):
    row0, col0 = parse_coordinate("A7")
    keyset = extract_keys(grid2, row0, col0, Orientation.VERTICAL)
    assert len(keyset.keys) == 577


def test_file3_keys(grid3):
    row0, col0 = parse_coordinate("A9")
    keyset = extract_keys(grid3, row0, col0, Orientation.VERTICAL)
    assert len(keyset.keys) == 25


def _build_aggregator(grid1, grid2, grid3):
    """Сквозной сценарий: файл 1 вручную, файлы 2 и 3 через matcher."""
    agg = Aggregator(key_title="№ препятствия")

    ks1 = extract_keys(grid1, *parse_coordinate("A6"), Orientation.VERTICAL)
    agg.add_source(ks1, SourceConfig(
        file_path=FILE1, sheet_name=SHEET1, anchor=parse_coordinate("A6"),
        orientation=Orientation.VERTICAL, key_index=0,
        char_columns=[2, 14],
        column_titles=["Наименование (каталог)", "Высота, м (каталог)"],
    ), grid1)

    candidates2 = find_key_candidates(grid2, ks1.keys, Orientation.VERTICAL)
    best2 = candidates2[0]
    ks2 = extract_keys(grid2, best2.first_row0, best2.index, Orientation.VERTICAL)
    agg.add_source(ks2, SourceConfig(
        file_path=FILE2, sheet_name=SHEET2, anchor=(best2.first_row0, best2.index),
        orientation=Orientation.VERTICAL, key_index=best2.index,
        char_columns=[1, 14],
        column_titles=["Наименование (ПЗ)", "Высота норм., м (ПЗ)"],
    ), grid2)

    candidates3 = find_key_candidates(grid3, ks1.keys, Orientation.VERTICAL)
    best3 = candidates3[0]
    ks3 = extract_keys(grid3, best3.first_row0, best3.index, Orientation.VERTICAL)
    agg.add_source(ks3, SourceConfig(
        file_path=FILE3, sheet_name=SHEET3, anchor=(best3.first_row0, best3.index),
        orientation=Orientation.VERTICAL, key_index=best3.index,
        char_columns=[1, 5],
        column_titles=["Наименование (превышающие)", "Высота, м (превышающие)"],
    ), grid3)
    return agg


def test_end_to_end_union_order_and_values(grid1, grid2, grid3):
    agg = _build_aggregator(grid1, grid2, grid3)
    headers, rows = agg.build_table()

    assert headers[0] == "№ препятствия"
    assert len(headers) == 1 + 2 + 2 + 2
    assert len(rows) == 577

    # union-порядок: сначала 48 ключей файла 1, затем новые из файла 2
    first_keys = [row[0] for row in rows[:48]]
    assert first_keys[0] == "2" and first_keys[-1] == "50"
    assert rows[48][0] not in first_keys   # первый новый ключ файла 2

    # ключ "22" есть во всех трёх источниках -> значения из всех трёх
    row22 = next(row for row in rows if row[0] == "22")
    assert row22[1] is not None   # файл 1, наименование
    assert row22[3] is not None   # файл 2, наименование
    assert row22[5] == "Строение на горе"  # файл 3, наименование

    # ключ, которого нет в файле 3 -> пустые ячейки
    row2 = next(row for row in rows if row[0] == "2")
    assert row2[1] is not None and row2[3] is not None
    assert row2[5] is None and row2[6] is None


def test_end_to_end_coverage(grid1, grid2, grid3):
    agg = _build_aggregator(grid1, grid2, grid3)
    stats = agg.coverage_stats()
    assert [covered for _, covered, _ in stats] == [48, 577, 25]
    assert all(total == 577 for _, _, total in stats)
