"""Красная команда, зона 2: новые правила распознавания в core/layout.py (коммит 831f9e4)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_aggregator.core.layout import (
    column_paths,
    detect_layout,
    header_cells,
    header_text,
)
from excel_aggregator.core.models import SheetGrid
from excel_aggregator.core.reader import load_workbook

N = None
SAMPLES = Path(__file__).resolve().parents[3] / "Примеры таблиц"


def G(rows, merged=None, formats=None):
    return SheetGrid("t", rows, merged or [], formats)


def _layout(grid):
    layout = detect_layout(grid)
    assert layout is not None
    return layout


# --- _merged_share: считает только горизонтальные объединения ----------------

def test_vertical_only_header_text_table_subheader_is_not_object():
    """Шапка 2 яруса, «Объект»/«Тип» объединены вертикально, остальное — текст:
    строка подзаголовков «подгр/подкл» не должна стать объектом."""
    grid = G(
        [
            ["Объект", "Тип", "Группа", "Класс"],
            [N, N, "подгр", "подкл"],
            ["О1", "А", "x", "y"], ["О2", "Б", "x", "y"], ["О3", "В", "x", "y"],
        ],
        [(0, 2, 0, 1), (0, 2, 1, 2)],
    )
    layout = _layout(grid)
    assert layout.data_start == 2, f"data_start={layout.data_start}, header_rows={layout.header_rows}"


def test_vertical_only_header_single_data_row():
    """Одна строка данных под двухъярусной шапкой с вертикальными объединениями:
    подзаголовки «макс/мин» не должны стать объектом."""
    grid = G(
        [
            ["Объект", "Тип", "Высота", "Ширина"],
            [N, N, "макс", "мин"],
            ["О1", "А", 10, 20],
        ],
        [(0, 2, 0, 1), (0, 2, 1, 2)],
    )
    layout = _layout(grid)
    assert layout.data_start == 2, f"data_start={layout.data_start}, header_rows={layout.header_rows}"


def test_vertical_only_three_level_header_two_data_rows():
    """Три яруса шапки (вертикальные объединения), две строки данных."""
    grid = G(
        [
            ["Объект", "Тип", "Высота", "Ширина"],
            [N, N, "макс", "мин"],
            [N, N, "лето", "зима"],
            ["О1", "А", 10, 20], ["О2", "Б", 30, 40],
        ],
        [(0, 3, 0, 1), (0, 3, 1, 2)],
    )
    layout = _layout(grid)
    assert layout.data_start == 3, f"data_start={layout.data_start}, header_rows={layout.header_rows}"


# --- text_object: разреженная строка над блоком данных -----------------------

def test_text_object_does_not_swallow_subheader_row():
    """«Адрес» разбит на «улица/дом» (подзаголовок только в текстовых столбцах),
    над ним плотная строка шапки. Строка «улица/дом» — шапка, а не объект."""
    grid = G(
        [
            ["Объект", "Адрес", N, "Высота", "Ширина", "Глубина", "Год", "Прим"],
            [N, "улица", "дом", N, N, N, N, N],
            ["О1", "Ленина", "1а", 10, 20, 30, 2001, "a"],
            ["О2", "Мира", "2б", 11, 21, 31, 2002, "b"],
            ["О3", "Тверская", "3в", 12, 22, 32, 2003, "c"],
        ]
    )
    layout = _layout(grid)
    assert layout.data_start == 2, f"data_start={layout.data_start}, header_rows={layout.header_rows}"


# --- нумерация граф над одними числовыми столбцами ---------------------------

def test_numbering_row_in_all_numeric_table_is_not_first_object():
    """Таблица из одних чисел (ID, X, Y, Z) с настоящей нумерацией граф «1 2 3 4»:
    нумерация не объект."""
    grid = G(
        [
            ["ID", "X", "Y", "Z"],
            [1, 2, 3, 4],
            [101, 1.5, 2.5, 3.5], [102, 1.6, 2.6, 3.6], [103, 1.7, 2.7, 3.7],
        ]
    )
    layout = _layout(grid)
    assert (layout.numbering_row, layout.data_start) == (1, 2), (
        f"numbering={layout.numbering_row}, data_start={layout.data_start}"
    )


def test_numbering_over_numeric_columns_with_unnumbered_name_column():
    """Столбец названий без номера, нумерация только над числовыми графами."""
    grid = G(
        [
            ["Объект", "X", "Y", "Z"],
            [N, 1, 2, 3],
            ["А", 1.5, 2.5, 3.5], ["Б", 1.6, 2.6, 3.6], ["В", 1.7, 2.7, 3.7],
        ]
    )
    layout = _layout(grid)
    assert (layout.numbering_row, layout.data_start) == (1, 2), (
        f"numbering={layout.numbering_row}, data_start={layout.data_start}"
    )


# --- подпись раздела в строке единиц; расширенный список единиц --------------

def test_unit_row_subheader_left_of_units_keeps_its_text():
    """Подзаголовок «Код» под «Склад» стоит левее «шт/кг/т» (теперь единицы):
    header_text не должен стирать настоящий заголовок."""
    grid = G(
        [
            ["Склад", "Остаток", N, N],
            ["Код", "шт", "кг", "т"],
            ["С1", 1, 2, 3], ["С2", 4, 5, 6], ["С3", 7, 8, 9],
        ],
        [(0, 1, 1, 4)],
    )
    layout = _layout(grid)
    assert layout.header_rows == [0, 1]
    assert header_text(grid, layout, 1, 0) == "Код"


def _weather():
    return G(
        [
            ["Станция", "Погода", N, N],
            [N, "Давление", "Т", "%"],
            ["С1", 1, 2, 3], ["С2", 4, 5, 6], ["С3", 7, 8, 9],
        ],
        [(0, 2, 0, 1), (0, 1, 1, 4)],
    )


def test_real_header_t_and_percent_do_not_erase_neighbour_level():
    """Подгруппа «Погода»: «Давление | Т | %». «Т» и «%» стали единицами, строка
    считается строкой единиц, а «Давление» — «подписью раздела» и пропадает из шапки."""
    grid = _weather()
    layout = _layout(grid)
    levels = header_cells(grid, layout, [1])[0]
    assert [t for _, t in levels] == ["Погода", "Давление"], f"levels={levels}"


def test_column_paths_keeps_neighbour_level_next_to_t_and_percent():
    """column_paths: «Давление» не должно пропадать из пути столбца."""
    grid = _weather()
    layout = _layout(grid)
    assert column_paths(grid, layout)[1] == ["Погода", "Давление"]


# --- подпись над таблицей: «все частые столбцы» -------------------------------

def test_group_over_all_frequent_columns_is_header_level_not_title():
    """Группа «Размеры» над тремя числовыми столбцами; «Код» и «Прим» — редкие
    столбцы по краям (заполнены в 1 из 3 строк), но у них есть заголовки в шапке.
    Группа — уровень шапки; как «подпись» она исчезает из шапки итога."""
    grid = G(
        [
            [N, "Размеры", N, N, N],
            ["Код", "Высота", "Ширина", "Глубина", "Прим"],
            [N, 10, 20, 30, N], [N, 11, 21, 31, N], ["к", 12, 22, 32, "x"],
        ],
        [(0, 1, 1, 4)],
    )
    layout = _layout(grid)
    assert layout.title == "" and 0 in layout.header_rows, (
        f"title={layout.title!r}, header_rows={layout.header_rows}"
    )


# --- реальные примеры: не регрессировать ---------------------------------------

_REAL = [
    ("Каталог Сочи 2024 МОС ФАП-262.xls", "Каталог 2024", [0, 1, 2, 3], 5),
    ("Каталог Сочи 2024 ПЗ-90.11 00.00.xls", "Препятствия ПЗ-90.11_2024 00.00", [0, 1, 2, 4], 6),
    ("Превышающие препятствия 2024.xls", "Лист1", [3, 5, 6], 8),
]


@pytest.mark.parametrize("fname,sheet,header_rows,data_start", _REAL)
def test_real_samples_layout_and_header_levels_intact(fname, sheet, header_rows, data_start):
    path = SAMPLES / fname
    if not path.is_file():
        pytest.fail(f"нет примера {path}")
    book = load_workbook(path)
    grid = next(s for s in book.sheets if s.name == sheet)
    layout = _layout(grid)
    assert (layout.header_rows, layout.data_start) == (header_rows, data_start)
    # шапка итога не теряет уровней: у каждого столбца столько уровней, сколько строк шапки
    for levels in header_cells(grid, layout, list(range(grid.ncols))):
        assert len(levels) == len(layout.header_rows)
