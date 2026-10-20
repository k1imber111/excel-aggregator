"""Красная команда, волна 3: взаимодействие правил в core/layout.py на типичных реестрах."""

from __future__ import annotations

import pytest

from excel_aggregator.core.layout import detect_layout, header_cells, layout_from
from excel_aggregator.core.models import SheetGrid

N = None


def G(rows, merged=None):
    return SheetGrid("t", rows, merged or [])


def _layout(grid):
    layout = detect_layout(grid)
    assert layout is not None, "таблица не распознана (detect_layout вернул None)"
    return layout


def _levels(grid, layout, c):
    return [t for _, t in header_cells(grid, layout, [c])[0]]


def _info(layout):
    return f"title={layout.title!r} header_rows={layout.header_rows} data_start={layout.data_start}"


# --- числа в шапке: годы, номера месяцев --------------------------------------

@pytest.mark.parametrize("y1,y2", [(2023, 2024), ("2023", "2024"), (1, 2)])
def test_year_columns_single_header_row_is_not_data(y1, y2):
    """«Объект | 2023 | 2024»: заголовки-годы — шапка, а не первый объект."""
    grid = G([["Объект", y1, y2], ["О1", 1.5, 2.5], ["О2", 3.5, 4.5], ["О3", 5.5, 6.5]])
    layout = _layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0], 1), _info(layout)


def test_year_columns_two_level_header_with_vertical_merge():
    """«Объект» объединён по вертикали, годы в первом ярусе, «план/факт» во втором."""
    grid = G(
        [["Объект", 2023, 2024], [N, "план", "факт"],
         ["О1", 1, 2], ["О2", 3, 4], ["О3", 5, 6]],
        [(0, 2, 0, 1)],
    )
    layout = _layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0, 1], 2), _info(layout)


def test_years_as_subheader_under_group_are_not_an_object():
    """«Показатели» над годами 2022–2024: строка лет — уровень шапки."""
    grid = G(
        [["Объект", "Показатели", N, N], [N, 2022, 2023, 2024],
         ["О1", 1, 2, 3], ["О2", 4, 5, 6], ["О3", 7, 8, 9]],
        [(0, 2, 0, 1), (0, 1, 1, 4)],
    )
    layout = _layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0, 1], 2), _info(layout)


# --- текстовая таблица: подзаголовок под объединённой группой -----------------

def test_text_table_subheader_under_hmerged_group_without_vertical_merge():
    """Текстовый реестр: «Адрес» объединён по горизонтали над «улица | дом»,
    «Организация» не объединена вниз. Строка «улица | дом» не объект."""
    grid = G(
        [["Организация", "Адрес", N, "Телефон"], [N, "улица", "дом", N],
         ["ООО А", "Ленина", "1а", "111-22"], ["ООО Б", "Мира", "2б", "333-44"],
         ["ООО В", "Тверская", "3в", "555-66"]],
        [(0, 1, 1, 3)],
    )
    layout = _layout(grid)
    assert layout.data_start == 2, _info(layout)


# Решение (волна 3, L3-3): в узкой текстовой таблице без чисел объединение
# «Района» по вертикали в данных неотличимо от вертикально объединённой шапки.
# Таблица не распознаётся автоматически, но не молча: шаг 2 просит указать
# первую строку данных, и с ней шапка и данные определяются верно.

def _narrow_district_grid():
    return G(
        [["Район", "Адрес", N], [N, "улица", "дом"],
         ["Север", "Ленина", "1а"], [N, "Мира", "2б"],
         ["Юг", "Тверская", "3в"], [N, "Арбат", "4г"]],
        [(0, 2, 0, 1), (0, 1, 1, 3), (2, 4, 0, 1), (4, 6, 0, 1)],
    )


def test_text_table_three_columns_vertically_merged_district_two_level_header():
    grid = _narrow_district_grid()
    layout = detect_layout(grid)
    if layout is not None:                      # распознана — тогда верно
        assert layout.data_start == 2, _info(layout)
        return
    manual = layout_from(grid, 2)               # пользователь указал строку 3
    assert manual.header_rows == [0, 1]
    assert _levels(grid, manual, 2) == ["Адрес", "дом"]


def test_unrecognised_sheet_asks_for_first_data_row(tmp_path):
    import openpyxl

    from excel_aggregator.app import AggregatorApp
    from excel_aggregator.core import load_workbook

    book = openpyxl.Workbook()
    ws = book.active
    for row in [["Север", "Ленина", "школа"], [N, "Мира", "сад"], [N, "Тверская", "дом"],
                ["Юг", "Арбат", "парк"], [N, "Тихая", "кафе"]]:
        ws.append(row)
    ws.merge_cells("A1:A3")
    ws.merge_cells("A4:A5")
    book.save(tmp_path / "р.xlsx")
    app = AggregatorApp()
    app.books = [load_workbook(tmp_path / "р.xlsx")]
    app._analyse()
    table = app.tables[0]
    if table.layout is None:                    # не молча: источник помечен, причина названа
        assert not table.enabled and "первую строку данных" in table.problem
    else:
        assert table.layout.data_start == 0


# --- text_object / подзаголовок с заполненным первым столбцом -----------------

def test_subheader_row_with_first_column_label_is_not_text_object():
    """«Объект | Адрес(улица, дом)»: во втором ярусе первого столбца тоже стоит
    подпись («Название»), строка «Название | улица | дом» — шапка, не объект."""
    grid = G(
        [["Объект", "Адрес", N, "Высота", "Ширина", "Год"],
         ["Название", "улица", "дом", N, N, N],
         ["О1", "Ленина", "1а", 10, 20, 30], ["О2", "Мира", "2б", 11, 21, 31],
         ["О3", "Тверская", "3в", 12, 22, 32]],
        [(0, 1, 1, 3)],
    )
    layout = _layout(grid)
    assert layout.data_start == 2, _info(layout)


# --- подпись + дата составления -------------------------------------------------

def test_title_and_date_in_one_sparse_row_are_not_header_levels():
    """«Реестр объектов» слева и «Дата: …» справа в одной строке над таблицей:
    это подпись, а не уровень шапки (иначе названия столбцов портятся)."""
    grid = G(
        [["Реестр объектов", N, N, "Дата: 01.01.2024"],
         ["Объект", "Высота", "Ширина", "Год"],
         ["О1", 1, 2, 3], ["О2", 4, 5, 6], ["О3", 7, 8, 9]],
    )
    layout = _layout(grid)
    assert layout.header_rows == [1], _info(layout)
    assert _levels(grid, layout, 3) == ["Год"]


def test_title_then_company_and_date_row_not_header_levels():
    """Строка 0 — подпись на всю ширину, строка 1 — «Организация | … | Дата»."""
    grid = G(
        [["Реестр объектов", N, N, N],
         ["Организация: ООО Ромашка", N, N, "Дата: 01.01.2024"],
         ["Объект", "Высота", "Ширина", "Год"],
         ["О1", 1, 2, 3], ["О2", 4, 5, 6], ["О3", 7, 8, 9]],
        [(0, 1, 0, 4)],
    )
    layout = _layout(grid)
    assert layout.header_rows == [2], _info(layout)


# --- _is_section_caption: расщеплённое «№ / п/п» ----------------------------

def test_split_number_header_keeps_second_line_next_to_units():
    """«№» над «п/п», в этой же строке «м | м» под «Высота | Широта»: строка
    считается строкой единиц, и «п/п» стирается как подпись раздела."""
    grid = G(
        [["№", "Наименование", "Высота", "Ширина"],
         ["п/п", N, "м", "м"],
         [1, "О1", 1, 2], [2, "О2", 3, 4], [3, "О3", 5, 6]],
        [(0, 2, 1, 2)],
    )
    layout = _layout(grid)
    assert layout.data_start == 2, _info(layout)
    assert _levels(grid, layout, 0) == ["№", "п/п"]


# --- комбинация: подпись + дата + многоуровневая шапка + единицы + нумерация --

def _registry(with_numbering=True, with_title_cells=True):
    rows = [
        ["Реестр аэродромов", N, N, N, N, N],
        ["Составлено 01.01.2024", N, N, N, N, N],
        [N] * 6,
        ["№", "Объект", "Координаты", N, "Высота", "Год"],
        [N, N, "Широта", "Долгота", N, N],
        [N, N, "град", "град", "м", N],
        [1, 2, 3, 4, 5, 6] if with_numbering else [N] * 6,
        ["Район 1", N, N, N, N, N],
        [1, "О1", 55.1, 37.2, 100, 2001],
        [2, "О2", 55.2, 37.3, 101, 2002],
        [3, "О3", 55.3, 37.4, 102, 2003],
        ["Район 2", N, N, N, N, N],
        [4, "О4", 55.4, 37.5, 103, 2004],
        [5, "О5", 55.5, 37.6, 104, 2005],
        ["Итого", N, N, N, 5, N],
    ]
    merged = [(0, 1, 0, 6), (1, 2, 0, 6), (3, 6, 0, 1), (3, 6, 1, 2), (3, 4, 2, 4),
              (3, 6, 4, 5), (3, 6, 5, 6), (7, 8, 0, 6), (11, 12, 0, 6)]
    return G(rows, merged)


@pytest.mark.parametrize("numbering", [True, False])
def test_registry_combo_title_date_multilevel_units_numbering_districts(numbering):
    grid = _registry(numbering)
    layout = _layout(grid)
    assert layout.title == "Реестр аэродромов"
    assert layout.header_rows == [3, 4, 5], _info(layout)
    assert layout.data_start == 8, _info(layout)
    assert _levels(grid, layout, 2) == ["Координаты", "Широта", "град"]
    assert _levels(grid, layout, 4)[0] == "Высота"
