"""Независимая сверка итога трёх примеров с исходниками.

Итог собирается программой, а ожидаемые значения читаются из .xls напрямую
через xlrd — мимо reader/layout/aggregator. Так ловится ошибка в самой
сборке (съезд строк, чужой столбец, потерянный формат), которую внутренняя
проверка записи не видит: та сверяет файл только с собранной таблицей.
"""

from __future__ import annotations

import math
from pathlib import Path

import openpyxl
import pytest
import xlrd
from openpyxl.styles.numbers import BUILTIN_FORMATS

from excel_aggregator.app import AggregatorApp
from excel_aggregator.core import (
    Aggregator,
    Orientation,
    SourceConfig,
    export,
    header_cells,
    load_workbook,
)

SAMPLES = Path(__file__).resolve().parents[2] / "Примеры таблиц"
# файл, лист, первая строка данных (0-based), столбец названий
SOURCES = [
    (SAMPLES / "Каталог Сочи 2024 МОС ФАП-262.xls", "Каталог 2024", 5),
    (SAMPLES / "Каталог Сочи 2024 ПЗ-90.11 00.00.xls", "Препятствия ПЗ-90.11_2024 00.00", 6),
    (SAMPLES / "Превышающие препятствия 2024.xls", "Лист1", 8),
]

pytestmark = pytest.mark.skipif(
    not all(path.is_file() for path, _, _ in SOURCES), reason="Примеры таблиц недоступны"
)


def _key(value) -> str | None:
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    text = str(value).strip()
    return text or None


def _format(book, sheet, r: int, c: int) -> str | None:
    key = book.xf_list[sheet.cell_xf_index(r, c)].format_key
    text = BUILTIN_FORMATS.get(key) or book.format_map[key].format_str
    return None if text == "General" else text


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    """Итог трёх примеров, собранный как в программе (все столбцы)."""
    app = AggregatorApp()
    app.books = [load_workbook(path) for path, _, _ in SOURCES]
    app._analyse()
    app.fold = False
    tables = app._enabled()
    aggregator = Aggregator(app._key_title(), match=app._match)
    for table in tables:
        columns = app._data_columns(table)
        aggregator.add_source(
            table.keyset,
            SourceConfig(
                table.path, table.sheet.name, (table.layout.data_start, table.key_col),
                Orientation.VERTICAL, table.key_col, columns,
                [table.title(c) for c in columns], [table.paths[c] for c in columns],
                header_cells(table.grid, table.layout, columns),
                header_cells(table.grid, table.layout, [table.key_col])[0],
            ),
            table.grid,
        )
    headers, rows, formats = aggregator.build()
    out = tmp_path_factory.mktemp("oracle") / "итог.xlsx"
    export(out, headers, rows, headers[0], title="итог",
           header_blocks=aggregator.header_blocks(), formats=formats)
    return out, [app._data_columns(t) for t in tables]


def test_every_cell_matches_its_source(result):
    out, columns = result
    ws = openpyxl.load_workbook(out, read_only=True).active
    data_row = 7
    written = list(ws.iter_rows(min_row=data_row))
    assert len(written) == 3500

    offset = 1
    mismatches = []
    for (path, sheet_name, start), used in zip(SOURCES, columns):
        book = xlrd.open_workbook(str(path), formatting_info=True)
        sheet = book.sheet_by_name(sheet_name)
        rows = {}
        for r in range(start, sheet.nrows):
            key = _key(sheet.cell_value(r, 0))
            if key is not None and key not in rows:
                rows[key] = r
        for cells in written:
            source_row = rows.get(_key(cells[0].value))
            for j, c in enumerate(used):
                cell = cells[offset + j]
                expected = None if source_row is None else sheet.cell_value(source_row, c)
                if isinstance(expected, str) and not expected.strip():
                    expected = None
                same = cell.value == expected or (
                    isinstance(expected, float) and isinstance(cell.value, (int, float))
                    and math.isclose(cell.value, expected, rel_tol=1e-12)
                )
                if same and isinstance(expected, float):
                    fmt = _format(book, sheet, source_row, c)
                    same = (cell.number_format if fmt else "General") == (fmt or "General")
                if not same:
                    mismatches.append((path.name, cells[0].value, c, expected, cell.value))
        offset += len(used)
    assert not mismatches, mismatches[:10]


def test_header_follows_sources(result):
    out, _ = result
    ws = openpyxl.load_workbook(out).active
    merged = {str(r) for r in ws.merged_cells.ranges}
    values = {
        cell.coordinate: cell.value for row in ws.iter_rows(min_row=1, max_row=6) for cell in row
        if cell.value is not None
    }
    # каталог: «Ап» объединено над «град» и «мин»; высота — над «м» и «фут»
    assert values["E4"] == "Ап" and "E4:F4" in merged
    assert values["E5"] == "град" and values["F5"] == "мин"
    assert values["O3"] == "Абсолютная высота препятствия" and "O3:P5" in merged
    assert values["O6"] == "м" and values["P6"] == "фут"
    # ПЗ-90: «Широта» над четырьмя столбцами, подписи раздела «4 Препятствия» нет
    assert values["S5"] == "Широта" and "S5:V5" in merged
    assert values["Q3"] == "Аэронавигационные данные"
    assert "Препятствия" not in values.values() and "4" not in values.values()
    assert values["AC5"] == "Высота нормаль-\nная\nБалт.1977г"   # переносы как в исходнике
    assert values["S6"] == "°"
    # необъединённая в исходнике группа «Высота геодези-ческая» над «м» и «фут»
    assert "AA5:AB5" in merged


def test_data_rows_have_default_height(result):
    out, _ = result
    ws = openpyxl.load_workbook(out).active
    assert all(ws.row_dimensions[r].height in (None, 15) for r in range(7, ws.max_row + 1))
