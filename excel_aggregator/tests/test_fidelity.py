"""Верность исходникам: форматы чисел, шапка как в исходнике, скрытые листы."""

from __future__ import annotations

import datetime
import os
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import openpyxl

from excel_aggregator.app import AggregatorApp
from excel_aggregator.core import (
    Aggregator,
    Orientation,
    SheetGrid,
    SourceConfig,
    detect_layout,
    export,
    extract_keys,
    header_cells,
    layout_from,
    load_workbook,
)
from excel_aggregator.core import exporter, reader

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _xlsx(path: Path, sheets: dict[str, list[list]], hidden: tuple[str, ...] = (),
          formats: dict[tuple[int, int], str] | None = None,
          merges: list[str] | None = None) -> Path:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
        for (r, c), fmt in (formats or {}).items():
            ws.cell(row=r, column=c).number_format = fmt
        for rng in merges or []:
            ws.merge_cells(rng)
        if name in hidden:
            ws.sheet_state = "hidden"
    wb.save(path)
    return path


# --- форматы -------------------------------------------------------------------

def test_reader_keeps_number_formats_and_drops_blank_text(tmp_path):
    path = _xlsx(tmp_path / "ф.xlsx", {"Л": [
        ["Код", "Длина", "Доля", "Дата", "Примечание"],
        ["a", 1.23456, 0.25, datetime.datetime(2024, 3, 5), "   "],
    ]}, formats={(2, 2): "0.0", (2, 3): "0.00%", (2, 4): "DD.MM.YYYY"})
    grid = load_workbook(path).sheets[0]
    assert grid.number_format(1, 1) == "0.0"
    assert grid.number_format(1, 2) == "0.00%"
    assert grid.number_format(1, 3) == "DD.MM.YYYY"
    assert grid.number_format(1, 0) is None          # общий формат
    assert grid.value(1, 4) is None                  # ячейка из пробелов — пустая


def test_xls_builtin_format_uses_excel_code():
    # Встроенный код 14 (дата) xlrd отдаёт как «m/d/yy»; нужен код Excel
    book = SimpleNamespace(
        xf_list=[SimpleNamespace(format_key=14), SimpleNamespace(format_key=170)],
        format_map={14: SimpleNamespace(format_str="m/d/yy"),
                    170: SimpleNamespace(format_str="0.000")},
    )
    assert reader._xls_format(book, 0) == "mm-dd-yy"
    assert reader._xls_format(book, 1) == "0.000"
    assert reader._xls_format(book, 99) is None


def test_export_applies_source_formats(tmp_path):
    out = tmp_path / "итог.xlsx"
    when = datetime.datetime(2024, 3, 5)
    rows = [["k", 41189.053493908126, 0.25, when, datetime.datetime(2024, 3, 5, 14, 30), "040"]]
    formats = [[None, "0.0", "0%", None, None, "0.0"]]
    export(out, ["К", "S", "Доля", "Дата", "Время", "Текст"], rows, "К", formats=formats)
    ws = openpyxl.load_workbook(out).active
    assert ws["B2"].number_format == "0.0"
    assert ws["C2"].number_format == "0%"
    assert ws["D2"].number_format == "DD.MM.YYYY"          # дата без формата — русский вид
    assert ws["E2"].number_format == "DD.MM.YYYY HH:MM"
    assert ws["F2"].value == "040" and ws["F2"].number_format == "@"   # текст — текстом
    assert ws.row_dimensions[2].height in (None, 15)       # длинная дробь не раздувает строку
    assert ws.column_dimensions["B"].width < 15             # ширина по виду «41189.1»


def test_display_text_follows_format():
    assert exporter._display_text(41189.053493908126, "0.0") == "41189.1"
    assert exporter._display_text(0.256, "0.0%") == "25.6%"
    assert exporter._display_text(1234567.0, "#,##0") == "1,234,567"
    assert exporter._display_text(40, "000") == "040"
    assert exporter._display_text(-5.5, "0.00") == "-5.50"
    assert len(exporter._display_text(1 / 3, None)) <= 12


def test_verify_detects_lost_format(tmp_path):
    out = tmp_path / "итог.xlsx"
    export(out, ["A", "B"], [["k", 1.5]], "A", formats=[[None, "0.0"]])
    assert exporter._verify(out, [["k", 1.5]], 2, 2, [[None, "0.0"]]) == 2
    try:
        exporter._verify(out, [["k", 1.5]], 2, 2, [[None, "0.000"]])
    except exporter.OutputWriteError as exc:
        assert "формат" in exc.message
    else:
        raise AssertionError("потерянный формат не замечен")


# --- шапка как в исходнике -----------------------------------------------------

def _grid_with_units() -> SheetGrid:
    # Шапка в духе ПЗ-90: группа «Широта» над ° ' ", строка единиц начинается
    # с подписи раздела «4 | Препятствия», «Высота» не объединена над м/фут
    data = [
        ["№", "Название", "Широта", None, None, "Высота", None],
        ["4", "Препятствия", 0, "'", '"', "м", "фут"],
        [1, "Мачта", "43", "41", "51.30", 10.5, 34],
        [2, "Труба", "43", "39", "04.61", 20.0, 66],
        [3, "Дерево", "43", "38", "42.73", 5.25, 17],
    ]
    return SheetGrid("Л", data, [(0, 1, 2, 5)])


def test_header_cells_drop_section_caption_and_keep_units():
    grid = _grid_with_units()
    layout = layout_from(grid, 2)
    cells = header_cells(grid, layout, [0, 1, 2, 3, 4, 5, 6])
    texts = [[text for _, text in col] for col in cells]
    assert texts[0] == ["№", ""] and texts[1] == ["Название", ""]   # «4 Препятствия» — нет
    assert texts[2] == ["Широта", "°"] and texts[3] == ["Широта", "'"]
    assert texts[5] == ["Высота", "м"] and texts[6] == ["Высота", "фут"]
    assert cells[2][0][0] == cells[4][0][0]       # одна объединённая ячейка «Широта»
    assert cells[5][0][0] == cells[6][0][0]       # необъединённая группа «Высота» — тоже


def _export_grid(tmp_path, grid, columns, layout=None):
    layout = layout or layout_from(grid, 2)
    keyset = extract_keys(grid, layout.data_start, 0, Orientation.VERTICAL)
    aggregator = Aggregator("№")
    aggregator.add_source(keyset, SourceConfig(
        tmp_path / "исх.xlsx", "Л", (layout.data_start, 0), Orientation.VERTICAL, 0, columns,
        [str(c) for c in columns], [[str(c)] for c in columns],
        header_cells(grid, layout, columns), header_cells(grid, layout, [0])[0],
    ), grid)
    headers, rows, formats = aggregator.build()
    out = tmp_path / "итог.xlsx"
    export(out, headers, rows, "№", title="итог",
           header_blocks=aggregator.header_blocks(), formats=formats)
    ws = openpyxl.load_workbook(out).active
    return ws, {str(r) for r in ws.merged_cells.ranges}


def test_export_header_matches_source_structure(tmp_path):
    ws, merged = _export_grid(tmp_path, _grid_with_units(), [1, 2, 3, 4, 5, 6])
    # строка 2 — источник, 3–4 — шапка как в исходнике
    assert ws["A2"].value == "№" and "A2:A4" in merged
    assert ws["B3"].value == "Название" and "B3:B4" in merged    # пусто под ним — к нему
    assert ws["C3"].value == "Широта" and "C3:E3" in merged
    assert [ws["C4"].value, ws["D4"].value, ws["E4"].value] == ["°", "'", '"']
    assert ws["F3"].value == "Высота" and "F3:G3" in merged      # необъединённая группа
    assert ws["F4"].value == "м" and ws["G4"].value == "фут"
    assert ws.auto_filter.ref == "A4:G7"


def test_export_header_with_gap_in_selection(tmp_path):
    # Выбраны «°» и «"» без «'»: объединение «Широта» сужается до выбранных
    ws, merged = _export_grid(tmp_path, _grid_with_units(), [2, 4])
    assert ws["B3"].value == "Широта" and "B3:C3" in merged
    assert ws["B4"].value == "°" and ws["C4"].value == '"'


def test_sources_with_different_header_depth(tmp_path):
    deep = _grid_with_units()
    flat = SheetGrid("Л2", [["№", "Цвет"], [1, "красный"], [2, "синий"], [3, "белый"]])
    aggregator = Aggregator("№")
    for grid, layout, columns in ((deep, layout_from(deep, 2), [5, 6]),
                                  (flat, detect_layout(flat), [1])):
        keyset = extract_keys(grid, layout.data_start, 0, Orientation.VERTICAL)
        aggregator.add_source(keyset, SourceConfig(
            tmp_path / f"{grid.name}.xlsx", grid.name, (layout.data_start, 0),
            Orientation.VERTICAL, 0, columns, ["x"] * len(columns), [["x"]] * len(columns),
            header_cells(grid, layout, columns), header_cells(grid, layout, [0])[0],
        ), grid)
    headers, rows, formats = aggregator.build()
    out = tmp_path / "итог.xlsx"
    export(out, headers, rows, "№", header_blocks=aggregator.header_blocks(), formats=formats)
    ws = openpyxl.load_workbook(out).active
    merged = {str(r) for r in ws.merged_cells.ranges}
    assert ws["D2"].value == "Цвет" and "D2:D3" in merged      # мелкая шапка растянута вниз
    assert ws["B2"].value == "Высота" and "B2:C2" in merged
    assert [row[3] for row in ws.iter_rows(min_row=4, values_only=True)] == [
        "красный", "синий", "белый"
    ]


def test_non_rectangular_region_is_not_merged():
    ids, texts = exporter._resolve_header([
        [("s", "A"), ("x", "Группа")],
        [("s", "A"), ("y", "Б")],
        [("s", "A"), ("x", "Группа")],     # тот же id через столбец — не прямоугольник
    ])
    regions = exporter._regions(ids)
    pieces = [r for r in regions if r[0] == "x"]
    assert all(k0 == k1 and j0 == j1 for _, k0, k1, j0, j1 in pieces) and len(pieces) == 2


# --- скрытые листы и запуск с путями -------------------------------------------

def _table(rows_count: int = 4) -> list[list]:
    return [["Код", "Значение"]] + [[f"к{i}", i * 1.5] for i in range(1, rows_count + 1)]


def test_hidden_sheets_are_skipped(tmp_path):
    path = _xlsx(tmp_path / "к.xlsx", {"Видимый": _table(), "Скрытый": _table(6)},
                 hidden=("Скрытый",))
    app = AggregatorApp()
    app.books = [load_workbook(path)]
    app._analyse()
    assert [t.sheet.name for t in app.tables] == ["Видимый"]
    assert app.hidden_sheets == 1


def test_hidden_table_used_when_no_visible_table(tmp_path):
    path = _xlsx(tmp_path / "к.xlsx", {"Пусто": [["заметка"]], "Скрытый": _table()},
                 hidden=("Скрытый",))
    app = AggregatorApp()
    app.books = [load_workbook(path)]
    app._analyse()
    assert [t.sheet.name for t in app.tables] == ["Скрытый"]
    assert app.tables[0].keyset is not None


def test_files_from_command_line_skip_main_menu(tmp_path):
    source = _xlsx(tmp_path / "исх.xlsx", {"Л": _table()})
    out = tmp_path / "итог.xlsx"
    env = dict(os.environ, AGGREGATOR_SCRIPTED="1", PYTHONIOENCODING="utf-8")
    script = "\n".join(["", "", "", str(out), "н", "н", "3", ""])
    result = subprocess.run(
        [sys.executable, str(PROJECT_ROOT / "main.py"), str(source)],
        input=script, capture_output=True, text=True, encoding="utf-8",
        cwd=tmp_path, env=env, timeout=120,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Файл загружен: исх.xlsx" in result.stdout
    assert "Запись проверена" in result.stdout
    assert openpyxl.load_workbook(out).active["B4"].value == 1.5
