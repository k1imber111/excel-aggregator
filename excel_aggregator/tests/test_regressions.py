"""Регрессионные тесты по аудиту 2026-10-06: по одному на найденную ошибку."""

from __future__ import annotations

import datetime
import io
from pathlib import Path

import openpyxl
import pytest
from rich.console import Console

import excel_aggregator.app as app_module
from excel_aggregator.app import AggregatorApp
from excel_aggregator.core.errors import OutputWriteError
from excel_aggregator.core.exporter import export
from excel_aggregator.core.models import Orientation, SheetGrid
from excel_aggregator.core.paths import resolve_input_path
from excel_aggregator.core.reader import load_workbook
from excel_aggregator.ui.widgets import UI

SAMPLE_WITH_DATES = (
    Path(__file__).resolve().parents[2] / "Примеры таблиц" / "Каталог Сочи 2024 МОС ФАП-262.xls"
)


def _make_xlsx(path: Path, rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


@pytest.fixture
def table(tmp_path: Path) -> Path:
    """Таблица с шапкой: ключи в A2:A4, две графы характеристик."""
    return _make_xlsx(tmp_path / "таблица.xlsx", [
        ["Название", "Высота", "Ширина"],
        ["A", 1, 10],
        ["B", 2, 20],
        ["C", 3, 30],
    ])


def _drive(answers: list[str], monkeypatch) -> tuple[AggregatorApp, str]:
    """Прогоняет приложение в памяти с заранее заданными ответами."""
    monkeypatch.setenv("AGGREGATOR_SCRIPTED", "1")
    app = AggregatorApp()
    buf = io.StringIO()
    app.ui.console = Console(file=buf, width=200)
    queue = list(answers)
    app.ui.console.input = lambda prompt="": queue.pop(0)
    app.run()
    assert not queue, f"Остались непрочитанные ответы: {queue}"
    return app, buf.getvalue()


# --- (1) «назад» через границу блока «очередной файл» -----------------------

def test_back_across_file_blocks_does_not_duplicate_sources(table, monkeypatch):
    first = ["2", str(table), "A2", "1", "1"]     # меню, файл, якорь, вертикально, графа B
    second = ["1", str(table), "1", "1"]          # добавить файл, путь, подтвердить, графа B
    # добавить 3-й файл -> «н» на вводе пути -> «н» в меню -> перенастроить 2-й
    back = ["1", "н", "н", "1", "1"]
    app, _ = _drive(first + second + back + ["о", "д", "3"], monkeypatch)
    assert len(app.sources) == 2
    assert len(app._blocks) == 1


# --- страховка: непредвиденная ошибка не выбрасывает из приложения ----------

def test_unexpected_error_is_logged_and_offers_retry(tmp_path, monkeypatch):
    monkeypatch.setattr(app_module, "_program_dir", lambda: tmp_path)

    def boom(self):
        raise RuntimeError("сбой шага")

    monkeypatch.setattr(AggregatorApp, "_step_first_file", boom)
    _, out = _drive(["2", "н", "3"], monkeypatch)   # запуск, не повторять, выход
    assert "Непредвиденная ошибка" in out
    assert "сбой шага" in (tmp_path / "aggregator_error.log").read_text(encoding="utf-8")


# --- (2) даты .xls, (3) даты и (4) недопустимые символы при экспорте --------

@pytest.mark.skipif(not SAMPLE_WITH_DATES.is_file(), reason="Пример недоступен")
def test_xls_date_cells_are_dates():
    grid = load_workbook(SAMPLE_WITH_DATES).sheet("Т12")
    assert grid.value(0, 1) == datetime.datetime(2018, 4, 24)


def test_export_keeps_dates_and_strips_illegal_characters(tmp_path):
    out = tmp_path / "итог.xlsx"
    when = datetime.datetime(2024, 1, 15)
    export(out, ["Ключ", "Да\x0bта"], [["k\x01", when]], key_title="Ключ")
    ws = openpyxl.load_workbook(out).active
    assert ws["B1"].value == "Дата"
    assert ws["A2"].value == "k"
    assert ws["B2"].value == when


# --- (5) плохой путь сохранения -> понятная ошибка ---------------------------

def test_export_unwritable_folder_is_output_write_error(tmp_path):
    blocker = tmp_path / "это_файл"
    blocker.write_text("x", encoding="utf-8")
    with pytest.raises(OutputWriteError):
        export(blocker / "папка" / "итог.xlsx", ["A"], [["1"]], key_title="A")


# --- (6) rich-разметка в данных ----------------------------------------------

def test_brackets_in_data_are_printed_literally():
    ui = UI()
    buf = io.StringIO()
    ui.console = Console(file=buf, width=200)
    ui.info("Таблица [final].xls и хвост [/]")
    ui.error("значение [/b]")
    ui.show_preview(SheetGrid("лист [a]", [["x [/b]", "y [red]"]]), title="Лист [a]")
    out = buf.getvalue()
    for fragment in ("[final]", "хвост [/]", "значение [/b]", "x [/b]", "y [red]", "Лист [a]"):
        assert fragment in out


# --- (7) расширение дописывается к имени, а не заменяет хвост ----------------

def test_output_name_with_dot_is_not_truncated(table, tmp_path, monkeypatch):
    target = tmp_path / "отчёт v1.2"
    answers = ["2", str(table), "A2", "1", "1", "2", str(target), "н", "н", "3"]
    _drive(answers, monkeypatch)
    assert (tmp_path / "отчёт v1.2.xlsx").is_file()


def test_input_name_with_dot_gets_extension(tmp_path):
    f = tmp_path / "данные.2024.xls"
    f.write_bytes(b"\xd0\xcf\x11\xe0")
    assert resolve_input_path(str(tmp_path / "данные.2024")) == f


# --- (8) заголовки в таблице из двух столбцов --------------------------------

def test_two_column_table_headers_are_detected():
    grid = SheetGrid("s", [["Название", "Высота"], ["A", 1], ["B", 2]])
    assert AggregatorApp._guess_header(grid, 0, 1, 0, Orientation.VERTICAL) == "Название"
    assert AggregatorApp._guess_header(grid, 1, 1, 0, Orientation.VERTICAL) == "Высота"


# --- (9) однобуквенное название графы — не команда ---------------------------

def test_single_letter_title_is_not_a_command(tmp_path, monkeypatch):
    bare = _make_xlsx(tmp_path / "без шапки.xlsx", [["A", 1], ["B", 2]])
    # заголовков нет -> программа спрашивает название ключа и графы
    answers = ["2", str(bare), "A1", "1", "Y", "1", "Н", "о", "д", "3"]
    app, _ = _drive(answers, monkeypatch)
    assert app.key_title == "Y"
    assert app.sources[0][1].column_titles == ["Н"]
