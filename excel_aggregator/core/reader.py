"""Чтение книг Excel (.xls через xlrd, .xlsx через openpyxl) в SheetGrid.

Формат определяется по сигнатуре файла, а не по расширению.
Merged cells разворачиваются, ячейки-ошибки (#REF! и т.п.) и ячейки из одних
пробелов -> None, реальная область данных вычисляется по факту заполнения
(DIMENSIONS не доверяем). Вместе со значениями читаются форматы чисел
(«0.0», «DD.MM.YYYY») и признак скрытого листа.
"""

from __future__ import annotations

import re
import zipfile
from datetime import date, datetime, time, timedelta
from pathlib import Path

import xlrd
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.utils.datetime import to_excel

from .errors import (
    CorruptedFileError,
    EncryptedFileError,
    NotExcelFileError,
    PathNotFoundError,
)
from .models import SheetGrid, Workbook
from .normalize import strip_invisible

_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0"
_ZIP_SIGNATURE = b"PK"

# Строковые значения ячеек-ошибок Excel
_ERROR_VALUES = {
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    "#GETTING_DATA", "#SPILL!", "#CALC!",
}


def _is_empty(v) -> bool:
    """Пустая ячейка: None или строка из одних пробелов и невидимых символов."""
    return v is None or (isinstance(v, str) and strip_invisible(v).strip() == "")


def _real_area(data: list[list]) -> tuple[int, int]:
    """(nrows, ncols) реальной области по фактически заполненным ячейкам."""
    max_r = -1
    max_c = -1
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            if not _is_empty(v):
                if r > max_r:
                    max_r = r
                if c > max_c:
                    max_c = c
    return max_r + 1, max_c + 1


def _format(text: str | None) -> str | None:
    """Формат числа ячейки; общий формат -> None."""
    if not text or text.casefold() == "general":
        return None
    return text


# Знаки цифр формата вне кавычек и экранирования: «0», «#», «?»
_DIGIT_PLACEHOLDER_RE = re.compile(r"[0#?]")
_FORMAT_LITERAL_RE = re.compile(r'"[^"]*"|\\.|\[[^\]]*\]')


def _xlsx_value(value, fmt: str | None):
    """Значение ячейки .xlsx; число, принятое openpyxl за дату, — снова число.

    openpyxl считает датой любой формат с латинскими d/m/y/h/s, в том числе
    «0 m» (метры, пробел экранирован). Формат с цифровыми знаками 0/#/? — числовой: такую ячейку
    возвращаем порядковым числом, как она записана в файле.
    """
    if isinstance(value, (datetime, date, time, timedelta)) and fmt and (
        _DIGIT_PLACEHOLDER_RE.search(_FORMAT_LITERAL_RE.sub("", fmt))
    ):
        return to_excel(value)
    return value


def _xls_format(book, xf_index: int) -> str | None:
    """Формат ячейки .xls по индексу XF.

    Встроенные форматы (код < 164) xlrd отдаёт американской строкой
    («m/d/yy» для даты): берём код Excel, чтобы дата показывалась так, как
    принято в системе пользователя.
    """
    try:
        key = book.xf_list[xf_index].format_key
    except (IndexError, AttributeError):
        return None
    if key in BUILTIN_FORMATS:
        return _format(BUILTIN_FORMATS[key])
    fmt = book.format_map.get(key)
    return _format(fmt.format_str if fmt is not None else None)


def _build_grid(
    name: str,
    data: list[list],
    merged_ranges,
    error_cells: int = 0,
    formats: list[list] | None = None,
    hidden: bool = False,
) -> SheetGrid:
    """Обрезает данные до реальной области и собирает SheetGrid."""
    nrows, ncols = _real_area(data)
    trimmed = [row[:ncols] for row in data[:nrows]]
    clipped = [
        (rlo, min(rhi, nrows), clo, min(chi, ncols))
        for rlo, rhi, clo, chi in merged_ranges
        if rlo < nrows and clo < ncols
    ]
    trimmed_formats = [row[:ncols] for row in formats[:nrows]] if formats else None
    grid = SheetGrid(name, trimmed, clipped, trimmed_formats)
    grid.error_cells = error_cells
    grid.hidden = hidden
    return grid


def _map_open_error(exc: Exception, path: Path):
    """Превращает исключения движков чтения в пользовательские ошибки."""
    text = str(exc).lower()
    if "password" in text or "encrypt" in text:
        return EncryptedFileError(f"Файл защищён паролем: {path.name}")
    return CorruptedFileError(f"Не удалось прочитать файл «{path.name}»: {exc}")


def _load_xls(path: Path) -> Workbook:
    try:
        book = xlrd.open_workbook(str(path), formatting_info=True)
    except Exception:
        # formatting_info может мешать на экзотических файлах (xlrd бросает
        # не только XLRDError) — пробуем без него, без merged cells
        try:
            book = xlrd.open_workbook(str(path))
        except Exception as exc:
            raise _map_open_error(exc, path) from exc
    return _xls_grids(path, book)


def _xls_grids(path: Path, book) -> Workbook:
    sheets = []
    with_formats = bool(getattr(book, "xf_list", None))
    for sh in book.sheets():
        data = []
        formats = []
        errors = 0
        for r in range(sh.nrows):
            row = []
            row_formats = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                row_formats.append(
                    _xls_format(book, cell.xf_index)
                    if with_formats and cell.xf_index is not None else None
                )
                if cell.ctype == xlrd.XL_CELL_ERROR:
                    errors += 1
                    row.append(None)
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    row.append(None)
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    # xlrd отдаёт дату порядковым числом (43214.0) — в дату
                    try:
                        row.append(xlrd.xldate_as_datetime(cell.value, book.datemode))
                    except (ValueError, OverflowError):
                        row.append(cell.value)
                elif _is_empty(cell.value):
                    row.append(None)
                else:
                    row.append(cell.value)
            data.append(row)
            formats.append(row_formats)
        merged = getattr(sh, "merged_cells", []) or []
        hidden = getattr(sh, "visibility", 0) != 0
        sheets.append(_build_grid(sh.name, data, merged, errors, formats, hidden))
    return Workbook(path=path, sheets=sheets)


def _count_uncached_formulas(path: Path, book) -> dict[str, int]:
    """Лист -> число формул, у которых в файле нет сохранённого значения.

    Такие ячейки читаются пустыми (файл сохранён не из Excel или не пересчитан).
    Подсчёт вспомогательный: при любой ошибке возвращается пустой словарь.
    """
    import openpyxl

    counts: dict[str, int] = {}
    try:
        formulas = openpyxl.load_workbook(path, data_only=False, read_only=False)
        for ws in formulas.worksheets:
            values = book[ws.title]
            counts[ws.title] = sum(
                1
                for row in ws.iter_rows()
                for cell in row
                if cell.data_type == "f" and values.cell(cell.row, cell.column).value is None
            )
    except Exception:
        return {}
    return counts


def _load_xlsx(path: Path) -> Workbook:
    import openpyxl

    try:
        book = openpyxl.load_workbook(path, data_only=True, read_only=False)
    except (zipfile.BadZipFile, OSError) as exc:
        raise _map_open_error(exc, path) from exc
    except Exception as exc:
        raise _map_open_error(exc, path) from exc

    uncached = _count_uncached_formulas(path, book)
    sheets = []
    for ws in book.worksheets:
        data = []
        formats = []
        errors = 0
        for cells in ws.iter_rows():
            row = [_xlsx_value(cell.value, cell.number_format) for cell in cells]
            errors += sum(1 for v in row if isinstance(v, str) and v in _ERROR_VALUES)
            data.append([
                None if _is_empty(v) or v in _ERROR_VALUES else v
                for v in row
            ])
            formats.append([_format(cell.number_format) for cell in cells])
        merged = [
            (rng.min_row - 1, rng.max_row, rng.min_col - 1, rng.max_col)
            for rng in ws.merged_cells.ranges
        ]
        grid = _build_grid(
            ws.title, data, merged, errors, formats, ws.sheet_state != "visible"
        )
        grid.uncached_formulas = uncached.get(ws.title, 0)
        sheets.append(grid)
    return Workbook(path=path, sheets=sheets)


def load_workbook(path: Path) -> Workbook:
    """Читает книгу Excel в Workbook. Формат — по сигнатуре файла."""
    path = Path(path)
    if not path.is_file():
        raise PathNotFoundError(f"Файл не найден: {path}")
    try:
        with open(path, "rb") as fh:
            signature = fh.read(8)
    except OSError as exc:
        raise CorruptedFileError(f"Не удалось открыть файл «{path.name}»: {exc}") from exc

    if signature[:4] == _OLE2_SIGNATURE:
        return _load_xls(path)
    if signature[:2] == _ZIP_SIGNATURE:
        return _load_xlsx(path)
    raise NotExcelFileError(f"Файл «{path.name}» не является таблицей Excel (.xls или .xlsx).")
