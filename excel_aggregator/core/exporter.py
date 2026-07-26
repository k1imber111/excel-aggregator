"""Запись итоговой таблицы в форматированный .xlsx."""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook as _XlsxWorkbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .errors import OutputWriteError

# Число с ведущим нулём в целой части: "040", "04.61", "-05" — сохраняем как текст
_LEADING_ZERO_RE = re.compile(r"-?0\d")

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
_CELL_ALIGNMENT = Alignment(vertical="center", wrap_text=True)
_THIN_SIDE = Side(style="thin")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

_MIN_WIDTH = 10
_MAX_WIDTH = 50
_DEFAULT_ROW_HEIGHT = 15
_MAX_ROW_HEIGHT = 409


def _display_len(v) -> int:
    """Длина самой длинной строки значения при показе (для автоподбора ширины)."""
    if v is None:
        return 0
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return max(len(line) for line in str(v).split("\n"))


def _display_lines(v, width: float) -> int:
    """Примерное число строк после переноса текста в Excel."""
    if v is None:
        return 1
    usable = max(1, int(width) - 2)
    lines = 0
    for line in str(v).split("\n"):
        length = len(line)
        lines += max(1, (length + usable - 1) // usable)
    return lines


def export(path: Path, headers: list[str], rows: list[list], key_title: str = "Название") -> None:
    """Пишет итоговую таблицу в .xlsx с оформлением.

    - шапка: тёмно-синяя заливка, белый жирный текст, перенос, центрирование;
    - данные: тонкие границы, вертикальное выравнивание по центру;
    - автоподбор ширины столбцов (мин 10, макс 50);
    - freeze_panes="A2", автофильтр, без скрытых строк/столбцов;
    - числа пишутся числами; строки с ведущими нулями — текстом ("@").
    """
    path = Path(path)
    if not headers:
        headers = [key_title]
    if path.parent and not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)

    wb = _XlsxWorkbook()
    ws = wb.active
    ws.title = "Итог"

    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _THIN_BORDER

    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c)
            if isinstance(value, bool) or value is None:
                cell.value = value
            elif isinstance(value, (int, float)):
                cell.value = value  # настоящие числа остаются числами
            else:
                text = str(value)
                cell.value = text
                if _LEADING_ZERO_RE.match(text):
                    cell.number_format = "@"  # не терять ведущие нули
            cell.border = _THIN_BORDER
            cell.alignment = _CELL_ALIGNMENT

    widths: dict[int, float] = {}
    for c in range(1, len(headers) + 1):
        max_len = _display_len(ws.cell(row=1, column=c).value)
        for r in range(2, len(rows) + 2):
            max_len = max(max_len, _display_len(ws.cell(row=r, column=c).value))
        width = min(_MAX_WIDTH, max(_MIN_WIDTH, max_len + 2))
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = width
        widths[c] = width

    for r in range(1, len(rows) + 2):
        max_lines = max(
            _display_lines(ws.cell(row=r, column=c).value, widths[c])
            for c in range(1, len(headers) + 1)
        )
        ws.row_dimensions[r].height = min(
            _MAX_ROW_HEIGHT, max(_DEFAULT_ROW_HEIGHT, max_lines * _DEFAULT_ROW_HEIGHT)
        )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    try:
        wb.save(path)
    except PermissionError as exc:
        raise OutputWriteError(
            f"Не удалось записать файл «{path}»: он открыт в Excel — закройте его и повторите."
        ) from exc
    except OSError as exc:
        raise OutputWriteError(f"Не удалось записать файл «{path}»: {exc}") from exc
