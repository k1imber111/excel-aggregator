"""Запись итоговой таблицы в форматированный .xlsx и проверка записанного."""

from __future__ import annotations

import math
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import Workbook as _XlsxWorkbook
from openpyxl import load_workbook as _load_xlsx
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

from .errors import OutputWriteError

# Число с ведущим нулём в целой части: "040", "04.61", "-05" — сохраняем как текст
_LEADING_ZERO_RE = re.compile(r"-?0\d")
# Название объекта из одних цифр («22») — пишем числом, чтобы фильтр сортировал 1, 2, 10
_INT_KEY_RE = re.compile(r"-?(0|[1-9]\d{0,14})")

_KEY_FILL = PatternFill("solid", fgColor="1F4E78")
# Источники чередуют заливку, чтобы блоки столбцов разных файлов различались
_SOURCE_FILLS = (PatternFill("solid", fgColor="2E75B6"), _KEY_FILL)
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
# Подпись широкого блока и название таблицы — слева: по центру блока в 15
# столбцов она оказалась бы за краем экрана
_WIDE_ALIGNMENT = Alignment(horizontal="left", vertical="center", indent=1)
_WIDE_GROUP = 6       # группа шире стольких столбцов подписывается слева
_FLOAT_DISPLAY = 12   # сколько знаков дробного числа учитываем в ширине столбца
_CELL_ALIGNMENT = Alignment(vertical="center", wrap_text=True)
_THIN_SIDE = Side(style="thin")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

_MIN_WIDTH = 10
_MAX_WIDTH = 50
_DEFAULT_ROW_HEIGHT = 15
_MAX_ROW_HEIGHT = 409
_TITLE_ROW_HEIGHT = 30

_AS_IS = (bool, int, float, datetime, date, time, timedelta)


def _clean(text) -> str:
    """Текст, допустимый в .xlsx: без управляющих символов, переводы строк — «\\n»."""
    return ILLEGAL_CHARACTERS_RE.sub("", str(text)).replace("\r\n", "\n").replace("\r", "\n")


def _cell_value(value, is_key: bool = False):
    """Значение, которое попадёт в ячейку: числа и даты — как есть, остальное — текст."""
    if value is None or isinstance(value, _AS_IS):
        return value
    text = _clean(value)
    if is_key and _INT_KEY_RE.fullmatch(text):
        return int(text)
    return text


def _display_len(v) -> int:
    """Длина самой длинной строки значения при показе (для автоподбора ширины)."""
    if v is None:
        return 0
    if isinstance(v, float):
        if v.is_integer():
            v = int(v)
        else:
            # Excel сам округляет показ дробного числа под ширину столбца;
            # 17 знаков точности не должны раздувать столбец
            return min(len(str(v)), _FLOAT_DISPLAY)
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


def _same(expected, actual) -> bool:
    """Совпадает ли прочитанное из файла значение с записанным."""
    if expected is None or expected == "":
        return actual is None or actual == ""
    if isinstance(expected, datetime):
        return isinstance(actual, datetime) and abs((actual - expected).total_seconds()) < 0.001
    if isinstance(expected, date):
        return (actual.date() if isinstance(actual, datetime) else actual) == expected
    if isinstance(expected, (time, timedelta)):
        return actual is not None  # Excel хранит время долей суток — сверяем наличие
    if isinstance(expected, float) and isinstance(actual, (int, float)):
        # .xlsx хранит 15–16 значащих цифр: 17-я цифра исходного числа теряется
        return math.isclose(actual, expected, rel_tol=1e-12, abs_tol=0.0)
    return actual == expected


def _verify(path: Path, rows: list[list], data_row: int, ncols: int) -> int:
    """Перечитывает записанный файл и сверяет данные ячейка в ячейку.

    Возвращает число сверенных ячеек; расхождение -> OutputWriteError.
    """
    try:
        book = _load_xlsx(path, read_only=True)
    except Exception as exc:
        raise OutputWriteError(
            f"Файл «{path}» записан, но не открывается для проверки: {exc}"
        ) from exc
    try:
        written = list(book.active.iter_rows(
            min_row=data_row, max_row=data_row + len(rows) - 1, max_col=ncols, values_only=True
        )) if rows else []
    finally:
        book.close()
    if len(written) != len(rows):
        raise OutputWriteError(
            f"Проверка записи не пройдена: в файле {len(written)} строк данных, "
            f"ожидалось {len(rows)}."
        )
    checked = 0
    for r, (expected_row, actual_row) in enumerate(zip(rows, written)):
        for c, value in enumerate(expected_row):
            expected = _cell_value(value, c == 0)
            actual = actual_row[c] if c < len(actual_row) else None
            if not _same(expected, actual):
                raise OutputWriteError(
                    "Проверка записи не пройдена: ячейка "
                    f"{get_column_letter(c + 1)}{data_row + r} — в файле {actual!r}, "
                    f"ожидалось {expected!r}."
                )
            checked += 1
    return checked


def export(
    path: Path,
    headers: list[str],
    rows: list[list],
    key_title: str = "Название",
    header_paths: list[list[str]] | None = None,
    title: str | None = None,
) -> int:
    """Пишет итоговую таблицу в .xlsx с оформлением и проверяет записанное.

    header_paths — шапка цепочками уровней на каждый столбец:
    [источник, группа…, название]; одинаковые соседние уровни объединяются.
    Без header_paths шапка — одна строка из headers. title — строка названия
    таблицы над шапкой.

    - шапка: заливка (блоки разных источников чередуют цвет), белый жирный
      текст, перенос, центрирование;
    - данные: тонкие границы, вертикальное выравнивание по центру;
    - ширина столбцов по данным и нижнему уровню шапки (мин 10, макс 50);
    - закреплены шапка и первый столбец, автофильтр на нижней строке шапки;
    - числа и даты пишутся как есть; строки с ведущими нулями — текстом ("@");
    - управляющие символы, недопустимые в .xlsx, из текста удаляются;
    - печать: альбомная, по ширине страницы, шапка на каждом листе.

    Возвращает число ячеек данных, сверенных с файлом после записи.
    """
    path = Path(path)
    banded = header_paths is not None
    paths = header_paths or [[h] for h in (headers or [key_title])]
    paths = [[_clean(level) for level in p] or [""] for p in paths]
    ncols = len(paths)
    depth = max(len(p) for p in paths)
    top = 2 if title else 1            # первая строка шапки
    leaf_row = top + depth - 1         # нижняя строка шапки — на ней фильтр
    data_row = leaf_row + 1
    last_col = get_column_letter(ncols)

    wb = _XlsxWorkbook()
    ws = wb.active
    ws.title = "Итог"

    if title:
        # Без объединения: текст сам продолжается вправо по пустым ячейкам строки
        cell = ws.cell(row=1, column=1, value=_clean(title))
        cell.font = _TITLE_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = _TITLE_ROW_HEIGHT

    # --- данные ---------------------------------------------------------
    values = [[_cell_value(v, c == 0) for c, v in enumerate(row)] for row in rows]
    for r, row in enumerate(values, start=data_row):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            if isinstance(value, str):
                cell.data_type = "s"  # текст, начинающийся с «=», — не формула
                if _LEADING_ZERO_RE.match(value):
                    cell.number_format = "@"  # не терять ведущие нули
            cell.border = _THIN_BORDER
            cell.alignment = _CELL_ALIGNMENT

    widths: list[float] = []
    for j, column_path in enumerate(paths):
        longest_word = max((len(word) for word in column_path[-1].split()), default=0)
        longest_value = max((_display_len(row[j]) for row in values if j < len(row)), default=0)
        width = min(_MAX_WIDTH, max(_MIN_WIDTH, max(longest_word, longest_value) + 2))
        ws.column_dimensions[get_column_letter(j + 1)].width = width
        widths.append(width)

    for r, row in enumerate(values, start=data_row):
        lines = max((_display_lines(v, widths[c]) for c, v in enumerate(row)), default=1)
        ws.row_dimensions[r].height = min(_MAX_ROW_HEIGHT, lines * _DEFAULT_ROW_HEIGHT)

    # --- шапка ------------------------------------------------------------
    fills = []
    band, previous = -1, None
    for j, column_path in enumerate(paths):
        if banded and j > 0:
            if column_path[0] != previous:
                band, previous = band + 1, column_path[0]
            fills.append(_SOURCE_FILLS[band % len(_SOURCE_FILLS)])
        else:
            fills.append(_KEY_FILL)
    for k in range(depth):
        for j in range(ncols):
            cell = ws.cell(row=top + k, column=j + 1)
            cell.fill = fills[j]
            cell.font = _HEADER_FONT
            cell.alignment = _HEADER_ALIGNMENT
            cell.border = _THIN_BORDER

    heights = [_DEFAULT_ROW_HEIGHT] * depth
    tall: list[tuple[int, int]] = []   # (уровень, нужная высота) ячеек на несколько строк
    for k in range(depth):
        j = 0
        while j < ncols:
            column_path = paths[j]
            if k >= len(column_path):
                j += 1     # место занято названием столбца, растянутым вниз
                continue
            is_leaf = k == len(column_path) - 1
            end = j
            if not is_leaf:
                # Группа: соседние столбцы с той же цепочкой до этого уровня
                while (
                    end + 1 < ncols
                    and len(paths[end + 1]) - 1 > k
                    and paths[end + 1][:k + 1] == column_path[:k + 1]
                ):
                    end += 1
            cell = ws.cell(row=top + k, column=j + 1, value=column_path[k])
            if end - j >= _WIDE_GROUP:
                cell.alignment = _WIDE_ALIGNMENT
            last_row = leaf_row if is_leaf else top + k
            if end > j or last_row > top + k:
                ws.merge_cells(
                    start_row=top + k, start_column=j + 1, end_row=last_row, end_column=end + 1
                )
            need = _display_lines(column_path[k], sum(widths[j:end + 1]) - 1) * _DEFAULT_ROW_HEIGHT
            if last_row > top + k:
                tall.append((k, need))
            else:
                heights[k] = max(heights[k], need)
            j = end + 1
    for k, need in tall:
        heights[-1] += max(0, need - sum(heights[k:]))
    for k, height in enumerate(heights):
        ws.row_dimensions[top + k].height = min(_MAX_ROW_HEIGHT, height)

    ws.freeze_panes = ws.cell(row=data_row, column=2 if ncols > 1 else 1)
    ws.auto_filter.ref = f"A{leaf_row}:{last_col}{max(leaf_row, data_row + len(rows) - 1)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"1:{leaf_row}"

    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)
    except PermissionError as exc:
        raise OutputWriteError(
            f"Не удалось записать файл «{path}»: он открыт в Excel (закройте его "
            "и повторите) или в эту папку нет прав на запись."
        ) from exc
    except OSError as exc:
        raise OutputWriteError(f"Не удалось записать файл «{path}»: {exc}") from exc

    return _verify(path, rows, data_row, ncols)
