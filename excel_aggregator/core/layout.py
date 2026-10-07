"""Распознавание таблицы на листе: шапка, начало данных, названия столбцов.

Таблица читается сверху вниз (один объект — одна строка). Повёрнутый лист
сначала разворачивается через SheetGrid.transposed().
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .models import SheetGrid, col_letter
from .normalize import format_cell, normalize_key

# Строка нумерации граф: «1», «2», «1а», «15»
_NUMBERING_RE = re.compile(r"\d{1,3}[а-яa-z]?")
_NUMERIC_RE = re.compile(r"-?\d+([.,]\d+)?")

# Единицы измерения: не название графы, а уточнение к нему («Ап» + «град»)
_UNIT_VALUES = {
    "м", "метр", "метры", "км", "см", "мм",
    "фт", "фут", "футы",
    "°", "'", '"', "′", "″", "с", "град", "град.", "мин", "мин.", "сек", "сек.",
}
# Знак градуса, набранный нулём или буквой «о» (стоит слева от знака минут)
_DEGREE_MARKS = {"0", "о", "o"}
_MINUTE_MARKS = {"'", "′"}

_MERGED_SHARE = 0.3   # доля объединённых ячеек, выше которой строка — шапка
_DENSE_SHARE = 0.5    # строка данных заполнена не меньше чем на половину самой широкой
_SIMILARITY = 0.8     # совпадение набора заполненных столбцов у соседних строк данных
_LOOKAHEAD = 2        # со сколькими следующими строками сверяем
_SCAN_ROWS = 200      # в скольких верхних непустых строках ищем нумерацию граф
_PROFILE_ROWS = 50    # по скольким строкам определяем числовые столбцы
_NUMERIC_COLUMN = 0.6


@dataclass
class TableLayout:
    """Где на листе находится таблица.

    title — подпись над таблицей (или ""). header_rows — 0-based строки шапки
    без подписи, строки нумерации и строк-разделителей. data_start — 0-based
    первая строка данных.
    """

    title: str = ""
    header_rows: list[int] = field(default_factory=list)
    numbering_row: int | None = None
    data_start: int = 0


def _text(value) -> str:
    """Текст ячейки шапки: число без «.0», пробелы и переносы — одним пробелом."""
    return " ".join(str(format_cell(value)).split())


def _filled(grid: SheetGrid, r: int) -> list[int]:
    """Столбцы строки с собственным (не унаследованным от объединения) значением."""
    return [c for c in range(grid.ncols) if normalize_key(grid.raw(r, c)) is not None]


def _is_numeric(value) -> bool:
    key = normalize_key(value)
    return key is not None and _NUMERIC_RE.fullmatch(key) is not None


def _is_numbering(grid: SheetGrid, r: int) -> bool:
    """Строка нумерации граф: 1, 1а, 2, 3… — с нуля или единицы и по возрастанию.

    Номером должна быть каждая ячейка строки, включая первую: строка данных
    «c | 1 | 2 | 3» — не нумерация. Шаг 0 допустим изредка («1», «1а»), но
    номера обязаны расти: строка из одних нулей или единиц — тоже не нумерация.
    """
    values = [normalize_key(grid.raw(r, c)).casefold() for c in _filled(grid, r)]
    if len(values) < 3 or not all(_NUMBERING_RE.fullmatch(v) for v in values):
        return False
    ints = [int(re.match(r"\d+", v).group()) for v in values]
    steps = [b - a for a, b in zip(ints, ints[1:])]
    return (
        ints[0] in (0, 1)
        and all(step in (0, 1) for step in steps)
        and steps.count(1) >= 0.6 * len(steps)
    )


def _merged_share(grid: SheetGrid, r: int) -> float:
    """Доля ячеек строки, входящих в объединения."""
    merged = sum(
        1 for c in range(grid.ncols)
        if (span := grid.span(r, c))[1] - span[0] > 1 or span[3] - span[2] > 1
    )
    return merged / max(1, grid.ncols)


def _similar(grid: SheetGrid, r1: int, r2: int) -> float:
    """Насколько заполненные столбцы более пустой строки входят в более полную (0..1).

    Пустые ячейки в строке данных — норма (необязательные графы), поэтому
    сравниваем по меньшему набору, а не по объединению.
    """
    a, b = set(_filled(grid, r1)), set(_filled(grid, r2))
    return len(a & b) / min(len(a), len(b)) if a and b else 0.0


def _is_unit_row(grid: SheetGrid, r: int) -> bool:
    """Строка единиц измерения («м», «фут», «°», «'»…) — часть шапки."""
    cells = _filled(grid, r)
    units = sum(1 for c in cells if _is_unit(grid, r, c, _text(grid.raw(r, c))))
    return bool(cells) and units / len(cells) >= _NUMERIC_COLUMN


def _is_lone_cell_row(grid: SheetGrid, r: int) -> bool:
    """В строке одна ячейка — подпись или разделитель («район 2»), не данные."""
    return len(_filled(grid, r)) == 1


def _numeric_columns(grid: SheetGrid, rows: list[int]) -> set[int]:
    """Столбцы, где в строках rows преобладают числа."""
    result = set()
    for c in range(grid.ncols):
        values = [grid.raw(r, c) for r in rows if normalize_key(grid.raw(r, c)) is not None]
        if values and sum(map(_is_numeric, values)) / len(values) >= _NUMERIC_COLUMN:
            result.add(c)
    return result


def _is_header_like(grid: SheetGrid, r: int, numeric_cols: set[int]) -> bool:
    """Плотная строка — ещё шапка: в числовых столбцах у неё текст, а не числа."""
    cells = [c for c in _filled(grid, r) if c in numeric_cols]
    return bool(cells) and not any(_is_numeric(grid.raw(r, c)) for c in cells)


def detect_layout(grid: SheetGrid) -> TableLayout | None:
    """Находит таблицу на листе. None — таблицы нет или она не распознана.

    1. Строка нумерации граф, если есть: данные начинаются под ней.
    2. Первая плотная строка без объединений, похожая на две следующие
       непустые строки, — начало блока таблицы.
    3. Если над блоком нет шапки, первые его строки с текстом на месте чисел
       считаются шапкой (обычная таблица с заголовком в одну строку).
    """
    fills = [len(_filled(grid, r)) for r in range(grid.nrows)]
    widest = max(fills, default=0)
    if widest < 2:
        return None
    nonempty = [r for r, n in enumerate(fills) if n]
    # Нумерация граф стоит под шапкой, поэтому первую непустую строку не проверяем
    numbering = next((r for r in nonempty[1:_SCAN_ROWS] if _is_numbering(grid, r)), None)

    def dense(r: int) -> bool:
        return fills[r] >= _DENSE_SHARE * widest

    for i, r in enumerate(nonempty):
        if numbering is not None and r <= numbering:
            continue
        if not dense(r) or _merged_share(grid, r) > _MERGED_SHARE:
            continue
        following = nonempty[i + 1:i + 1 + _LOOKAHEAD]
        if not all(dense(n) and _similar(grid, r, n) >= _SIMILARITY for n in following):
            continue

        block = nonempty[i:i + _PROFILE_ROWS]
        numeric_cols = _numeric_columns(grid, block[1:] or block)

        # Вверх от блока: разреженные строки с числами — тоже данные. Без этого
        # таблица с неполным началом теряла бы первые объекты молча.
        first = i
        for j in range(i - 1, -1, -1):
            above = nonempty[j]
            if (numbering is not None and above <= numbering) or (
                _merged_share(grid, above) > _MERGED_SHARE
            ):
                break
            if _is_lone_cell_row(grid, above):
                continue  # подпись между строками данных не прерывает таблицу
            has_number = any(
                _is_numeric(grid.raw(above, c)) for c in _filled(grid, above) if c in numeric_cols
            )
            if _is_unit_row(grid, above) or not has_number:
                break
            first = j
        if first < i:
            return layout_from(grid, nonempty[first], numbering)

        start = i
        if numeric_cols:
            # Строки блока с текстом на месте чисел — шапка без объединений
            while start < len(nonempty) - 1 and _is_header_like(
                grid, nonempty[start], numeric_cols
            ):
                start += 1
        elif i + 1 < len(nonempty) and not layout_from(grid, r, numbering).header_rows:
            start = i + 1  # таблица из одного текста без шапки сверху: первая строка — заголовок
        return layout_from(grid, nonempty[start], numbering)
    return None


def layout_from(
    grid: SheetGrid, data_start: int, numbering: int | None = None
) -> TableLayout:
    """Раскладка по известной первой строке данных: всё, что выше, — шапка."""
    layout = TableLayout(data_start=data_start)
    for r in range(min(data_start, grid.nrows)):
        cells = _filled(grid, r)
        if not cells:
            continue
        if r == numbering or (numbering is None and _is_numbering(grid, r)):
            layout.numbering_row = r
            continue
        if len(cells) == 1:
            rlo, _, clo, chi = grid.span(r, cells[0])
            wide = chi - clo >= max(2, grid.ncols / 2)
            # Одиночная ячейка вне объединений — подпись или разделитель («район 2»)
            lone = chi - clo == 1 and not any(
                grid.span(r, c)[0] < r for c in range(grid.ncols)
            )
            if wide or lone:
                if not layout.title and not layout.header_rows:
                    layout.title = _text(grid.raw(rlo, clo))
                continue
        layout.header_rows.append(r)
    return layout


def _is_unit(grid: SheetGrid, r: int, c: int, text: str) -> str:
    """Единица измерения, если ячейка — единица; иначе ""."""
    if text.casefold() in _UNIT_VALUES:
        return text
    if text.casefold() in _DEGREE_MARKS and c + 1 < grid.ncols:
        right = grid.raw(r, c + 1)
        if right is not None and _text(right) in _MINUTE_MARKS:
            return "°"
    return ""


def column_paths(grid: SheetGrid, layout: TableLayout) -> dict[int, list[str]]:
    """Название каждого столбца цепочкой уровней шапки, сверху вниз.

    «Прямоугольные координаты» / «ИВПП 06/24» / «X». Уровни с одинаковым
    охватом столбцов — одна фраза и склеиваются пробелом; единица измерения
    дописывается к последнему уровню («Ап, град»). Столбец, у которого своя
    только единица, берёт название у соседа слева (необъединённая группа:
    «Высота» над «м» и «фут»). Столбец без шапки -> «Столбец B».
    """
    # столбец -> [(строка, (clo, chi), текст, единица)]
    parts: dict[int, list[tuple[int, tuple[int, int], str, str]]] = {}
    for c in range(grid.ncols):
        seen: set[tuple[int, int]] = set()
        column: list[tuple[int, tuple[int, int], str, str]] = []
        for r in layout.header_rows:
            anchor = grid.anchor(r, c)
            if anchor in seen:
                continue
            seen.add(anchor)
            value = grid.raw(*anchor)
            if normalize_key(value) is None:
                continue
            text = _text(value)
            _, _, clo, chi = grid.span(*anchor)
            column.append((anchor[0], (clo, chi), text, _is_unit(grid, *anchor, text)))

        names = [p for p in column if not p[3]]
        own_name = any(p[1] == (c, c + 1) for p in names)
        if not own_name and any(p[3] for p in column) and c > 0:
            # Необъединённая группа: недостающие уровни — у соседа слева
            have = {p[0] for p in column}
            borrowed = [
                p for p in parts[c - 1]
                if not p[3] and p[0] not in have and p[1] == (c - 1, c)
            ]
            column = sorted(column + [(r, (c, c + 1), t, u) for r, _, t, u in borrowed])
        parts[c] = column

    result: dict[int, list[str]] = {}
    for c, column in parts.items():
        levels: list[tuple[tuple[int, int], str]] = []
        unit = ""
        for _, cols, text, is_unit in column:
            if is_unit:
                unit = is_unit
            elif levels and levels[-1][0] == cols:
                levels[-1] = (cols, f"{levels[-1][1]} {text}")  # одна фраза в две строки
            elif not levels or levels[-1][1] != text:
                levels.append((cols, text))
        path = [text for _, text in levels]
        if unit:
            if path:
                path[-1] = f"{path[-1]}{' ' if path[-1].endswith(',') else ', '}{unit}"
            else:
                path = [unit]
        result[c] = path or [f"Столбец {col_letter(c)}"]
    return result
