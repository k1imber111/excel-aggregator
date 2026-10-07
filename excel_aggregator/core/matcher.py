"""Автоматический поиск столбца/строки ключей в новом источнике."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .models import Orientation, SheetGrid
from .normalize import format_cell, normalize_key

# Минимальное число совпадений, чтобы кандидат не считался мусором
_MIN_MATCHES = 2
_SAMPLE_SIZE = 10


@dataclass
class Candidate:
    """Кандидат на столбец/строку ключей.

    index — 0-based индекс столбца (VERTICAL) или строки (HORIZONTAL).
    match_count — сколько РАЗЛИЧНЫХ ключей из base_keys встретилось
    (повторы одного значения, напр. «43» в каждой строке, не накручивают счёт).
    coverage — доля покрытых ключей (match_count / len(base_keys)).
    sample — первые 10 непустых значений в display-форме.
    first_row0 — 0-based индекс первой ячейки с совпадением (якорь).
    """

    index: int
    match_count: int
    coverage: float
    sample: list = field(default_factory=list)
    first_row0: int = 0


def find_key_candidates(
    grid: SheetGrid,
    base_keys: list[str],
    orientation: Orientation,
    top: int = 3,
) -> list[Candidate]:
    """Ищет столбцы (VERTICAL) или строки (HORIZONTAL) с ключами из base_keys.

    Кандидаты отсортированы по match_count по убыванию; кандидаты с
    match_count < 2 отсеиваются. Пустой результат — это нормально
    (возвращается [], решение принимает UI).
    """
    base = {normalize_key(k).casefold() for k in base_keys if normalize_key(k) is not None}
    candidates: list[Candidate] = []

    outer = range(grid.ncols) if orientation is Orientation.VERTICAL else range(grid.nrows)
    for index in outer:
        values = grid.col(index) if orientation is Orientation.VERTICAL else grid.row(index)
        matched: set[str] = set()   # различные совпавшие ключи (дубли не накручивают счёт)
        first_row0 = -1
        sample: list = []
        for i, raw in enumerate(values):
            key = normalize_key(raw)
            if key is None:
                continue
            if len(sample) < _SAMPLE_SIZE:
                sample.append(format_cell(raw))
            folded = key.casefold()
            if folded in base and folded not in matched:
                matched.add(folded)
                if first_row0 < 0:
                    first_row0 = i
        match_count = len(matched)
        if match_count >= _MIN_MATCHES:
            coverage = match_count / len(base) if base else 0.0
            candidates.append(Candidate(index, match_count, coverage, sample, first_row0))

    candidates.sort(key=lambda c: c.match_count, reverse=True)
    return candidates[:top]


# --- автоматический выбор столбца, по которому соединять таблицы ----------------

# Столбец названий одиночной таблицы: заполнен и значения в основном не повторяются
_SOLO_UNIQUENESS = 0.5
_SOLO_FILL = 0.8
# Дробные числа и даты — измерения, а не названия объектов
_MEASURE_RE = re.compile(r"-?\d+[.,]\d+|\d{4}-\d{2}-\d{2}.*")
_MEASURE_PENALTY = 0.3
_NUMERIC_PENALTY = 0.9
# Ряд 1…N — нумерация строк («№ п/п»): совпадает у любых двух таблиц
_ROW_NUMBER_PENALTY = 0.3
_SAME_TITLE_BONUS = 1.5   # одинаковый заголовок столбца в обеих таблицах
_POSITION_STEP = 0.05     # каждый столбец правее чуть менее вероятен
# Ниже этой оценки пересечение считается случайным, а не соединением:
# два общих значения в левых столбцах проходят, «1, 2, 3» в обоих файлах — нет
_MIN_JOIN_SCORE = 1.5


@dataclass
class _Column:
    """Столбец как кандидат в названия объектов."""

    index: int
    values: set[str]          # различные значения области данных
    uniqueness: float         # доля неповторяющихся значений
    title: str = ""

    @property
    def is_row_number(self) -> bool:
        """Значения — ровно 1, 2, …, N."""
        n = len(self.values)
        return n >= 3 and all(v.isdigit() for v in self.values) and (
            {int(v) for v in self.values} == set(range(1, n + 1))
        )

    @property
    def quality(self) -> float:
        """Насколько столбец похож на названия объектов (0..1)."""
        half = len(self.values) / 2
        if sum(1 for v in self.values if _MEASURE_RE.fullmatch(v)) > half:
            penalty = _MEASURE_PENALTY
        elif sum(1 for v in self.values if v.isdigit()) > half:
            penalty = _NUMERIC_PENALTY   # номера бывают названиями, но текст вероятнее
        else:
            penalty = 1.0
        return penalty / (1 + _POSITION_STEP * self.index)


def _columns(grid: SheetGrid, data_start: int, titles: dict[int, str] | None) -> list[_Column]:
    """Непустые столбцы области данных."""
    columns = []
    for c in range(grid.ncols):
        keys = (normalize_key(grid.value(r, c)) for r in range(data_start, grid.nrows))
        values = [k.casefold() for k in keys if k is not None]
        if values:
            distinct = set(values)
            columns.append(
                _Column(c, distinct, len(distinct) / len(values), (titles or {}).get(c, ""))
            )
    return columns


def _normal_title(title: str) -> str:
    """Заголовок без регистра, пробелов и знаков: «№ Препят-ствия» == «№ препятствия»."""
    return re.sub(r"[\W_]+", "", title.casefold())


def _pair_score(a: _Column, b: _Column) -> float:
    """Оценка пары столбцов как общего столбца названий двух таблиц.

    Основа — число общих значений × меньшая уникальность: высоты в футах
    совпадают у разных файлов почти полностью, но в столбце повторяются.
    Поправки отличают названия от измерений: штраф за дробные числа и даты,
    за нумерацию строк 1…N в обоих столбцах, за положение правее; бонус за
    одинаковый заголовок.
    """
    shared = len(a.values & b.values)
    if shared < _MIN_MATCHES:
        return 0.0
    score = shared * min(a.uniqueness, b.uniqueness) * a.quality * b.quality
    if a.is_row_number and b.is_row_number:
        score *= _ROW_NUMBER_PENALTY
    if a.title and _normal_title(a.title) == _normal_title(b.title):
        score *= _SAME_TITLE_BONUS
    return score


def _solo_column(columns: list[_Column]) -> int | None:
    """Столбец названий таблицы без пары: заполненный, с наименьшим числом повторов.

    При равной оценке берётся самый левый. По одной уникальности выигрывали бы
    измерения: у названий повторы бывают («Мачта 2» дважды), у дробных чисел —
    почти нет. Нумерация строк 1…N названием не считается.
    """
    if not columns:
        return None
    longest = max(len(col.values) for col in columns)
    scored = [
        (col.uniqueness * col.quality * (_ROW_NUMBER_PENALTY if col.is_row_number else 1.0),
         -col.index)
        for col in columns
        if col.uniqueness >= _SOLO_UNIQUENESS and len(col.values) >= _SOLO_FILL * longest
    ]
    return -max(scored)[1] if scored else None


def suggest_join(
    tables: list[tuple[SheetGrid, int]],
    titles: list[dict[int, str]] | None = None,
) -> list[int | None]:
    """Столбец названий объектов для каждой таблицы (grid, data_start).

    titles — заголовки столбцов каждой таблицы (столбец -> текст), если известны.

    Таблицы соединяются по столбцам, значения которых пересекаются между
    таблицами. Первая таблица с уверенным пересечением задаёт базу; остальные
    подключаются по убыванию оценки к уже набранным названиям. Таблица без
    уверенного пересечения получает свой «одиночный» столбец: соединится она
    или нет, покажет проверка — молча приписывать объектам чужие данные по
    случайному совпадению чисел нельзя.
    """
    profiles = [
        _columns(grid, start, titles[i] if titles else None)
        for i, (grid, start) in enumerate(tables)
    ]
    solo = [_solo_column(columns) for columns in profiles]

    def best_against(i: int, known: _Column) -> tuple[float, int | None]:
        scored = [(_pair_score(col, known), -col.index) for col in profiles[i]]
        score, neg_index = max(scored, default=(0.0, 0))
        return (score, -neg_index) if score >= _MIN_JOIN_SCORE else (0.0, None)

    # База: столбец первой таблицы с наибольшим суммарным пересечением с остальными
    anchor: _Column | None = None
    anchor_table = -1
    for i, columns in enumerate(profiles):
        totals = {
            col.index: sum(best_against(j, col)[0] for j in range(len(tables)) if j != i)
            for col in columns
        }
        best = max(totals, key=lambda c: (totals[c], -c), default=None)
        if best is not None and totals[best] > 0:
            anchor = next(col for col in columns if col.index == best)
            anchor_table = i
            break
    if anchor is None:
        return solo

    result = list(solo)
    result[anchor_table] = anchor.index
    known = _Column(0, set(anchor.values), 1.0, anchor.title)
    pending = set(range(len(tables))) - {anchor_table}
    while pending:
        scored = {i: best_against(i, known) for i in pending}
        i = max(scored, key=lambda k: (scored[k][0], -k))
        column = scored[i][1]
        if column is None:
            break  # у оставшихся таблиц уверенного пересечения нет — им остаётся solo
        result[i] = column
        known.values |= next(col for col in profiles[i] if col.index == column).values
        pending.remove(i)
    return result
