"""Автоматический поиск столбца/строки ключей в новом источнике."""

from __future__ import annotations

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
