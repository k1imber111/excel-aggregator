"""Извлечение набора ключей с листа по якорной координате."""

from __future__ import annotations

from dataclasses import dataclass, field

from .errors import EmptyKeySetError
from .models import Orientation, SheetGrid
from .normalize import normalize_key


@dataclass
class KeySet:
    """Набор ключей одного источника.

    keys — канонические строки в порядке появления, без дубликатов.
    duplicates — канонические строки повторных вхождений (для предупреждения).
    raw_positions — ключ -> 0-based индекс строки (VERTICAL) или столбца
    (HORIZONTAL) первого вхождения на листе.
    """

    keys: list[str] = field(default_factory=list)
    duplicates: list[str] = field(default_factory=list)
    raw_positions: dict[str, int] = field(default_factory=dict)


def extract_keys(
    grid: SheetGrid,
    row0: int,
    col0: int,
    orientation: Orientation,
) -> KeySet:
    """Читает ключи от якоря вниз (VERTICAL) или вправо (HORIZONTAL).

    Пустые ячейки пропускаются и не обрывают чтение; стоп — только по концу
    заполненной области. Дубликаты внутри источника: берётся первое вхождение,
    повторы попадают в KeySet.duplicates.
    """
    keyset = KeySet()
    seen: set[str] = set()

    if orientation is Orientation.VERTICAL:
        indices = range(row0, grid.nrows)
        values = (grid.value(i, col0) for i in indices)
    else:
        indices = range(col0, grid.ncols)
        values = (grid.value(row0, i) for i in indices)

    for index, raw in zip(indices, values):
        key = normalize_key(raw)
        if key is None:
            continue
        folded = key.casefold()
        if folded in seen:
            keyset.duplicates.append(key)
            continue
        seen.add(folded)
        keyset.keys.append(key)
        keyset.raw_positions[key] = index

    if not keyset.keys:
        raise EmptyKeySetError(
            f"По координате (строка {row0 + 1}, столбец {col0 + 1}) "
            "не найдено ни одного значения ключа."
        )
    return keyset
