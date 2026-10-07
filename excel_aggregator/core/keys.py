"""Извлечение набора ключей с листа по якорной координате."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .errors import EmptyKeySetError
from .models import Orientation, SheetGrid
from .normalize import normalize_key

# Строки итогов и примечаний — не объекты
_SUMMARY_RE = re.compile(r"(итого|всего|примечани[ея]|total)\b", re.IGNORECASE)


@dataclass
class KeySet:
    """Набор ключей одного источника.

    keys — канонические строки в порядке появления, без дубликатов.
    duplicates — канонические строки повторных вхождений (для предупреждения),
    duplicate_rows — их 0-based позиции на листе, в том же порядке.
    raw_positions — ключ -> 0-based индекс строки (VERTICAL) или столбца
    (HORIZONTAL) первого вхождения на листе.
    skipped — (позиция, текст) строк-подписей, не ставших ключами.
    """

    keys: list[str] = field(default_factory=list)
    duplicates: list[str] = field(default_factory=list)
    duplicate_rows: list[int] = field(default_factory=list)
    raw_positions: dict[str, int] = field(default_factory=dict)
    skipped: list[tuple[int, str]] = field(default_factory=list)


def _is_caption(grid: SheetGrid, index: int, key_line: int, vertical: bool, key: str) -> bool:
    """Строка — подпись, а не объект?

    Итог или примечание («Итого», «Примечание: …»), либо текст из нескольких
    слов, рядом с которым в строке больше ничего нет («район 2»). Одиночные
    коды и номера без характеристик остаются объектами.
    """
    if _SUMMARY_RE.match(key):
        return True
    if " " not in key:
        return False
    width = grid.ncols if vertical else grid.nrows
    others = (
        grid.raw(index, i) if vertical else grid.raw(i, index)
        for i in range(width) if i != key_line
    )
    return width > 1 and all(normalize_key(v) is None for v in others)


def extract_keys(
    grid: SheetGrid,
    row0: int,
    col0: int,
    orientation: Orientation,
) -> KeySet:
    """Читает ключи от якоря вниз (VERTICAL) или вправо (HORIZONTAL).

    Пустые ячейки пропускаются и не обрывают чтение; стоп — только по концу
    заполненной области. Дубликаты внутри источника: берётся первое вхождение,
    повторы попадают в KeySet.duplicates. Строки-подписи (см. _is_caption)
    ключами не становятся и попадают в KeySet.skipped.
    """
    keyset = KeySet()
    seen: set[str] = set()
    vertical = orientation is Orientation.VERTICAL

    if vertical:
        indices = range(row0, grid.nrows)
        values = (grid.value(i, col0) for i in indices)
    else:
        indices = range(col0, grid.ncols)
        values = (grid.value(row0, i) for i in indices)
    key_line = col0 if vertical else row0

    for index, raw in zip(indices, values):
        key = normalize_key(raw)
        if key is None:
            continue
        if _is_caption(grid, index, key_line, vertical, key):
            keyset.skipped.append((index, key))
            continue
        folded = key.casefold()
        if folded in seen:
            keyset.duplicates.append(key)
            keyset.duplicate_rows.append(index)
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
