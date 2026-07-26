"""Сборка итоговой таблицы из нескольких источников.

Итоговый набор ключей — объединение ключей всех источников: сначала ключи
первого источника в их порядке, затем новые ключи последующих источников
в порядке появления. Ключ, отсутствующий в источнике, -> пустая ячейка (None).
"""

from __future__ import annotations

from pathlib import Path

from .keys import KeySet
from .models import Orientation, SheetGrid, SourceConfig


class Aggregator:
    """Накапливает источники и собирает итоговую таблицу."""

    def __init__(self, key_title: str = "Название"):
        self.key_title = key_title
        self._sources: list[tuple[KeySet, SourceConfig, SheetGrid]] = []

    def add_source(self, keyset: KeySet, config: SourceConfig, grid: SheetGrid) -> None:
        """Добавляет источник: его ключи, конфигурацию и сетку листа."""
        self._sources.append((keyset, config, grid))

    def union_keys(self) -> list[str]:
        """Объединение ключей всех источников в утверждённом порядке."""
        keys: list[str] = []
        seen: set[str] = set()
        for keyset, _, _ in self._sources:
            for key in keyset.keys:
                folded = key.casefold()
                if folded not in seen:
                    seen.add(folded)
                    keys.append(key)
        return keys

    def build_table(self) -> tuple[list[str], list[list]]:
        """Собирает (headers, rows) итоговой таблицы."""
        headers = [self.key_title]
        lookups: list[dict[str, int]] = []
        for keyset, config, _ in self._sources:
            headers.extend(config.column_titles)
            lookups.append({key.casefold(): keyset.raw_positions[key] for key in keyset.keys})

        rows: list[list] = []
        for key in self.union_keys():
            row = [key]
            for (_, config, grid), lookup in zip(self._sources, lookups):
                pos = lookup.get(key.casefold())
                for char in config.char_columns:
                    if pos is None:
                        row.append(None)
                    elif config.orientation is Orientation.VERTICAL:
                        row.append(grid.value(pos, char))
                    else:
                        row.append(grid.value(char, pos))
            rows.append(row)
        return headers, rows

    def coverage_stats(self) -> list[tuple[str, int, int]]:
        """(имя источника, сколько ключей покрыто, всего ключей в объединении)."""
        keys = self.union_keys()
        stats = []
        for keyset, config, _ in self._sources:
            folded = {key.casefold() for key in keyset.keys}
            covered = sum(1 for key in keys if key.casefold() in folded)
            name = f"{Path(config.file_path).name} / {config.sheet_name}"
            stats.append((name, covered, len(keys)))
        return stats

    def duplicate_warnings(self) -> list[str]:
        """Сообщения о дубликатах ключей внутри источников (для UI)."""
        warnings = []
        for keyset, config, _ in self._sources:
            if keyset.duplicates:
                name = f"{Path(config.file_path).name} / {config.sheet_name}"
                examples = ", ".join(keyset.duplicates[:5])
                warnings.append(
                    f"{name}: дубликаты ключей ({len(keyset.duplicates)} шт., "
                    f"использовано первое вхождение): {examples}"
                )
        return warnings
