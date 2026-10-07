"""Сборка итоговой таблицы из нескольких источников.

Итоговый набор ключей — объединение ключей всех источников: сначала ключи
первого источника в их порядке, затем новые ключи последующих источников
в порядке появления. Ключ, отсутствующий в источнике, -> пустая ячейка (None).
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Callable

from .keys import KeySet
from .models import Orientation, SheetGrid, SourceConfig


class Aggregator:
    """Накапливает источники и собирает итоговую таблицу.

    match — форма ключа для сравнения между источниками: по умолчанию без
    учёта регистра; fold_lookalikes — ещё и без различия латиницы/кириллицы.
    """

    def __init__(self, key_title: str = "Название", match: Callable[[str], str] = str.casefold):
        self.key_title = key_title
        self._match = match
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
                form = self._match(key)
                if form not in seen:
                    seen.add(form)
                    keys.append(key)
        return keys

    def build_table(self) -> tuple[list[str], list[list]]:
        """Собирает (headers, rows) итоговой таблицы."""
        headers = [self.key_title]
        lookups: list[dict[str, int]] = []
        for keyset, config, _ in self._sources:
            headers.extend(config.column_titles)
            lookup: dict[str, int] = {}
            for key in keyset.keys:
                lookup.setdefault(self._match(key), keyset.raw_positions[key])
            lookups.append(lookup)

        rows: list[list] = []
        for key in self.union_keys():
            row = [key]
            for (_, config, grid), lookup in zip(self._sources, lookups):
                pos = lookup.get(self._match(key))
                for char in config.char_columns:
                    if pos is None:
                        row.append(None)
                    elif config.orientation is Orientation.VERTICAL:
                        row.append(grid.value(pos, char))
                    else:
                        row.append(grid.value(char, pos))
            rows.append(row)
        return headers, rows

    def source_labels(self) -> list[str]:
        """Подпись каждого источника: имя файла; лист — если из файла их несколько."""
        per_file = Counter(Path(config.file_path) for _, config, _ in self._sources)
        labels: list[str] = []
        for _, config, _ in self._sources:
            path = Path(config.file_path)
            sheet = config.sheet_name.strip()
            label = f"{path.stem} — {sheet}" if per_file[path] > 1 else path.stem
            # Одноимённые файлы из разных папок: подписи должны различаться,
            # иначе их блоки столбцов слились бы в шапке в один
            n = sum(1 for used in labels if used == label or used.startswith(f"{label} ("))
            labels.append(f"{label} ({n + 1})" if n else label)
        return labels

    def header_paths(self) -> list[list[str]]:
        """Шапка итога: на каждый столбец — цепочка [источник, группа…, название].

        Первый столбец (названия объектов) источника не имеет: [key_title].
        """
        paths = [[self.key_title]]
        for label, (_, config, _) in zip(self.source_labels(), self._sources):
            columns = config.column_paths or [[title] for title in config.column_titles]
            paths.extend([label, *path] for path in columns)
        return paths

    def coverage_stats(self) -> list[tuple[str, int, int]]:
        """(имя источника, сколько ключей покрыто, всего ключей в объединении)."""
        keys = self.union_keys()
        stats = []
        for keyset, config, _ in self._sources:
            forms = {self._match(key) for key in keyset.keys}
            covered = sum(1 for key in keys if self._match(key) in forms)
            name = f"{Path(config.file_path).name} / {config.sheet_name}"
            stats.append((name, covered, len(keys)))
        return stats

    def duplicate_warnings(self) -> list[str]:
        """Сообщения о дубликатах ключей внутри источников (для UI)."""
        warnings = []
        for keyset, config, _ in self._sources:
            if keyset.duplicates:
                name = f"{Path(config.file_path).name} / {config.sheet_name}"
                examples = ", ".join(
                    f"{key} (строка {row + 1})"
                    for key, row in zip(keyset.duplicates[:5], keyset.duplicate_rows)
                ) or ", ".join(keyset.duplicates[:5])
                warnings.append(
                    f"{name}: дубликаты ключей ({len(keyset.duplicates)} шт., "
                    f"использовано первое вхождение): {examples}"
                )
        return warnings
