"""Базовые модели данных core-слоя.

Все координаты внутри core — 0-based (row0, col0). Пользователь вводит
Excel-координаты (A6, счёт с 1); конвертация делается только здесь.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from .errors import AppError, CoordinateParseError

# Русские буквы-двойники латинских (похожие начертания)
_RU_TO_LAT = {
    "А": "A", "В": "B", "С": "C", "Е": "E", "Н": "H",
    "К": "K", "М": "M", "О": "O", "Р": "P", "Т": "T", "Х": "X",
}

_COORD_RE = re.compile(r"([A-ZА-ЯЁ]+)(\d+)")


def parse_coordinate(text: str) -> tuple[int, int]:
    """Excel-нотация ("A6", "А6", "AA10") -> (row0, col0), 0-based.

    Принимает латиницу и русские буквы-двойники, любой регистр,
    многобуквенные столбцы. Пробелы внутри игнорируются.
    Неверный ввод -> CoordinateParseError.
    """
    if text is None:
        raise CoordinateParseError("Координата не указана (пример: A6).")
    cleaned = re.sub(r"\s+", "", str(text)).upper()
    match = _COORD_RE.fullmatch(cleaned)
    if not match:
        raise CoordinateParseError(f"Не удалось распознать координату «{text}» (пример: A6).")
    letters, digits = match.groups()

    col = 0
    for ch in letters:
        lat = _RU_TO_LAT.get(ch, ch)
        if not ("A" <= lat <= "Z"):
            raise CoordinateParseError(
                f"Буква «{ch}» не похожа на столбец Excel в координате «{text}»."
            )
        col = col * 26 + (ord(lat) - ord("A") + 1)

    row = int(digits)
    if row < 1:
        raise CoordinateParseError(f"Номер строки в координате «{text}» должен быть не меньше 1.")
    return row - 1, col - 1


def col_letter(idx0: int) -> str:
    """0-based индекс столбца -> буквы Excel (0 -> "A", 26 -> "AA")."""
    if idx0 < 0:
        raise ValueError("Индекс столбца не может быть отрицательным.")
    letters = []
    n = idx0 + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


class Orientation(Enum):
    """Направление чтения ключей/характеристик на листе."""

    VERTICAL = "vertical"
    HORIZONTAL = "horizontal"


class SheetGrid:
    """Прямоугольная проекция листа: реальная область + развёрнутые merged cells.

    data — плотная матрица значений (уже обрезанная до реальной области,
    пустые ячейки — None). merged_ranges — список кортежей
    (rlo, rhi, clo, chi) с НЕвключёнными rhi/chi (конвенция xlrd):
    значение верхне-левой ячейки доступно по всем координатам диапазона.
    """

    def __init__(
        self,
        name: str,
        data: list[list],
        merged_ranges: list[tuple[int, int, int, int]] | None = None,
    ):
        self.name = name
        self.nrows = len(data)
        self.ncols = max((len(r) for r in data), default=0)
        # Выровнять строки до одной длины
        self._data = [list(r) + [None] * (self.ncols - len(r)) for r in data]
        self._merged: dict[tuple[int, int], tuple[int, int]] = {}
        # якорь диапазона -> (rlo, rhi, clo, chi)
        self._spans: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for rlo, rhi, clo, chi in merged_ranges or []:
            self._spans[(rlo, clo)] = (rlo, rhi, clo, chi)
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if (r, c) != (rlo, clo):
                        self._merged[(r, c)] = (rlo, clo)
        # Счётчики для отчёта (заполняет reader)
        self.error_cells = 0          # ячейки-ошибки (#REF! и т.п.), прочитаны пустыми
        self.uncached_formulas = 0    # формулы .xlsx без сохранённого значения

    def value(self, r: int, c: int):
        """Значение ячейки (0-based) с учётом развёрнутых merged cells."""
        src = self._merged.get((r, c))
        if src is not None:
            r, c = src
        if 0 <= r < self.nrows and 0 <= c < self.ncols:
            return self._data[r][c]
        return None

    def anchor(self, r: int, c: int) -> tuple[int, int]:
        """Верхне-левая ячейка объединения, в которое входит (r, c)."""
        return self._merged.get((r, c), (r, c))

    def raw(self, r: int, c: int):
        """Значение без разворота объединений: у «хвоста» объединения — None."""
        if (r, c) in self._merged:
            return None
        return self.value(r, c)

    def span(self, r: int, c: int) -> tuple[int, int, int, int]:
        """Диапазон (rlo, rhi, clo, chi) объединения с ячейкой (r, c); rhi/chi не включены."""
        ar, ac = self.anchor(r, c)
        return self._spans.get((ar, ac), (ar, ar + 1, ac, ac + 1))

    def transposed(self) -> "SheetGrid":
        """Тот же лист, повёрнутый: строки становятся столбцами.

        Нужен для таблиц, где названия объектов идут в строку: после поворота
        с ними работает обычная «вертикальная» логика.
        """
        grid = SheetGrid(
            self.name,
            [list(col) for col in zip(*self._data)],
            [(clo, chi, rlo, rhi) for rlo, rhi, clo, chi in self._spans.values()],
        )
        grid.error_cells = self.error_cells
        grid.uncached_formulas = self.uncached_formulas
        return grid

    def row(self, r: int) -> list:
        """Строка целиком (0-based)."""
        return [self.value(r, c) for c in range(self.ncols)]

    def col(self, c: int) -> list:
        """Столбец целиком (0-based)."""
        return [self.value(r, c) for r in range(self.nrows)]

    def preview_rows(self, n: int) -> list[list]:
        """Первые n непустых строк для превью."""
        result = []
        for r in range(self.nrows):
            row = self.row(r)
            if any(v is not None and str(v).strip() != "" for v in row):
                result.append(row)
                if len(result) >= n:
                    break
        return result


@dataclass
class SourceConfig:
    """Настройки одного источника данных (лист одного файла).

    anchor — (row0, col0) первой ячейки ключей, 0-based.
    key_index — номер столбца (VERTICAL) или строки (HORIZONTAL) ключа.
    char_columns — 0-based индексы столбцов (VERTICAL) или строк
    (HORIZONTAL) характеристик. column_titles — имена итоговых столбцов
    в том же порядке. column_paths — те же названия цепочкой уровней шапки
    («Полярные координаты», «Ап, град») для многоуровневой шапки итога;
    если пусто, берётся column_titles.
    """

    file_path: Path
    sheet_name: str
    anchor: tuple[int, int]
    orientation: Orientation
    key_index: int
    char_columns: list[int] = field(default_factory=list)
    column_titles: list[str] = field(default_factory=list)
    column_paths: list[list[str]] = field(default_factory=list)


@dataclass
class Workbook:
    """Прочитанная книга Excel: путь + список листов."""

    path: Path
    sheets: list[SheetGrid]

    def sheet(self, name: str) -> SheetGrid:
        """Лист по имени; имена с хвостовыми пробелами сравниваются мягко."""
        for sh in self.sheets:
            if sh.name == name:
                return sh
        stripped = name.strip()
        for sh in self.sheets:
            if sh.name.strip() == stripped:
                return sh
        available = ", ".join(repr(sh.name) for sh in self.sheets)
        raise AppError(f"Лист «{name}» не найден. Доступные листы: {available}")
