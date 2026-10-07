"""Сценарий приложения: файлы → проверка найденного → столбцы → сборка.

Программа сама находит на листах таблицы, их шапки и столбец с названиями
объектов; пользователь подтверждает найденное или поправляет его.

Каждый шаг — метод, возвращающий:
  None           -> перейти к следующему шагу;
  Command.BACK   -> вернуться к предыдущему шагу;
  Command.BUILD  -> перейти к сборке файла;
  Command.CANCEL -> подтвердить и выйти в главное меню;
  "files"        -> вернуться к списку файлов;
  "restart"      -> начать заново (после сборки).
"""

from __future__ import annotations

import os
import re
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from .core import (
    Aggregator,
    AppError,
    KeySet,
    Orientation,
    SheetGrid,
    SourceConfig,
    TableLayout,
    Workbook,
    col_letter,
    column_paths,
    detect_layout,
    expand_inputs,
    export,
    extract_keys,
    find_key_candidates,
    fold_lookalikes,
    format_cell,
    layout_from,
    load_workbook,
    normalize_key,
    suggest_join,
)
from .ui import Command, UI
from .ui.screens import main_menu, show_instruction

_DEFAULT_OUTPUT_NAME = "Агрегированная таблица.xlsx"
_OUTPUT_PATH_EXAMPLE = r"C:\Users\Рабочий стол\таблица"

_EDGE_KEYS = 3        # сколько первых и последних названий показываем на проверке
_SAMPLE_VALUES = 3    # сколько примеров значений столбца показываем
_LOW_OVERLAP = 0.5    # доля общих названий, ниже которой предупреждаем
_EDIT_PREVIEW_ROWS = 12


def _plain(title: str) -> str:
    """Заголовок без регистра, пробелов и знаков — для сравнения между таблицами."""
    return re.sub(r"[\W_]+", "", title.casefold())


def _program_dir() -> Path:
    """Папка exe, а при запуске из исходников — корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def log_unexpected_error() -> None:
    """Дописывает текущее исключение в aggregator_error.log рядом с программой."""
    try:
        with open(_program_dir() / "aggregator_error.log", "a", encoding="utf-8") as fh:
            traceback.print_exc(file=fh)
    except OSError:
        pass


# Результаты _safe: повторить шаг / выйти в главное меню
_RETRY = object()
_ABORT = object()


@dataclass
class Table:
    """Источник данных: лист файла с таблицей.

    sheet — лист как в файле; grid — он же либо повёрнутый (если названия
    объектов идут в строку). layout и key_col находит программа; None значит
    «не распознано» — тогда problem объясняет, чего не хватает.
    chosen — выбранные столбцы характеристик; None — все.
    """

    path: Path
    sheet: SheetGrid
    grid: SheetGrid
    layout: TableLayout | None
    key_col: int | None = None
    keyset: KeySet | None = None
    enabled: bool = True
    chosen: list[int] | None = None
    paths: dict[int, list[str]] = field(default_factory=dict)
    problem: str = ""

    @property
    def rotated(self) -> bool:
        return self.grid is not self.sheet

    @property
    def label(self) -> str:
        return f"{self.path.name} / {self.sheet.name.strip()}"

    def title(self, c: int) -> str:
        """Полное название столбца одной строкой."""
        return " / ".join(self.paths.get(c) or [f"Столбец {col_letter(c)}"])


class AggregatorApp:
    """Приложение-агрегатор: главное меню и сценарий из трёх шагов."""

    def __init__(self) -> None:
        self.ui = UI()
        self._reset()

    def _reset(self) -> None:
        """Полный сброс настройки (новый запуск из главного меню)."""
        self.books: list[Workbook] = []
        self.tables: list[Table] = []
        self.skipped_sheets = 0            # листы, на которых таблица не найдена
        self.fold: bool | None = None      # считать ли буквы-двойники A/А одинаковыми
        self._notes: list[tuple[str, str]] = []   # сообщения для следующего экрана

    def _invalidate(self) -> None:
        """Список файлов изменился — найденное нужно определить заново."""
        self.tables = []
        self.fold = None

    def _match(self, key: str) -> str:
        """Форма названия для сравнения между источниками."""
        return fold_lookalikes(key) if self.fold else key.casefold()

    def _enabled(self) -> list[Table]:
        return [t for t in self.tables if t.enabled and t.keyset is not None]

    def _ready(self) -> bool:
        return bool(self._enabled())

    # ------------------------------------------------------------------
    # главный цикл

    def run(self) -> None:
        """Главное меню: Начать работу / Инструкция / Выход."""
        while True:
            choice = main_menu(self.ui)
            if choice == 0:
                self._reset()
                self._run_flow()
            elif choice == 1:
                show_instruction(self.ui)
            else:  # «Выход» или команда «отмена»
                self.ui.info("Работа программы завершена.")
                return

    def _run_flow(self) -> None:
        """Шаги сценария с навигацией назад/отмена/собрать."""
        steps = [self._step_files, self._step_review, self._step_columns, self._step_build]
        i = 0
        while 0 <= i < len(steps):
            action = self._safe(steps[i])
            if action is _RETRY:
                continue
            if action is _ABORT:
                return

            if action is None:
                i += 1
            elif action == "files":
                i = 0
            elif action == "restart":
                self._reset()
                i = 0
            elif action is Command.BACK:
                if i == 0:
                    self._notes.append(("warn", "Это первый шаг — назад вернуться нельзя."))
                else:
                    i -= 1
            elif action is Command.BUILD:
                i = len(steps) - 1
            elif action is Command.CANCEL:
                answer = self.ui.confirm(
                    "Отменить работу и выйти в главное меню? "
                    "Добавленные файлы и настройка будут сброшены",
                    default=False,
                )
                if answer is True:
                    return

    def _safe(self, step):
        """Выполняет шаг, не давая ошибке стереть настройку.

        AppError -> сообщение и _RETRY. Любая другая ошибка -> журнал и вопрос:
        повторить шаг (_RETRY) или выйти в главное меню (_ABORT).
        """
        try:
            return step()
        except AppError as exc:
            # Понятная ошибка core-слоя: показать и повторить шаг
            self.ui.error(exc.message)
            self.ui.console.input("[dim]Нажмите Enter, чтобы повторить шаг...[/]")
            return _RETRY
        except EOFError:
            raise  # конец ввода обрабатывает main.py
        except Exception:
            log_unexpected_error()
            self.ui.error(
                "Непредвиденная ошибка на этом шаге. Настройка сохранена; "
                "подробности записаны в файл aggregator_error.log."
            )
            retry = self.ui.confirm(
                "Повторить шаг? «н» — выйти в главное меню", default=True
            )
            return _RETRY if retry is True else _ABORT

    # ------------------------------------------------------------------
    # шаг 1: файлы

    def _step_files(self):
        while True:
            self.ui.screen("Шаг 1 из 3. Файлы")
            self.ui.explain(
                "Программа собирает сведения об одних и тех же объектах из нескольких таблиц\n"
                "Excel в одну общую таблицу. Укажите файлы с таблицами (.xls или .xlsx):\n"
                "  • путь к файлу — по одному, после каждого нажимайте Enter;\n"
                "  • или путь к папке — программа возьмёт из неё все таблицы Excel.\n"
                "Путь можно перетащить мышью в это окно или вставить из буфера обмена;\n"
                "кавычки вокруг пути не мешают."
            )
            if self.books:
                self.ui.info("Добавлены файлы:")
                for n, book in enumerate(self.books, start=1):
                    self.ui.info(f"  {n}. {book.path.name} — листов: {len(book.sheets)}")
                self.ui.info(
                    "Чтобы убрать файл из списка, введите его номер со знаком минус, например: -2"
                )
            for kind, message in self._notes:
                getattr(self.ui, kind)(message)
            self._notes = []
            self.ui.hint(allow_build=self._ready())

            raw = self.ui.ask(
                "Ещё файл или папка (Enter — перейти к проверке)"
                if self.books else "Путь к файлу или папке",
                default="" if self.books else None,
                allow_build=self._ready(),
            )
            if isinstance(raw, Command):
                return raw
            if raw == "":
                return None

            removal = re.fullmatch(r"-\s*(\d+)", raw)
            if removal:
                n = int(removal.group(1))
                if 1 <= n <= len(self.books):
                    removed = self.books.pop(n - 1)
                    self._invalidate()
                    self._notes.append(("success", f"Файл убран из списка: {removed.path.name}"))
                else:
                    self._notes.append(("warn", f"Файла с номером {n} в списке нет."))
                continue

            try:
                files = expand_inputs(raw)
            except AppError as exc:
                self._notes.append(("error", exc.message))
                continue
            for path in files:
                if len(files) > 1 and path.name == _DEFAULT_OUTPUT_NAME:
                    self._notes.append(("warn", f"Пропущен итоговый файл программы: {path.name}"))
                    continue
                if any(path.resolve() == book.path.resolve() for book in self.books):
                    self._notes.append(("warn", f"Файл уже добавлен: {path.name}"))
                    continue
                try:
                    book = load_workbook(path)
                except AppError as exc:
                    self._notes.append(("error", exc.message))
                    continue
                self.books.append(book)
                self._invalidate()
                self._notes.append(
                    ("success", f"Файл загружен: {path.name} (листов: {len(book.sheets)})")
                )

    # ------------------------------------------------------------------
    # автоматический разбор

    def _refresh(self, table: Table) -> None:
        """Пересчитывает названия столбцов и набор названий объектов источника."""
        table.keyset = None
        table.chosen = None
        table.problem = ""
        table.paths = {}
        if table.layout is None:
            table.problem = "таблица не распознана — укажите первую строку данных"
            return
        table.paths = column_paths(table.grid, table.layout)
        if table.key_col is None:
            table.problem = "не найден столбец с названиями объектов"
            return
        try:
            table.keyset = extract_keys(
                table.grid, table.layout.data_start, table.key_col, Orientation.VERTICAL
            )
        except AppError:
            table.problem = "в выбранном столбце нет названий объектов"

    def _analyse(self) -> None:
        """Находит таблицы на листах всех файлов и столбцы, по которым их соединять."""
        tables: list[Table] = []
        self.skipped_sheets = 0
        for book in self.books:
            found = [
                Table(book.path, sheet, sheet, layout)
                for sheet in book.sheets
                if (layout := detect_layout(sheet)) is not None
            ]
            self.skipped_sheets += len(book.sheets) - len(found)
            if not found and book.sheets:
                # Ни на одном листе таблица не распознана — показываем самый
                # большой лист, первую строку данных укажет пользователь
                biggest = max(book.sheets, key=lambda s: s.nrows * s.ncols)
                found = [Table(book.path, biggest, biggest, None)]
                self.skipped_sheets -= 1
            tables.extend(found)

        recognised = [t for t in tables if t.layout is not None]
        for table in recognised:
            table.paths = column_paths(table.grid, table.layout)
        columns = suggest_join(
            [(t.grid, t.layout.data_start) for t in recognised],
            [{c: t.title(c) for c in t.paths} for t in recognised],
        )
        for table, column in zip(recognised, columns):
            table.key_col = column
        for table in tables:
            self._refresh(table)
            table.enabled = table.keyset is not None
        self.tables = tables

    def _guess_key(self, table: Table) -> int | None:
        """Столбец названий для одного источника: по совпадениям с остальными."""
        known = [k for t in self._enabled() if t is not table for k in t.keyset.keys]
        if known:
            candidates = find_key_candidates(table.grid, known, Orientation.VERTICAL, top=1)
            if candidates:
                return candidates[0].index
        titles = column_paths(table.grid, table.layout)
        return suggest_join(
            [(table.grid, table.layout.data_start)],
            [{c: " / ".join(path) for c, path in titles.items()}],
        )[0]

    def _shared(self) -> dict[int, int]:
        """id(источника) -> сколько его названий встречается в других источниках."""
        tables = self._enabled()
        forms = [{self._match(k) for k in t.keyset.keys} for t in tables]
        result = {}
        for i, table in enumerate(tables):
            others = set().union(*(f for j, f in enumerate(forms) if j != i))
            result[id(table)] = len(forms[i] & others)
        return result

    def _lookalike_pairs(self) -> list[tuple[str, ...]]:
        """Названия, отличающиеся только латинскими/русскими буквами-двойниками."""
        forms: dict[str, set[str]] = {}
        for table in self._enabled():
            for key in table.keyset.keys:
                forms.setdefault(fold_lookalikes(key), set()).add(key)
        groups = []
        for variants in forms.values():
            distinct = {v.casefold(): v for v in variants}
            if len(distinct) > 1:
                groups.append(tuple(sorted(distinct.values())))
        return groups

    def _key_title(self) -> str:
        """Название первого столбца итога — заголовок столбца названий первого источника."""
        base = self._enabled()[0]
        return base.title(base.key_col)

    def _samples(self, table: Table, c: int) -> str:
        """Несколько первых значений столбца из области данных — для подсказки."""
        values = []
        for r in range(table.layout.data_start, table.grid.nrows):
            value = table.grid.value(r, c)
            if normalize_key(value) is not None:
                if isinstance(value, float) and not value.is_integer():
                    value = f"{value:.6g}"  # для подсказки хватит шести знаков
                values.append(str(format_cell(value)).replace("\n", " ")[:20])
                if len(values) >= _SAMPLE_VALUES:
                    break
        return ", ".join(values)

    def _data_columns(self, table: Table) -> list[int]:
        """Столбцы характеристик: все, кроме столбца названий и пустых."""
        rows = list(table.keyset.raw_positions.values())
        return [
            c for c in range(table.grid.ncols)
            if c != table.key_col
            and any(normalize_key(table.grid.value(r, c)) is not None for r in rows)
        ]

    # ------------------------------------------------------------------
    # шаг 2: проверка найденного

    def _step_review(self):
        if not self.tables:
            self.ui.info("Анализирую таблицы...")
            self._analyse()
        while True:
            self.ui.screen("Шаг 2 из 3. Проверка")
            self.ui.explain(
                "Программа сама нашла в каждом файле таблицу, её шапку и столбец с названиями\n"
                "объектов — по этому столбцу таблицы соединяются. Проверьте, что всё определено\n"
                "верно. Если да — нажмите Enter. Если нет — выберите «Изменить источник»."
            )
            self._show_review()

            if self.fold is None and self._ready():
                pairs = self._lookalike_pairs()
                if pairs:
                    examples = "; ".join(" и ".join(f"«{v}»" for v in pair) for pair in pairs[:3])
                    self.ui.warn(
                        f"Есть названия ({len(pairs)} шт.), которые выглядят одинаково, но набраны "
                        f"разными буквами — латинскими и русскими: {examples}."
                    )
                    answer = self.ui.confirm("Считать такие названия одним объектом?", default=True)
                    if isinstance(answer, Command):
                        return answer
                    self.fold = answer
                    continue
                self.fold = False

            choice = self.ui.choose(
                "Всё определено верно?",
                ["Да, продолжить", "Изменить источник", "Добавить или убрать файлы"],
                allow_build=self._ready(),
                default=0,
            )
            if choice == 0:
                if self._ready():
                    return None
                self.ui.warn(
                    "Нет ни одного готового источника — выберите «Изменить источник» "
                    "и укажите, где находится таблица."
                )
                self.ui.pause()
            elif choice == 1:
                result = self._edit_source()
                if isinstance(result, Command):
                    return result
            elif choice == 2:
                return "files"
            else:
                return choice  # BACK / BUILD / CANCEL

    def _show_review(self) -> None:
        """Сводка найденного по всем источникам и предупреждения."""
        shared = self._shared()
        rows, dim = [], set()
        for n, table in enumerate(self.tables, start=1):
            layout = table.layout
            if layout is None:
                header = data = "—"
            else:
                head = [r + 1 for r in layout.header_rows]
                if not head:
                    header = "нет"
                elif len(head) == 1:
                    header = f"строка {head[0]}"
                else:
                    header = f"строки {head[0]}–{head[-1]}"
                data = f"строки {layout.data_start + 1}–{table.grid.nrows}"
            sheet = table.sheet.name.strip() + (" (повёрнут)" if table.rotated else "")
            if table.keyset is None:
                key, found, common = "—", table.problem, "—"
            else:
                key = f"{col_letter(table.key_col)} «{table.title(table.key_col)}»"
                found = str(len(table.keyset.keys))
                common = str(shared.get(id(table), 0)) if table.enabled else "не используется"
            if not table.enabled or table.keyset is None:
                dim.add(n - 1)
            rows.append([str(n), table.path.name, sheet, header, data, key, found, common])
        self.ui.show_table(
            "Что найдено",
            ["№", "Файл", "Лист", "Шапка", "Данные", "Названия объектов", "Найдено", "Общих"],
            rows,
            dim_rows=dim,
        )
        if self.skipped_sheets:
            self.ui.info(f"Листов без таблицы пропущено: {self.skipped_sheets}.")

        enabled = self._enabled()
        for n, table in enumerate(self.tables, start=1):
            if table not in enabled:
                continue
            keys = table.keyset.keys
            if len(keys) <= 2 * _EDGE_KEYS:
                edge = ", ".join(keys)
            else:
                edge = (f"{', '.join(keys[:_EDGE_KEYS])} … "
                        f"последние: {', '.join(keys[-_EDGE_KEYS:])}")
            self.ui.info(f"  {n}. Названия объектов: {edge}")
            for warning in self._warnings(table, shared.get(id(table), 0), len(enabled)):
                self.ui.warn(f"     {warning}")

        if len(enabled) == 1:
            self.ui.info(
                "Источник один: соединять не с чем, в итог попадут столбцы этой таблицы."
            )
        elif len(enabled) > 1:
            if not any(shared.get(id(t), 0) for t in enabled):
                self.ui.error(
                    "Общих названий объектов между таблицами не найдено: данные разных "
                    "файлов не соединятся, каждый объект получит строку только из своего "
                    "файла. Проверьте столбцы названий («Изменить источник»)."
                )
                return
            self.ui.success(
                f"Таблицы соединяются по названиям объектов («{self._key_title()}»)."
            )
            base = enabled[0]
            for table in enabled[1:]:
                if _plain(table.title(table.key_col)) != _plain(base.title(base.key_col)):
                    self.ui.warn(
                        f"В таблице «{table.label}» названия объектов взяты из столбца "
                        f"«{table.title(table.key_col)}», а в первой — из "
                        f"«{base.title(base.key_col)}». Убедитесь, что это одно и то же."
                    )

    def _warnings(self, table: Table, shared: int, enabled: int) -> list[str]:
        """Предупреждения по одному источнику."""
        keyset = table.keyset
        result = []
        if enabled > 1 and shared < _LOW_OVERLAP * len(keyset.keys):
            others = sum(len(t.keyset.keys) for t in self._enabled() if t is not table)
            if shared < _LOW_OVERLAP * min(len(keyset.keys), others):
                result.append(
                    f"с другими таблицами совпало только {shared} из {len(keyset.keys)} "
                    "названий — проверьте столбец названий"
                )
        if keyset.duplicates:
            where = ", ".join(str(r + 1) for r in keyset.duplicate_rows[:5])
            result.append(
                f"повторы названий: {len(keyset.duplicates)} (строки {where}) — "
                "берётся первое вхождение"
            )
        if keyset.skipped:
            examples = "; ".join(f"«{text[:30]}»" for _, text in keyset.skipped[:3])
            result.append(
                f"строки-подписи пропущены: {len(keyset.skipped)} ({examples}) — это не объекты"
            )
        if table.grid.error_cells:
            result.append(
                f"ячеек с ошибками (#REF!, #Н/Д и т.п.) на листе: {table.grid.error_cells} — "
                "они переносятся пустыми"
            )
        if table.grid.uncached_formulas:
            result.append(
                f"формул без сохранённого значения: {table.grid.uncached_formulas} — "
                "они читаются пустыми; откройте файл в Excel и сохраните его"
            )
        return result

    # ------------------------------------------------------------------
    # ручная правка источника

    def _edit_source(self):
        """Меню правки: какой источник и что в нём изменить."""
        self.ui.screen("Изменить источник")
        options = [t.label for t in self.tables] + ["Добавить другой лист из файла"]
        index = self.ui.choose("Какой источник изменить?", options)
        if isinstance(index, Command):
            return None if index is Command.BACK else index
        if index == len(self.tables):
            table = self._add_sheet()
            if not isinstance(table, Table):
                return None if table is Command.BACK else table
        else:
            table = self.tables[index]

        while True:
            self.ui.screen(f"Источник: {table.label}")
            self.ui.show_preview(
                table.grid, title="Начало листа (слева — номера строк)", limit=_EDIT_PREVIEW_ROWS
            )
            if table.keyset is not None:
                self.ui.info(
                    f"Сейчас: данные со строки {table.layout.data_start + 1}; названия объектов — "
                    f"столбец {col_letter(table.key_col)} «{table.title(table.key_col)}», "
                    f"найдено {len(table.keyset.keys)}."
                )
            else:
                self.ui.warn(f"Источник не готов: {table.problem}.")
            choice = self.ui.choose(
                "Что изменить?",
                [
                    "Указать первую строку данных",
                    "Выбрать столбец с названиями объектов",
                    "Вернуть обычное положение таблицы" if table.rotated
                    else "Таблица повёрнута: названия объектов идут в строку, а не в столбец",
                    "Использовать этот источник" if not table.enabled
                    else "Не использовать этот источник",
                    "Готово — вернуться к проверке",
                ],
                default=4,
            )
            if isinstance(choice, Command):
                return None if choice is Command.BACK else choice
            if choice == 4:
                return None
            self.fold = None  # названия изменились — буквы-двойники проверим заново
            if choice == 0:
                result = self._ask_data_start(table)
            elif choice == 1:
                result = self._ask_key_column(table)
            elif choice == 2:
                table.grid = table.sheet if table.rotated else table.sheet.transposed()
                table.layout = detect_layout(table.grid)
                table.key_col = self._guess_key(table) if table.layout else None
                self._refresh(table)
                table.enabled = table.keyset is not None
                result = None
            else:
                if table.enabled:
                    table.enabled = False
                elif table.keyset is not None:
                    table.enabled = True
                else:
                    self.ui.warn(f"Источник не готов: {table.problem}.")
                    self.ui.pause()
                result = None
            if isinstance(result, Command) and result is not Command.BACK:
                return result

    def _add_sheet(self):
        """Добавляет источником лист, который программа не выбрала сама."""
        book = self.books[0]
        if len(self.books) > 1:
            index = self.ui.choose("Из какого файла?", [b.path.name for b in self.books])
            if isinstance(index, Command):
                return index
            book = self.books[index]
        items = [f"{s.name.strip()}  ({s.nrows} строк × {s.ncols} столбцов)" for s in book.sheets]
        index = self.ui.choose("Какой лист добавить?", items)
        if isinstance(index, Command):
            return index
        sheet = book.sheets[index]
        existing = next(
            (t for t in self.tables if t.sheet is sheet), None
        )
        if existing is not None:
            return existing
        table = Table(book.path, sheet, sheet, detect_layout(sheet))
        if table.layout is not None:
            table.key_col = self._guess_key(table)
        self._refresh(table)
        table.enabled = table.keyset is not None
        self.tables.append(table)
        self.skipped_sheets = max(0, self.skipped_sheets - 1)
        return table

    def _ask_data_start(self, table: Table):
        """Ручной ввод номера первой строки данных."""
        self.ui.info(
            "Введите номер строки, в которой записан первый объект (номера строк — слева "
            "в таблице выше, как в Excel). Всё, что выше этой строки, программа считает шапкой."
        )
        while True:
            raw = self.ui.ask("Номер первой строки данных")
            if isinstance(raw, Command):
                return raw
            if raw.isdigit() and 1 <= int(raw) <= table.grid.nrows:
                break
            self.ui.warn(f"Введите число от 1 до {table.grid.nrows}.")
        table.layout = layout_from(table.grid, int(raw) - 1)
        if table.key_col is None:
            table.key_col = self._guess_key(table)
        self._refresh(table)
        table.enabled = table.keyset is not None
        return None

    def _ask_key_column(self, table: Table):
        """Ручной выбор столбца с названиями объектов."""
        if table.layout is None:
            self.ui.warn("Сначала укажите первую строку данных.")
            self.ui.pause()
            return None
        columns = [
            c for c in range(table.grid.ncols) if self._samples(table, c)
        ]
        items = [
            f"{col_letter(c)} — {table.title(c)} (напр.: {self._samples(table, c)})"
            for c in columns
        ]
        self.ui.info(
            "Выберите столбец, в котором записаны названия (номера, коды) объектов — "
            "те самые, что повторяются в других таблицах."
        )
        index = self.ui.choose("Столбец с названиями объектов", items)
        if isinstance(index, Command):
            return index
        table.key_col = columns[index]
        self._refresh(table)
        table.enabled = table.keyset is not None
        return None

    # ------------------------------------------------------------------
    # шаг 3: столбцы для переноса

    def _step_columns(self):
        tables = self._enabled()
        index = 0
        while index < len(tables):
            table = tables[index]
            columns = self._data_columns(table)
            if not columns:
                table.chosen = []
                index += 1
                continue
            self.ui.screen(f"Шаг 3 из 3. Столбцы ({index + 1} из {len(tables)})")
            self.ui.explain(
                f"Таблица: {table.label}\n"
                "Выберите столбцы, которые нужно перенести в итоговый файл. Нажмите Enter,\n"
                "чтобы перенести все, или введите номера нужных столбцов.\n"
                "Столбец с названиями объектов уже учтён — он станет первым столбцом итога."
            )
            items = [
                f"{col_letter(c)} — {table.title(c)} (напр.: {self._samples(table, c)})"
                for c in columns
            ]
            selected = self.ui.multi_select(
                "Столбцы для переноса", items, default_all=True, allow_build=True
            )
            if selected is Command.BACK:
                if index == 0:
                    return Command.BACK
                index -= 1
                continue
            if isinstance(selected, Command):
                return selected  # «собрать»: у оставшихся таблиц берутся все столбцы
            table.chosen = [columns[i] for i in selected]
            index += 1
        return None

    # ------------------------------------------------------------------
    # сборка итогового файла

    def _step_build(self):
        self.ui.screen("Сборка итогового файла")
        tables = self._enabled()
        key_title = self._key_title()
        aggregator = Aggregator(key_title, match=self._match)
        for table in tables:
            columns = table.chosen if table.chosen is not None else self._data_columns(table)
            aggregator.add_source(
                table.keyset,
                SourceConfig(
                    file_path=table.path,
                    sheet_name=table.sheet.name,
                    anchor=(table.layout.data_start, table.key_col),
                    orientation=Orientation.VERTICAL,
                    key_index=table.key_col,
                    char_columns=columns,
                    column_titles=[table.title(c) for c in columns],
                    column_paths=[table.paths[c] for c in columns],
                ),
                table.grid,
            )
        headers, rows = aggregator.build_table()

        self.ui.info(
            f"Готово к записи: объектов — {len(rows)}, столбцов — {len(headers)}, "
            f"источников — {len(tables)}."
        )
        output_dir = _program_dir()
        default_path = str(output_dir / _DEFAULT_OUTPUT_NAME)
        self.ui.explain(
            "Куда сохранить итоговый файл? Нажмите Enter — файл появится рядом с программой\n"
            f"под именем «{_DEFAULT_OUTPUT_NAME}». Либо введите своё имя файла или полный\n"
            "путь; расширение .xlsx добавится само. Имя файла станет названием таблицы."
        )
        while True:
            raw = self.ui.ask(
                "Путь и имя итогового файла",
                default=default_path,
                display_default=_OUTPUT_PATH_EXAMPLE,
            )
            if isinstance(raw, Command):
                if raw is Command.BUILD:
                    self.ui.warn("Файл уже собирается — укажите путь сохранения.")
                    continue
                return raw
            cleaned = raw.strip().strip('"').strip("'").strip()
            out = Path(cleaned)
            if not out.name:
                self.ui.warn("Укажите имя файла, а не только папку.")
                continue
            if not out.is_absolute():
                out = output_dir / out
            if out.suffix.lower() == ".xls":
                out = out.with_suffix(".xlsx")
            elif out.suffix.lower() != ".xlsx":
                # Дописываем к имени: with_suffix обрезал бы «отчёт v1.2» до «отчёт v1»
                out = out.with_name(out.name + ".xlsx")

            if any(out.resolve() == book.path.resolve() for book in self.books):
                self.ui.warn(
                    f"«{out.name}» — один из исходных файлов. Исходные файлы программа "
                    "не изменяет: выберите другое имя."
                )
                continue
            if out.exists():
                overwrite = self.ui.confirm(
                    f"Файл «{out.name}» уже существует — перезаписать?", default=False
                )
                if overwrite is not True:
                    if isinstance(overwrite, Command):
                        return overwrite
                    continue

            try:
                checked = export(
                    out, headers, rows, key_title,
                    header_paths=aggregator.header_paths(), title=out.stem,
                )
            except AppError as exc:
                self.ui.error(exc.message)
                retry = self.ui.confirm("Повторить запись?", default=True)
                if retry is True:
                    continue
                if isinstance(retry, Command):
                    return retry
                return None
            break

        self.ui.success(f"Файл собран: {out}")
        self._report(aggregator, tables, len(rows), len(headers), checked)

        open_it = self.ui.confirm("Открыть собранный файл?", default=True)
        if open_it is True:
            try:
                os.startfile(str(out))  # noqa: S606 — штатное действие Windows
            except OSError:
                self.ui.warn("Не удалось открыть файл автоматически — откройте его вручную.")
        elif isinstance(open_it, Command):
            return open_it

        again = self.ui.confirm(
            "Собрать ещё один файл из других таблиц?", default=False
        )
        if again is True:
            return "restart"
        if isinstance(again, Command):
            return again
        return None

    def _report(self, aggregator: Aggregator, tables: list[Table],
                nrows: int, ncols: int, checked: int) -> None:
        """Отчёт о сборке: что откуда взято и что проверено."""
        self.ui.success(f"Строк: {nrows}, столбцов: {ncols}")
        shared = self._shared()
        rows = []
        for table, (_, covered, total) in zip(tables, aggregator.coverage_stats()):
            columns = table.chosen if table.chosen is not None else self._data_columns(table)
            rows.append([
                table.label,
                f"строки {table.layout.data_start + 1}–{table.grid.nrows}",
                str(len(table.keyset.keys)),
                f"{covered} из {total}",
                str(len(columns)),
            ])
        self.ui.show_table(
            "Отчёт о сборке",
            ["Источник", "Данные", "Названий", "Объектов заполнено", "Столбцов перенесено"],
            rows,
        )
        for table in tables:
            for warning in self._warnings(table, shared.get(id(table), 0), len(tables)):
                self.ui.warn(f"{table.label}: {warning}")
        self.ui.success(
            f"Запись проверена: файл перечитан, {checked} ячеек данных совпали с исходными."
        )
