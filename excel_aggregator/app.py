"""Машина состояний приложения: шаги настройки и сборка итогового файла.

Каждый шаг — метод, возвращающий:
  None           -> перейти к следующему шагу;
  Command.BACK   -> вернуться к предыдущему шагу;
  Command.BUILD  -> перейти к сборке файла;
  Command.CANCEL -> подтвердить и выйти в главное меню;
  "add_file"     -> вставить блок шагов для ещё одного файла (только _step_more).
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from .core import (
    Aggregator,
    AppError,
    Orientation,
    SourceConfig,
    col_letter,
    export,
    extract_keys,
    find_key_candidates,
    format_cell,
    load_workbook,
    normalize_key,
    parse_coordinate,
    resolve_input_path,
)
from .ui import Command, UI
from .ui.screens import main_menu, show_instruction

_DEFAULT_OUTPUT_NAME = "Агрегированная таблица.xlsx"
_OUTPUT_PATH_EXAMPLE = r"C:\Users\Рабочий стол\таблица"

# Строки нумерации граф ("1", "2", "1а", "15") — не считаем их заголовками
_NUMBERING_RE = re.compile(r"\d{1,2}[а-яa-z]?")

# Значения из строк единиц измерения — тоже не заголовки
_UNIT_VALUES = {
    "м", "метр", "метры", "км", "см", "мм",
    "фт", "фут", "футы",
    "°", "'", '"', "′", "″", "с", "град", "мин", "сек",
}

# Доля одинаковых значений в строке/столбце, при которой это не заголовок
# графы, а общий титул/разделитель («район 2», название документа)
_SECTION_SHARE = 0.5


def _looks_like_header(value) -> bool:
    """Похоже ли значение ячейки на заголовок (а не на нумерацию/единицы)."""
    if value is None or str(value).strip() == "":
        return False
    text = str(format_cell(value)).strip()
    if _NUMBERING_RE.fullmatch(text):
        return False
    if text.casefold() in _UNIT_VALUES:
        return False
    return True


def _program_dir() -> Path:
    """Папка exe, а при запуске из исходников — корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _is_section_row(grid, j: int, value, vertical: bool) -> bool:
    """Строка (столбец) — титул/разделитель, а не заголовок графы?

    Признаки: то же значение занимает большую часть строки (объединённый
    титул) или в строке всего 1–2 различных значения при широкой таблице
    (строка-разделитель вроде «район 2»).
    """
    line = grid.row(j) if vertical else grid.col(j)
    width = len(line)
    normalized = [
        normalize_key(v) for v in line if normalize_key(v) is not None
    ]
    if not normalized:
        return True
    candidate = normalize_key(value)
    same = sum(1 for v in normalized if v.casefold() == candidate.casefold())
    if same / width >= _SECTION_SHARE:
        return True
    distinct = {v.casefold() for v in normalized}
    return width >= 6 and len(distinct) <= 2


class AggregatorApp:
    """Приложение-агрегатор: главное меню и пошаговая настройка."""

    def __init__(self) -> None:
        self.ui = UI()
        self._reset()

    def _reset(self) -> None:
        """Полный сброс настройки (новый запуск из главного меню)."""
        self.sources: list = []          # (KeySet, SourceConfig, SheetGrid)
        self.base_keys: list[str] = []   # ключи первого файла (база для поиска)
        self.key_title = "Название"
        self.first_file_dir: Path = Path.cwd()
        self._wb = None                  # текущая книга (первый/очередной файл)
        self._first_sheets: list = []
        self._next_sheets: list = []
        self._base_orientation = Orientation.VERTICAL
        self._block_base: int | None = None  # отметка начала блока «очередной файл»

    # ------------------------------------------------------------------
    # главный цикл

    def run(self) -> None:
        """Главное меню: Инструкция / Запуск программы / Выход."""
        while True:
            choice = main_menu(self.ui)
            if choice == 0:
                show_instruction(self.ui)
            elif choice == 1:
                self._reset()
                self._run_flow()
            else:  # «Выход» или команда «отмена»
                self.ui.info("Работа программы завершена.")
                return

    def _run_flow(self) -> None:
        """Пошаговая настройка с навигацией назад/отмена/собрать."""
        steps = [
            self._step_first_file,
            self._step_first_sheets,
            self._step_first_source,
            self._step_more,
        ]
        i = 0
        while 0 <= i < len(steps):
            try:
                action = steps[i]()
            except AppError as exc:
                # Понятная ошибка core-слоя: показать и повторить шаг
                self.ui.error(exc.message)
                self.ui.console.input("[dim]Нажмите Enter, чтобы повторить шаг...[/]")
                continue

            if action is None:
                i += 1
            elif action == "add_file":
                steps[i + 1:i + 1] = [
                    self._step_next_file,
                    self._step_next_sheets,
                    self._step_next_source,
                    self._step_more,
                ]
                i += 1
            elif action is Command.BACK:
                if i == 0:
                    self.ui.warn("Это первый шаг настройки — назад вернуться нельзя.")
                else:
                    i -= 1
            elif action is Command.BUILD:
                if not self.sources:
                    self.ui.warn("Сначала настройте хотя бы один источник данных.")
                    continue
                outcome = self._step_build()
                if outcome == "restart":
                    self._reset()
                    i = 0
                elif outcome is None:
                    return  # в главное меню
                elif outcome is Command.CANCEL:
                    return
                # Command.BACK — остаёмся на том же шаге
            elif action is Command.CANCEL:
                answer = self.ui.confirm(
                    "Отменить настройку и выйти в главное меню? "
                    "Проделанная настройка будет потеряна",
                    default=False,
                )
                if answer is True:
                    return

    # ------------------------------------------------------------------
    # шаг 1: путь к первому файлу

    def _step_first_file(self):
        self.ui.screen("Шаг 1. Первый файл")
        self.ui.info(
            "Укажите путь к файлу Excel (.xls или .xlsx).\n"
            "Путь можно перетащить мышью в окно консоли или вставить из буфера —\n"
            "кавычки программа уберёт автоматически."
        )
        self.ui.hint()
        raw = self.ui.ask("Путь к файлу")
        if isinstance(raw, Command):
            return raw
        path = resolve_input_path(raw)          # AppError -> шаг повторится
        wb = load_workbook(path)
        self._wb = wb
        self.first_file_dir = path.parent
        self.ui.success(f"Файл загружен: {path.name} (листов: {len(wb.sheets)})")
        self.ui.show_preview(wb.sheets[0], title=f"Лист «{wb.sheets[0].name}» — первые строки")
        return None

    # ------------------------------------------------------------------
    # шаг 2: выбор листов первого файла

    def _step_first_sheets(self):
        self.ui.screen("Шаг 2. Листы первого файла")
        sheets = self._wb.sheets
        if len(sheets) == 1:
            self.ui.info(f"В файле один лист: «{sheets[0].name}» — выбран автоматически.")
            self._first_sheets = list(sheets)
            return None
        items = [f"{sh.name}  ({sh.nrows} строк × {sh.ncols} столбцов)" for sh in sheets]
        selected = self.ui.multi_select("Выберите листы для обработки", items)
        if isinstance(selected, Command):
            return selected
        self._first_sheets = [sheets[i] for i in selected]
        self.ui.success(f"Выбрано листов: {len(self._first_sheets)}")
        return None

    # ------------------------------------------------------------------
    # шаг 3: ключи и характеристики первого файла

    def _step_first_source(self):
        self.ui.screen("Шаг 3. Названия объектов (первый файл)")
        self.sources.clear()
        self._block_base = None

        idx = 0
        while idx < len(self._first_sheets):
            grid = self._first_sheets[idx]
            result = self._configure_source_manual(grid, self._wb.path)
            if result is Command.BACK:
                if idx == 0:
                    return Command.BACK
                self.sources.pop()
                idx -= 1
                continue
            if isinstance(result, Command):
                return result
            idx += 1

        self.base_keys = [k for ks, _, _ in self.sources for k in ks.keys]
        self._base_orientation = self.sources[0][1].orientation
        return None

    # ------------------------------------------------------------------
    # шаг 4: что дальше (добавить файл / собрать)

    def _step_more(self):
        self.ui.screen("Источники данных")
        self.ui.info("Настроенные источники:")
        for n, (keyset, config, _) in enumerate(self.sources, start=1):
            self.ui.info(
                f"  {n}. {Path(config.file_path).name} / {config.sheet_name.strip()}"
                f" — названий: {len(keyset.keys)}, характеристик: {len(config.char_columns)}"
            )
        choice = self.ui.choose(
            "Что делаем дальше?",
            ["Добавить ещё файл", "Собрать итоговый файл"],
            allow_build=True,
        )
        if choice == 0:
            self._block_base = None  # новый блок «очередной файл»
            return "add_file"
        if choice == 1:
            return Command.BUILD
        return choice  # BACK / CANCEL

    # ------------------------------------------------------------------
    # очередной файл: путь, листы, автопоиск

    def _step_next_file(self):
        self.ui.screen("Очередной файл")
        if self._block_base is None:
            self._block_base = len(self.sources)
        else:
            # Повторный вход (шаг «назад»): убрать источники этого блока
            del self.sources[self._block_base:]
        self.ui.info("Укажите путь к следующему файлу Excel.")
        self.ui.hint(allow_build=True)
        raw = self.ui.ask("Путь к файлу", allow_build=True)
        if isinstance(raw, Command):
            return raw
        path = resolve_input_path(raw)
        wb = load_workbook(path)
        self._wb = wb
        self.ui.success(f"Файл загружен: {path.name} (листов: {len(wb.sheets)})")
        self.ui.show_preview(wb.sheets[0], title=f"Лист «{wb.sheets[0].name}» — первые строки")
        return None

    def _step_next_sheets(self):
        self.ui.screen("Листы очередного файла")
        sheets = self._wb.sheets
        if len(sheets) == 1:
            self.ui.info(f"В файле один лист: «{sheets[0].name}» — выбран автоматически.")
            self._next_sheets = list(sheets)
            return None
        items = [f"{sh.name}  ({sh.nrows} строк × {sh.ncols} столбцов)" for sh in sheets]
        selected = self.ui.multi_select("Выберите листы для обработки", items)
        if isinstance(selected, Command):
            return selected
        self._next_sheets = [sheets[i] for i in selected]
        self.ui.success(f"Выбрано листов: {len(self._next_sheets)}")
        return None

    def _step_next_source(self):
        self.ui.screen("Поиск названий в очередном файле")
        if self._block_base is not None:
            del self.sources[self._block_base:]

        idx = 0
        while idx < len(self._next_sheets):
            grid = self._next_sheets[idx]
            result = self._configure_source_auto(grid, self._wb.path)
            if result is Command.BACK:
                if idx == 0:
                    return Command.BACK
                self.sources.pop()
                idx -= 1
                continue
            if isinstance(result, Command):
                return result
            idx += 1
        return None

    # ------------------------------------------------------------------
    # настройка одного источника

    def _configure_source_manual(self, grid, file_path: Path):
        """Ручное указание координат ключей + выбор характеристик."""
        self.ui.screen(f"Лист «{grid.name.strip()}» ({Path(file_path).name})")
        self.ui.show_preview(grid, title=f"Лист «{grid.name}» — первые строки")

        while True:
            raw = self.ui.ask(
                "Координата первой ячейки с названиями объектов (например, A6)",
                allow_build=bool(self.sources),
            )
            if isinstance(raw, Command):
                return raw
            try:
                row0, col0 = parse_coordinate(raw)
            except AppError as exc:
                self.ui.error(exc.message)
                continue

            orient_choice = self.ui.choose(
                "Как расположены названия объектов?",
                ["Вертикально (сверху вниз)", "Горизонтально (слева направо)"],
                allow_build=bool(self.sources),
            )
            if isinstance(orient_choice, Command):
                return orient_choice
            orientation = (
                Orientation.VERTICAL if orient_choice == 0 else Orientation.HORIZONTAL
            )
            try:
                keyset = extract_keys(grid, row0, col0, orientation)
            except AppError as exc:
                self.ui.error(exc.message)
                continue
            break

        self._report_keyset(keyset)
        return self._finish_source(grid, file_path, keyset, row0, col0, orientation)

    def _configure_source_auto(self, grid, file_path: Path):
        """Автопоиск столбца/строки ключей, при неудаче — ручной режим."""
        self.ui.screen(f"Лист «{grid.name.strip()}» ({Path(file_path).name})")
        self.ui.info("Ищу совпадения с названиями из первого файла...")
        candidates = find_key_candidates(grid, self.base_keys, self._base_orientation)

        if not candidates:
            self.ui.warn(
                "Автоматически совпадения не найдены — "
                "укажите координаты вручную."
            )
            return self._configure_source_manual(grid, file_path)

        pos = 0
        while True:
            cand = candidates[pos]
            where = (
                f"столбце {col_letter(cand.index)}"
                if self._base_orientation is Orientation.VERTICAL
                else f"строке {cand.index + 1}"
            )
            self.ui.success(
                f"Найдено совпадений: {cand.match_count} из {len(self.base_keys)} "
                f"в {where}"
            )
            sample = ", ".join(str(v) for v in cand.sample)
            self.ui.info(f"Первые значения ({len(cand.sample)} шт.): {sample}")

            options = ["Подтвердить — это нужный столбец/строка"]
            if len(candidates) > 1:
                options.append("Показать следующий вариант")
            options.append("Указать координаты вручную")
            choice = self.ui.choose("Подтвердите вариант", options, allow_build=True)
            if isinstance(choice, Command):
                return choice
            if choice == 0:
                if self._base_orientation is Orientation.VERTICAL:
                    row0, col0 = cand.first_row0, cand.index
                else:
                    row0, col0 = cand.index, cand.first_row0
                orientation = self._base_orientation
                try:
                    keyset = extract_keys(grid, row0, col0, orientation)
                except AppError as exc:
                    self.ui.error(exc.message)
                    continue
                self._report_keyset(keyset)
                return self._finish_source(
                    grid, file_path, keyset, row0, col0, orientation
                )
            if len(candidates) > 1 and choice == 1:
                pos = (pos + 1) % len(candidates)
                continue
            return self._configure_source_manual(grid, file_path)

    def _report_keyset(self, keyset) -> None:
        """Обратная связь после извлечения ключей."""
        self.ui.success(f"Найдено значений: {len(keyset.keys)}")
        first_five = ", ".join(keyset.keys[:5])
        self.ui.info(f"Первые пять: {first_five}")
        if keyset.duplicates:
            examples = ", ".join(keyset.duplicates[:5])
            self.ui.warn(
                f"Найдены повторы названий ({len(keyset.duplicates)} шт., "
                f"использовано первое вхождение): {examples}"
            )

    def _finish_source(self, grid, file_path: Path, keyset, row0, col0, orientation):
        """Название ключевого столбца, выбор характеристик, добавление источника."""
        # Первый источник определяет название ключевого столбца итогового файла
        if not self.sources:
            key_index = col0 if orientation is Orientation.VERTICAL else row0
            key_header = self._guess_header(grid, key_index, row0, col0, orientation)
            if key_header:
                self.key_title = key_header
                self.ui.info(
                    "Столбец с названиями объектов в итоговом файле "
                    f"назван автоматически: «{key_header}»"
                )
            else:
                title = self.ui.ask(
                    "Не удалось прочитать заголовок столбца с названиями объектов —\n"
                    "введите его название в итоговом файле",
                    default=self.key_title,
                )
                if isinstance(title, Command):
                    return title
                self.key_title = title

        picked = self._select_characteristics(grid, row0, col0, orientation)
        if isinstance(picked, Command):
            return picked
        char_columns, column_titles = picked

        config = SourceConfig(
            file_path=Path(file_path),
            sheet_name=grid.name,
            anchor=(row0, col0),
            orientation=orientation,
            key_index=col0 if orientation is Orientation.VERTICAL else row0,
            char_columns=char_columns,
            column_titles=column_titles,
        )
        self.sources.append((keyset, config, grid))
        self.ui.success(
            f"Источник добавлен: {Path(file_path).name} / {grid.name.strip()}"
        )
        if self.base_keys:
            base = {k.casefold() for k in self.base_keys}
            covered = sum(1 for k in keyset.keys if k.casefold() in base)
            new = len(keyset.keys) - covered
            self.ui.info(
                f"Из названий первого файла найдено: {covered} из "
                f"{len(self.base_keys)}; новых названий добавлено: {new}"
            )
        return None

    # ------------------------------------------------------------------
    # выбор характеристик

    def _select_characteristics(self, grid, row0, col0, orientation):
        """Мультивыбор столбцов (строк) характеристик + названия столбцов."""
        vertical = orientation is Orientation.VERTICAL
        indices = range(grid.ncols) if vertical else range(grid.nrows)
        items: list[tuple[int, str]] = []
        for i in indices:
            if i == (col0 if vertical else row0):
                continue  # ключевой столбец/строка не предлагаем
            header = self._guess_header(grid, i, row0, col0, orientation)
            sample = format_cell(grid.value(row0, i) if vertical else grid.value(i, col0))
            label = col_letter(i) if vertical else str(i + 1)
            items.append((i, f"{label} — {header or 'без заголовка'} (напр.: {sample})"))

        what = "столбцы" if vertical else "строки"
        selected = self.ui.multi_select(
            f"Выберите {what} с характеристиками для переноса", [t for _, t in items]
        )
        if isinstance(selected, Command):
            return selected

        chosen = [items[i][0] for i in selected]
        all_titles = [
            title
            for _, config, _ in self.sources
            for title in config.column_titles
        ]
        titles: list[str] = []
        for i in chosen:
            header = self._guess_header(grid, i, row0, col0, orientation)
            label = col_letter(i) if vertical else str(i + 1)
            if header:
                # Заголовок прочитан из таблицы — вопрос пользователю не нужен
                title = self._unique_title(header, all_titles)
                self.ui.info(f"Графа {label}: название из таблицы — «{title}»")
            else:
                title = self.ui.ask(
                    f"Не удалось прочитать заголовок графы {label} "
                    "— введите её название в итоговом файле",
                    default=f"Графа {label}",
                    allow_build=bool(self.sources),
                )
                if isinstance(title, Command):
                    return title
                title = self._unique_title(title, all_titles)
            all_titles.append(title)
            titles.append(title)
        return chosen, titles

    @staticmethod
    def _unique_title(title: str, existing: list[str]) -> str:
        """Устраняет дубликаты названий столбцов: «X» -> «X (2)» и т.д."""
        if title not in existing:
            return title
        n = 2
        while f"{title} ({n})" in existing:
            n += 1
        return f"{title} ({n})"

    @staticmethod
    def _guess_header(grid, index: int, row0: int, col0: int,
                      orientation: Orientation) -> str:
        """Ближайший непустой текст выше (левее) якоря — предположительный заголовок.

        Пропускает строки нумерации граф («1», «2», «1а») и строки единиц
        измерения («м», «фт», «°»).
        """
        vertical = orientation is Orientation.VERTICAL
        rng = range(row0 - 1, -1, -1) if vertical else range(col0 - 1, -1, -1)
        skipped_unit = ""
        for j in rng:
            value = grid.value(j, index) if vertical else grid.value(index, j)
            if not _looks_like_header(value):
                # Запомним единицу измерения: заголовок может быть разорван
                # на две строки («Sп,» + «м»)
                if value is not None and str(value).strip().casefold() in _UNIT_VALUES:
                    if not skipped_unit:
                        skipped_unit = str(value).strip()
                continue
            if _is_section_row(grid, j, value, vertical):
                continue  # титул документа или строка-разделитель («район 2»)
            text = str(format_cell(value)).replace("\n", " ").strip()
            if skipped_unit:
                # Восстанавливаем единицу измерения в названии:
                # «Sп,» + «м» -> «Sп, м»; «Высота» + «м» -> «Высота, м»
                separator = " " if text.endswith(",") else ", "
                text = f"{text}{separator}{skipped_unit}"
            return text
        return ""

    # ------------------------------------------------------------------
    # сборка итогового файла

    def _step_build(self):
        self.ui.screen("Сборка итогового файла")
        aggregator = Aggregator(self.key_title)
        for keyset, config, grid in self.sources:
            aggregator.add_source(keyset, config, grid)
        headers, rows = aggregator.build_table()

        self.ui.info(
            f"Источников: {len(self.sources)}; названий объектов: {len(rows)}; "
            f"столбцов в итоге: {len(headers)}"
        )
        for name, covered, total in aggregator.coverage_stats():
            self.ui.info(f"  • {name}: покрыто {covered} из {total} названий")
        for warning in aggregator.duplicate_warnings():
            self.ui.warn(warning)

        output_dir = _program_dir()
        default_path = str(output_dir / _DEFAULT_OUTPUT_NAME)
        self.ui.info(
            "Нажмите Enter, чтобы сохранить файл рядом с программой, "
            "или введите свой путь с именем файла. Расширение .xlsx добавится автоматически."
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
            if not out.is_absolute():
                out = output_dir / out
            if out.suffix.lower() != ".xlsx":
                out = out.with_suffix(".xlsx")

            if out.exists():
                overwrite = self.ui.confirm(
                    f"Файл «{out.name}» уже существует — перезаписать?", default=False
                )
                if overwrite is not True:
                    if isinstance(overwrite, Command):
                        return overwrite
                    continue

            try:
                export(out, headers, rows, self.key_title)
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
        self.ui.success(f"Строк: {len(rows)}, столбцов: {len(headers)}")

        open_it = self.ui.confirm("Открыть собранный файл?", default=False)
        if open_it is True:
            try:
                os.startfile(str(out))  # noqa: S606 — штатное действие Windows
            except OSError:
                self.ui.warn("Не удалось открыть файл автоматически — откройте его вручную.")
        elif isinstance(open_it, Command):
            return open_it

        again = self.ui.confirm(
            "Собрать ещё один файл с другими настройками?", default=False
        )
        if again is True:
            return "restart"
        if isinstance(again, Command):
            return again
        return None
