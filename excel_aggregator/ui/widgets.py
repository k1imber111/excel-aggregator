"""Виджеты консольного интерфейса на rich: шапка, сообщения, превью, ввод."""

from __future__ import annotations

import os

from rich.console import Console
from rich.markup import escape
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from ..core.normalize import format_cell
from .nav import HINT, HINT_NO_BUILD, HINT_WORDS, Command, parse_command

_TITLE = "АГРЕГАТОР EXCEL-ТАБЛИЦ"
_AUTHOR = "Автор: Иван Журавлев"
_PREVIEW_ROWS = 8
_PREVIEW_COLS = 8
_CELL_MAX = 24


class UI:
    """Единая обёртка над rich: единый стиль панелей, ввода и сообщений."""

    def __init__(self) -> None:
        self.console = Console()

    # --- каркас экрана -------------------------------------------------

    def clear(self) -> None:
        """Очистка экрана (в скриптовом режиме отключена)."""
        if os.environ.get("AGGREGATOR_SCRIPTED"):
            return
        self.console.clear()

    def header(self, subtitle: str = "") -> None:
        """Шапка программы с автором."""
        text = Text()
        text.append(f"  {_TITLE}  \n", style="bold white")
        text.append(f"  {_AUTHOR}  ", style="italic cyan")
        if subtitle:
            text.append(f"\n  {subtitle}  ", style="dim")
        self.console.print(Panel(text, style="cyan", padding=(1, 4)))

    def screen(self, subtitle: str = "") -> None:
        """Новый экран: очистка + шапка."""
        self.clear()
        self.header(subtitle)

    # --- сообщения ------------------------------------------------------

    # Текст сообщений печатается буквально (markup=False): в нём бывают имена
    # файлов и значения ячеек, а «[final]» или «[/]» rich принял бы за разметку.

    def info(self, msg: str) -> None:
        self.console.print(msg, style="white", markup=False)

    def success(self, msg: str) -> None:
        self.console.print(f"✔ {msg}", style="bold green", markup=False)

    def warn(self, msg: str) -> None:
        self.console.print(f"! {msg}", style="bold yellow", markup=False)

    def error(self, msg: str) -> None:
        self.console.print(f"✖ {msg}", style="bold red", markup=False)

    def pause(self) -> None:
        """Ждёт Enter, чтобы сообщения не стёрла очистка следующего экрана.

        В скриптовом режиме отключена — как и clear().
        """
        if os.environ.get("AGGREGATOR_SCRIPTED"):
            return
        self.console.input("[dim]Нажмите Enter, чтобы продолжить...[/]")

    def hint(self, allow_build: bool = False) -> None:
        """Строка-подсказка с навигационными командами."""
        self.console.print(f"[dim]{HINT if allow_build else HINT_NO_BUILD}[/]")

    # --- превью ----------------------------------------------------------

    def show_preview(self, grid, title: str = "Предпросмотр") -> None:
        """Первые строки листа в виде таблицы с буквами столбцов."""
        rows = grid.preview_rows(_PREVIEW_ROWS)
        ncols = min(grid.ncols, _PREVIEW_COLS)
        from ..core.models import col_letter

        table = Table(title=escape(title), title_justify="left", show_lines=False)
        for c in range(ncols):
            table.add_column(col_letter(c), overflow="fold", max_width=_CELL_MAX)
        for row in rows:
            table.add_row(*[escape(str(format_cell(v))[:_CELL_MAX]) for v in row[:ncols]])
        self.console.print(table)
        note = []
        if grid.nrows > _PREVIEW_ROWS:
            note.append(f"строк всего: {grid.nrows}")
        if grid.ncols > _PREVIEW_COLS:
            note.append(f"столбцов всего: {grid.ncols}")
        if note:
            self.console.print(f"[dim](показаны первые строки; {', '.join(note)})[/]")

    # --- ввод ------------------------------------------------------------

    def ask(self, prompt: str, default: str | None = None,
            allow_build: bool = False, display_default: str | None = None,
            short_commands: bool = True) -> str | Command:
        """Текстовый ввод. Пустой ввод -> default (если есть).

        Навигационные команды возвращаются как Command. short_commands=False —
        принимаются только команды-слова (для ввода названий: «Y», «Н»).
        """
        shown_default = display_default if display_default is not None else default
        suffix = f" [{shown_default}]" if shown_default else ""
        if not short_commands:
            self.console.print(f"[dim]{HINT_WORDS}[/]")
        while True:
            raw = self.console.input(f"[bold cyan]{escape(prompt + suffix)}:[/] ").strip()
            cmd = parse_command(raw, short=short_commands)
            if cmd is not None:
                if cmd is Command.BUILD and not allow_build:
                    self.warn("Сначала настройте хотя бы один источник данных.")
                    continue
                return cmd
            if not raw:
                if default is not None:
                    return default
                self.warn("Введите значение (или команду навигации).")
                continue
            return raw

    def choose(self, title: str, options: list[str],
               allow_back: bool = True, allow_build: bool = False) -> int | Command:
        """Меню с нумерованными вариантами -> индекс (0-based) или Command."""
        self.console.print(f"[bold]{escape(title)}[/]")
        for i, opt in enumerate(options, start=1):
            self.console.print(f"  [cyan]{i}[/]. {escape(opt)}")
        if allow_back or allow_build:
            self.hint(allow_build=allow_build)
        while True:
            raw = self.console.input("[bold cyan]Ваш выбор:[/] ").strip()
            cmd = parse_command(raw)
            if cmd is not None:
                if cmd is Command.BACK and not allow_back:
                    self.warn("Здесь нет возврата назад.")
                    continue
                if cmd is Command.BUILD and not allow_build:
                    self.warn("Сначала настройте хотя бы один источник данных.")
                    continue
                return cmd
            try:
                num = int(raw)
            except ValueError:
                self.warn("Введите номер варианта.")
                continue
            if 1 <= num <= len(options):
                return num - 1
            self.warn(f"Варианта с номером {num} нет — выберите от 1 до {len(options)}.")

    def multi_select(self, title: str, items: list[str]) -> list[int] | Command:
        """Мультивыбор номеров: «1,3» / «1-3» / «все» / «кроме: 2,4»."""
        self.console.print(f"[bold]{escape(title)}[/]")
        for i, item in enumerate(items, start=1):
            self.console.print(f"  [cyan]{i}[/]. {escape(item)}")
        self.console.print(
            "[dim]Можно выбрать: 2,4; диапазоны: 2-5, 6, 8, 12-16; "
            "все пункты: «все»; всё кроме указанных: «кроме: 2,4» или "
            "«кроме: 2-5, 8».[/]"
        )
        self.hint(allow_build=False)
        while True:
            raw = self.console.input("[bold cyan]Ваш выбор:[/] ").strip()
            cmd = parse_command(raw)
            if cmd is not None:
                if cmd is Command.BUILD:
                    self.warn("Сначала настройте хотя бы один источник данных.")
                    continue
                return cmd
            lowered = raw.casefold()
            try:
                if lowered in ("все", "всё", "all", "*"):
                    return list(range(len(items)))
                if lowered.startswith("кроме"):
                    excluded_text = lowered.split(":", 1)[1] if ":" in lowered else lowered[5:]
                    excluded = self._parse_numbers(excluded_text, len(items))
                    result = [i for i in range(len(items)) if i not in excluded]
                else:
                    result = self._parse_numbers(raw, len(items))
            except ValueError as exc:
                self.warn(str(exc))
                continue
            if not result:
                self.warn("Ничего не выбрано — укажите хотя бы один номер.")
                continue
            return result

    @staticmethod
    def _parse_numbers(text: str, count: int) -> list[int]:
        """«1,3,5» / «1-3,5» -> 0-based индексы; проверка диапазона."""
        result: list[int] = []
        for chunk in text.replace(";", ",").split(","):
            chunk = chunk.strip()
            if not chunk:
                continue
            if "-" in chunk:
                left, right = (part.strip() for part in chunk.split("-", 1))
                try:
                    start, end = int(left), int(right)
                except ValueError:
                    raise ValueError(
                        f"«{chunk}» — не диапазон. Пример диапазона: 2-5."
                    )
                if start > end:
                    raise ValueError(f"Диапазон «{chunk}» указан наоборот.")
                numbers = range(start, end + 1)
            else:
                try:
                    numbers = range(int(chunk), int(chunk) + 1)
                except ValueError:
                    raise ValueError(
                        f"«{chunk}» — не номер. Пример: 2,4 или 2-5, 8."
                    )
            for num in numbers:
                if not 1 <= num <= count:
                    raise ValueError(f"Номера {num} нет в списке (доступно: 1–{count}).")
                if num - 1 not in result:
                    result.append(num - 1)
        return result

    def confirm(self, question: str, default: bool = True) -> bool | Command:
        """Вопрос да/нет. Возвращает bool или Command."""
        mark = "Д/н" if default else "д/Н"
        while True:
            raw = self.console.input(f"[bold cyan]{escape(question)} [{mark}]:[/] ").strip()
            if not raw:
                return default
            lowered = raw.casefold()
            # Сначала да/нет: «н» здесь — это «нет», а не команда «назад»
            if lowered in ("д", "да", "y", "yes"):
                return True
            if lowered in ("н", "нет", "n", "no"):
                return False
            cmd = parse_command(raw)
            if cmd is not None:
                return cmd
            self.warn("Ответьте «д» (да) или «н» (нет).")
