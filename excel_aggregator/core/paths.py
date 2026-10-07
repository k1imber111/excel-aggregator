"""Нормализация путей, введённых пользователем."""

from __future__ import annotations

import os
import re
from pathlib import Path

from .errors import PathNotFoundError

# Пары обрамляющих кавычек, которые пользователи часто копируют вместе с путём
_QUOTE_PAIRS = (('"', '"'), ("'", "'"), ("«", "»"))

# Расширения, которые пробуем подставить, если пользователь не указал своё
_DEFAULT_EXTENSIONS = (".xls", ".xlsx")


def _strip_quotes(text: str) -> str:
    """Снимает обрамляющие кавычки и лишние пробелы."""
    s = text.strip()
    changed = True
    while changed and len(s) >= 2:
        changed = False
        for left, right in _QUOTE_PAIRS:
            if s.startswith(left) and s.endswith(right):
                s = s[1:-1].strip()
                changed = True
    return s


def resolve_input_path(raw: str, base_dir: Path | None = None) -> Path:
    """Превращает сырую строку ввода в существующий путь к файлу.

    - снимает обрамляющие кавычки ("...", '...', «...») и пробелы;
    - раскрывает переменные окружения и ~;
    - относительные пути считает от base_dir (или от текущей директории);
    - если расширения нет — пробует добавить .xls / .xlsx;
    - если файл не найден — PathNotFoundError с понятным сообщением.
    """
    cleaned = _strip_quotes(raw or "")
    if not cleaned:
        raise PathNotFoundError("Путь к файлу не указан.")

    expanded = os.path.expandvars(os.path.expanduser(cleaned))
    path = Path(expanded)
    if not path.is_absolute():
        path = (base_dir or Path.cwd()) / path

    # Расширение дописываем к имени целиком: у «данные.2024» suffix — «.2024»,
    # и with_suffix его бы затёр
    candidates = [path]
    if path.name and path.suffix.lower() not in _DEFAULT_EXTENSIONS:
        candidates += [path.with_name(path.name + ext) for ext in _DEFAULT_EXTENSIONS]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    tried = ", ".join(str(c) for c in candidates)
    raise PathNotFoundError(f"Файл не найден: {tried}")


def _expand_one(text: str, base_dir: Path | None) -> list[Path]:
    """Один путь -> файлы: папка раскрывается в свои таблицы Excel."""
    cleaned = _strip_quotes(text)
    path = Path(os.path.expandvars(os.path.expanduser(cleaned))) if cleaned else None
    if path is not None and not path.is_absolute():
        path = (base_dir or Path.cwd()) / path
    if path is not None and path.is_dir():
        files = sorted(
            f for f in path.iterdir()
            if f.is_file()
            and f.suffix.lower() in _DEFAULT_EXTENSIONS
            and not f.name.startswith("~$")  # временные файлы открытых книг
        )
        if not files:
            raise PathNotFoundError(f"В папке «{path}» нет файлов Excel (.xls, .xlsx).")
        return files
    return [resolve_input_path(text, base_dir)]


def expand_inputs(raw: str, base_dir: Path | None = None) -> list[Path]:
    """Строка ввода -> список файлов Excel.

    Принимает путь к файлу, путь к папке (берутся все .xls/.xlsx в ней) или
    несколько путей в одной строке — так их вставляет консоль при
    перетаскивании нескольких файлов: "C:\\a b.xls" "C:\\c.xls".
    """
    try:
        return _expand_one(raw or "", base_dir)
    except PathNotFoundError:
        tokens = re.findall(r'"[^"]+"|\S+', raw or "")
        if len(tokens) < 2:
            raise
        try:
            return [f for token in tokens for f in _expand_one(token, base_dir)]
        except PathNotFoundError:
            pass
        raise  # строка не делится на пути — показываем ошибку для неё целиком
