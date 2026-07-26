"""Нормализация путей, введённых пользователем."""

from __future__ import annotations

import os
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

    candidates = [path]
    if not path.suffix:
        candidates = [path.with_suffix(ext) for ext in _DEFAULT_EXTENSIONS]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    tried = ", ".join(str(c) for c in candidates)
    raise PathNotFoundError(f"Файл не найден: {tried}")
