"""Навигационные команды, доступные на каждом шаге настройки.

Вместо экранных кнопок (это консоль) пользователь вводит короткие команды:
«н» — назад, «с» — собрать файл, «о» — отмена (выход в главное меню).
"""

from __future__ import annotations

from enum import Enum


class Command(Enum):
    """Специальные команды навигации."""

    BACK = "back"      # ← Назад (к предыдущему шагу)
    BUILD = "build"    # Собрать файл
    CANCEL = "cancel"  # Отмена (в главное меню)
    SKIP = "skip"      # Пропустить текущий лист (только на шаге выбора столбцов)


# Варианты ввода для каждой команды (русская и латинская раскладки)
_COMMAND_KEYS = {
    "н": Command.BACK, "y": Command.BACK, "назад": Command.BACK,
    "с": Command.BUILD, "c": Command.BUILD, "собрать": Command.BUILD,
    "о": Command.CANCEL, "j": Command.CANCEL, "отмена": Command.CANCEL,
}
# «Пропустить» не входит в _COMMAND_KEYS: она осмысленна только там, где
# её ждут (multi_select), и не должна «просачиваться» в остальные вопросы.
_SKIP_KEYS = {"п", "g", "пропустить", "далее", "skip"}

HINT = "н — назад | с — собрать файл | о — отмена"
HINT_NO_BUILD = "н — назад | о — отмена"
HINT_SKIP = "н — назад | п — пропустить лист | с — собрать файл | о — отмена"


def parse_command(text: str) -> Command | None:
    """Распознаёт навигационную команду; обычный ввод -> None."""
    return _COMMAND_KEYS.get((text or "").strip().casefold())


def is_skip(text: str) -> bool:
    """Ввод — команда «пропустить лист»."""
    return (text or "").strip().casefold() in _SKIP_KEYS
