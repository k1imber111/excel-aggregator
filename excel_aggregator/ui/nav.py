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


# Варианты ввода для каждой команды (русская и латинская раскладки)
_COMMAND_KEYS = {
    "н": Command.BACK, "y": Command.BACK, "назад": Command.BACK,
    "с": Command.BUILD, "c": Command.BUILD, "собрать": Command.BUILD,
    "о": Command.CANCEL, "j": Command.CANCEL, "отмена": Command.CANCEL,
}

HINT = "н — назад | с — собрать файл | о — отмена"
HINT_NO_BUILD = "н — назад | о — отмена"


def parse_command(text: str) -> Command | None:
    """Распознаёт навигационную команду; обычный ввод -> None."""
    return _COMMAND_KEYS.get((text or "").strip().casefold())
