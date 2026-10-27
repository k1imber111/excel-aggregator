"""Тесты общих виджетов ввода."""

import io

import pytest
from rich.console import Console

from excel_aggregator.ui.nav import Command
from excel_aggregator.ui.widgets import UI


def test_parse_numbers_supports_ranges() -> None:
    assert UI._parse_numbers("2-5, 6, 8, 12-16", 20) == [
        1, 2, 3, 4, 5, 7, 11, 12, 13, 14, 15
    ]


@pytest.mark.parametrize("text", ["5-2", "0", "11", "x-y"])
def test_parse_numbers_rejects_bad_input(text: str) -> None:
    with pytest.raises(ValueError):
        UI._parse_numbers(text, 10)


def _ui_with_input(answers: list[str]) -> tuple[UI, list[str]]:
    """UI с подставленным вводом; вывод собирается в буфер."""
    ui = UI()
    ui.console = Console(file=io.StringIO(), width=120)
    queue = list(answers)
    ui.console.input = lambda prompt="": queue.pop(0)
    return ui, queue


@pytest.mark.parametrize("word", ["п", "П", " далее ", "ПРОПУСТИТЬ", "g", "skip"])
def test_multi_select_skip_command(word: str) -> None:
    ui, _ = _ui_with_input([word])
    assert ui.multi_select("t", ["a", "b"], default_all=True, allow_skip=True) is Command.SKIP


@pytest.mark.parametrize("word", ["п", "далее"])
def test_multi_select_skip_is_not_a_number_when_disabled(word: str) -> None:
    # без allow_skip слово не команда: предупреждение и повторный ввод
    ui, queue = _ui_with_input([word, "2"])
    assert ui.multi_select("t", ["a", "b"], allow_skip=False) == [1]
    assert not queue


@pytest.mark.parametrize("answer, expected", [
    ("", [0, 1, 2]), ("все", [0, 1, 2]), ("1,3", [0, 2]), ("кроме: 2", [0, 2]),
])
def test_multi_select_old_inputs_work_with_skip_enabled(answer: str, expected: list[int]) -> None:
    ui, _ = _ui_with_input([answer])
    result = ui.multi_select("t", ["a", "b", "c"], default_all=True, allow_skip=True)
    assert result == expected


def test_multi_select_navigation_still_works_with_skip_enabled() -> None:
    ui, _ = _ui_with_input(["н", "с", "о"])
    for expected in (Command.BACK, Command.BUILD, Command.CANCEL):
        assert ui.multi_select("t", ["a"], allow_build=True, allow_skip=True) is expected


def test_skip_word_does_not_leak_into_other_prompts() -> None:
    from excel_aggregator.ui.nav import parse_command
    assert parse_command("п") is None and parse_command("далее") is None
