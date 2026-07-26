"""Тесты общих виджетов ввода."""

import pytest

from excel_aggregator.ui.widgets import UI


def test_parse_numbers_supports_ranges() -> None:
    assert UI._parse_numbers("2-5, 6, 8, 12-16", 20) == [
        1, 2, 3, 4, 5, 7, 11, 12, 13, 14, 15
    ]


@pytest.mark.parametrize("text", ["5-2", "0", "11", "x-y"])
def test_parse_numbers_rejects_bad_input(text: str) -> None:
    with pytest.raises(ValueError):
        UI._parse_numbers(text, 10)
