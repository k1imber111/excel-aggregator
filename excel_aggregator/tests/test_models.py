"""Unit-тесты: parse_coordinate / col_letter."""

import pytest

from excel_aggregator.core.errors import CoordinateParseError
from excel_aggregator.core.models import col_letter, parse_coordinate


@pytest.mark.parametrize("text,expected", [
    ("A6", (5, 0)),
    ("А6", (5, 0)),      # русская А
    ("а6", (5, 0)),      # строчная русская
    ("a6", (5, 0)),      # строчная латинская
    ("B2", (1, 1)),
    ("В2", (1, 1)),      # русская В = столбец B
    ("AA10", (9, 26)),   # многобуквенный столбец
    ("AB1", (0, 27)),
    (" A 6 ", (5, 0)),   # пробелы игнорируются
    ("С1", (0, 2)),      # русская С = столбец C
])
def test_parse_coordinate_ok(text, expected):
    assert parse_coordinate(text) == expected


@pytest.mark.parametrize("text", [
    "",
    "6A",      # цифры перед буквами
    "A",       # нет строки
    "6",       # нет столбца
    "A0",      # строка с 0
    "Д5",      # русская буква без латинского двойника
    "A-3",
    "A1B2",
])
def test_parse_coordinate_garbage(text):
    with pytest.raises(CoordinateParseError):
        parse_coordinate(text)


@pytest.mark.parametrize("idx,expected", [
    (0, "A"),
    (1, "B"),
    (25, "Z"),
    (26, "AA"),
    (27, "AB"),
    (51, "AZ"),
    (52, "BA"),
])
def test_col_letter(idx, expected):
    assert col_letter(idx) == expected


def test_col_letter_roundtrip():
    for idx in (0, 5, 26, 100):
        assert parse_coordinate(col_letter(idx) + "1") == (0, idx)
