"""Unit-тесты: normalize_key / keys_equal / format_cell."""

from excel_aggregator.core.normalize import (
    fold_lookalikes,
    format_cell,
    keys_equal,
    normalize_key,
)


def test_normalize_numbers():
    assert normalize_key(2) == "2"
    assert normalize_key(2.0) == "2"
    assert normalize_key(2.5) == "2.5"


def test_normalize_strings():
    assert normalize_key("2") == "2"
    assert normalize_key(" 22 ") == "22"
    assert normalize_key("2.0") == "2"
    assert normalize_key("\xa022\xa0") == "22"   # неразрывные пробелы
    assert normalize_key(" 2 2 ") == "2 2"       # пробелы внутри сохраняются


def test_normalize_whitespace_and_decimal_comma():
    assert normalize_key("Объект  1") == "Объект 1"      # повторные пробелы
    assert normalize_key("Дом\n№1") == "Дом №1"          # перенос строки в ячейке
    assert keys_equal(4.61, "4,61")                       # десятичная запятая
    assert not keys_equal("1.10", "1.1")                  # номера пунктов — разные ключи
    assert normalize_key(12345678901234567) == "12345678901234567"  # без потери точности


def test_normalize_leading_zeros_preserved():
    assert normalize_key("040") == "040"
    assert normalize_key("04.61") == "04.61"


def test_normalize_empty():
    assert normalize_key(None) is None
    assert normalize_key("") is None
    assert normalize_key("   ") is None
    assert normalize_key("\xa0") is None


def test_normalize_text_keeps_case():
    assert normalize_key("АбВ") == "АбВ"


def test_keys_equal():
    assert keys_equal(2, "2")
    assert keys_equal(2.0, "2.0")
    assert keys_equal(" 22 ", 22)
    assert keys_equal("Абв", "АБВ")     # casefold для текста
    assert not keys_equal("040", 40)    # ведущие нули — другой ключ
    assert not keys_equal(None, "2")
    assert not keys_equal(None, None)


def test_format_cell():
    assert format_cell(2.0) == 2        # float с целым значением -> int для превью
    assert format_cell(2.5) == 2.5
    assert format_cell(None) == ""
    assert format_cell("040") == "040"


def test_fold_lookalikes():
    assert fold_lookalikes("A1") == fold_lookalikes("А1")      # латинская и русская А
    assert fold_lookalikes("T14A") == fold_lookalikes("Т14А")
    assert fold_lookalikes("A1") != fold_lookalikes("B1")
