"""Unit-тесты: normalize_key / keys_equal / format_cell."""

from excel_aggregator.core.normalize import format_cell, keys_equal, normalize_key


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
