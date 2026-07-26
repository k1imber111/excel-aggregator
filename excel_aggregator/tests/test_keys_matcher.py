"""Unit-тесты: SheetGrid, extract_keys, find_key_candidates на синтетических данных."""

import pytest

from excel_aggregator.core.errors import EmptyKeySetError
from excel_aggregator.core.keys import extract_keys
from excel_aggregator.core.matcher import find_key_candidates
from excel_aggregator.core.models import Orientation, SheetGrid


def make_grid():
    # столбец 0 — ключи с пропусками и дубликатом, столбец 1 — мусор,
    # столбец 2 — частичное совпадение ключей
    return SheetGrid("Лист", [
        ["шапка", "x", "y"],
        [1, "foo", 10],
        [None, "bar", None],
        [2, "baz", 11],
        [3, "qux", 12],
        [2, "dup", 13],      # дубликат ключа "2"
        [4, None, None],
        [5, None, None],
    ])


def test_sheetgrid_basics():
    grid = make_grid()
    assert grid.nrows == 8
    assert grid.ncols == 3
    assert grid.value(1, 0) == 1
    assert grid.value(99, 99) is None
    assert grid.row(3) == [2, "baz", 11]
    assert grid.col(0)[0] == "шапка"


def test_sheetgrid_merged_expanded():
    grid = SheetGrid("Лист", [["A", None], [None, None]], merged_ranges=[(0, 2, 0, 2)])
    assert grid.value(0, 0) == "A"
    assert grid.value(1, 1) == "A"   # merged-диапазон развёрнут


def test_sheetgrid_preview_rows():
    grid = SheetGrid("Лист", [[None, None], [1, 2], [None, None], [3, 4], [5, 6]])
    assert grid.preview_rows(2) == [[1, 2], [3, 4]]


def test_extract_keys_vertical():
    grid = make_grid()
    ks = extract_keys(grid, 1, 0, Orientation.VERTICAL)
    assert ks.keys == ["1", "2", "3", "4", "5"]   # пропуски пропущены, дубликат убран
    assert ks.duplicates == ["2"]
    assert ks.raw_positions["2"] == 3             # первое вхождение


def test_extract_keys_horizontal():
    grid = SheetGrid("Лист", [["шапка", 1, None, 2, 3], [0, 0, 0, 0, 0]])
    ks = extract_keys(grid, 0, 1, Orientation.HORIZONTAL)
    assert ks.keys == ["1", "2", "3"]
    assert ks.raw_positions["3"] == 4


def test_extract_keys_empty():
    grid = SheetGrid("Лист", [[None, ""], [None, "  "], [None, None]])
    with pytest.raises(EmptyKeySetError):
        extract_keys(grid, 0, 0, Orientation.VERTICAL)


def test_matcher_best_candidate():
    grid = make_grid()
    candidates = find_key_candidates(grid, ["1", "2", "3", "4", "5"], Orientation.VERTICAL)
    assert candidates
    best = candidates[0]
    assert best.index == 0
    assert best.match_count == 5          # различные ключи (дубликат "2" не накручивает)
    assert best.coverage == pytest.approx(1.0)
    assert best.first_row0 == 1
    assert len(best.sample) <= 10


def test_matcher_filters_garbage():
    # столбец с одним случайным совпадением (< 2) отсеивается
    grid = SheetGrid("Лист", [[7, "a"], ["x", 1], ["y", "z"]])
    candidates = find_key_candidates(grid, ["1", "2", "3"], Orientation.VERTICAL)
    assert candidates == []               # пустой результат — не ошибка


def test_matcher_horizontal():
    hgrid = SheetGrid("Лист", [["х", 1, 2, 3], ["y", "a", "b", "c"]])
    candidates = find_key_candidates(hgrid, ["1", "2", "3"], Orientation.HORIZONTAL)
    assert candidates[0].index == 0
    assert candidates[0].match_count == 3
    assert candidates[0].first_row0 == 1
