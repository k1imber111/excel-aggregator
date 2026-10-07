"""Распознавание таблицы: шапка, начало данных, названия столбцов, соединение."""

from pathlib import Path

import pytest

from excel_aggregator.core.layout import column_paths, detect_layout, layout_from
from excel_aggregator.core.matcher import suggest_join
from excel_aggregator.core.models import SheetGrid
from excel_aggregator.core.reader import load_workbook

SAMPLES = Path(__file__).resolve().parents[2] / "Примеры таблиц"
FILE1 = SAMPLES / "Каталог Сочи 2024 МОС ФАП-262.xls"
FILE2 = SAMPLES / "Каталог Сочи 2024 ПЗ-90.11 00.00.xls"
FILE3 = SAMPLES / "Превышающие препятствия 2024.xls"
SHEET1, SHEET2, SHEET3 = "Каталог 2024", "Препятствия ПЗ-90.11_2024 00.00", "Лист1"

needs_samples = pytest.mark.skipif(
    not all(f.is_file() for f in (FILE1, FILE2, FILE3)), reason="Примеры таблиц недоступны"
)


def _paths(grid: SheetGrid) -> dict[int, list[str]]:
    return column_paths(grid, detect_layout(grid))


# --- образцы -----------------------------------------------------------------

@needs_samples
@pytest.mark.parametrize("path, sheet, header, numbering, data", [
    (FILE1, SHEET1, [1, 2, 3, 4], 5, 6),
    (FILE2, SHEET2, [1, 2, 3, 5], None, 7),     # строка 4 пуста, строка 6 — «район 2»
    (FILE3, SHEET3, [4, 6, 7], 8, 9),           # строки 1–2 — подпись таблицы
])
def test_samples_layout(path, sheet, header, numbering, data):
    layout = detect_layout(load_workbook(path).sheet(sheet))
    assert [r + 1 for r in layout.header_rows] == header
    assert (layout.numbering_row + 1 if layout.numbering_row is not None else None) == numbering
    assert layout.data_start + 1 == data


@needs_samples
def test_sample_title_is_not_a_header_level():
    layout = detect_layout(load_workbook(FILE3).sheet(SHEET3))
    assert layout.title.startswith("Препятствия")


@needs_samples
def test_working_sheets_have_no_table():
    book = load_workbook(FILE1)
    found = [s.name.strip() for s in book.sheets if detect_layout(s) is not None]
    assert found == [SHEET1]


@needs_samples
def test_sample_titles_follow_merged_hierarchy():
    paths = _paths(load_workbook(FILE1).sheet(SHEET1))
    assert paths[0] == ["№ препятствия"]
    assert paths[6] == ["Прямоугольные координаты, м", "ИВПП 06/24", "МКпос=058º", "X"]
    assert paths[7][-1] == "Y" and paths[8][2] == "МКпос=238º"
    assert paths[4][-1] == "Ап, град"            # единица измерения дописана к названию
    assert paths[14] == ["Абсолютная высота препятствия, м"]
    assert paths[15] == ["Абсолютная высота препятствия, фут"]


@needs_samples
def test_sample_unmerged_group_and_split_phrase():
    paths = _paths(load_workbook(FILE2).sheet(SHEET2))
    # «Высота геодези-ческая» стоит только над «м»; «фут» берёт название слева
    assert paths[11][-1] == "Высота геодези-ческая, м"
    assert paths[12][-1] == "Высота геодези-ческая, фут"
    # фраза в двух строках шапки — один уровень
    assert paths[15] == ["Точность определяемых характеристик", "Широ-та, м"]
    # «0» слева от знака минут — градусы; буква полушария единицу не наследует
    assert paths[3][-1] == "Широта, °"
    assert paths[6][-1] == "Широта"


@needs_samples
def test_join_column_found_by_overlap():
    grids = [load_workbook(f).sheet(s) for f, s in ((FILE1, SHEET1), (FILE2, SHEET2), (FILE3, SHEET3))]
    tables = [(g, detect_layout(g).data_start) for g in grids]
    assert suggest_join(tables) == [0, 0, 0]


# --- синтетика: обычные и «ломаные» шапки -------------------------------------

def test_one_row_header():
    grid = SheetGrid("s", [["Название", "Высота", "Ширина"], ["A", 1, 10], ["B", 2, 20], ["C", 3, 30]])
    layout = detect_layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0], 1)
    assert column_paths(grid, layout) == {0: ["Название"], 1: ["Высота"], 2: ["Ширина"]}


def test_two_column_table():
    grid = SheetGrid("s", [["Название", "Высота"], ["A", 1], ["B", 2], ["C", 3]])
    assert _paths(grid) == {0: ["Название"], 1: ["Высота"]}


def test_table_without_header():
    grid = SheetGrid("s", [["A", 1, 10], ["B", 2, 20], ["C", 3, 30]])
    layout = detect_layout(grid)
    assert (layout.header_rows, layout.data_start) == ([], 0)
    assert column_paths(grid, layout)[1] == ["Столбец B"]


def test_text_only_table_first_row_is_header():
    grid = SheetGrid("s", [["Код", "Имя"], ["a", "x"], ["b", "y"], ["c", "z"]])
    layout = detect_layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0], 1)


def test_caption_above_header():
    grid = SheetGrid("s", [
        ["Таблица 5", None, None],
        ["Название", "Высота", "Ширина"],
        ["A", 1, 10], ["B", 2, 20], ["C", 3, 30],
    ])
    layout = detect_layout(grid)
    assert layout.title == "Таблица 5"
    assert (layout.header_rows, layout.data_start) == ([1], 2)


def test_numbering_row_is_not_data():
    grid = SheetGrid("s", [
        ["Название", "Высота", "Ширина"],
        [1, 2, 3],
        ["A", 5, 10], ["B", 6, 20], ["C", 7, 30],
    ])
    layout = detect_layout(grid)
    assert (layout.header_rows, layout.numbering_row, layout.data_start) == ([0], 1, 2)


def test_merged_group_header():
    grid = SheetGrid(
        "s",
        [["Название", "Координаты", None], [None, "X", "Y"], ["A", 1, 2], ["B", 3, 4], ["C", 5, 6]],
        merged_ranges=[(0, 2, 0, 1), (0, 1, 1, 3)],
    )
    assert _paths(grid) == {0: ["Название"], 1: ["Координаты", "X"], 2: ["Координаты", "Y"]}


def test_unmerged_group_with_units():
    grid = SheetGrid("s", [
        ["Название", "Высота", None],
        [None, "м", "фут"],
        ["A", 1, 3], ["B", 2, 6], ["C", 3, 9],
    ])
    paths = _paths(grid)
    assert paths[1] == ["Высота, м"]
    assert paths[2] == ["Высота, фут"]


def test_column_without_own_header_does_not_steal_neighbour_title():
    # У третьего столбца нет ни названия, ни единицы — чужое название он не берёт
    grid = SheetGrid("s", [["Название", "Высота", None], ["A", 1, "x"], ["B", 2, "y"], ["C", 3, "z"]])
    assert _paths(grid)[2] == ["Столбец C"]


def test_manual_data_start():
    grid = SheetGrid("s", [["Название", "Высота"], ["A", 1], ["B", 2]])
    layout = layout_from(grid, 1)
    assert (layout.header_rows, layout.data_start) == ([0], 1)


def test_sheet_without_table():
    assert detect_layout(SheetGrid("s", [["одна ячейка"]])) is None
    assert detect_layout(SheetGrid("s", [])) is None


def test_rotated_table_reads_after_transposition():
    # Названия объектов идут в строку: после поворота — обычная таблица
    grid = SheetGrid("s", [["Название", "A", "B", "C"], ["Высота", 1, 2, 3], ["Ширина", 10, 20, 30]])
    turned = grid.transposed()
    assert _paths(turned) == {0: ["Название"], 1: ["Высота"], 2: ["Ширина"]}
    assert turned.value(1, 0) == "A" and turned.value(3, 2) == 30


def test_row_of_zeros_is_data_not_numbering():
    # Первая строка данных из одних нулей не должна пропасть как «нумерация граф»
    grid = SheetGrid("s", [
        ["Код", "П1", "П2", "П3"],
        ["obj-0", 0, 0, 0],
        ["obj-1", 0.5, 1.5, 2.5],
        ["obj-2", 1.0, 3.0, 5.0],
    ])
    layout = detect_layout(grid)
    assert (layout.numbering_row, layout.data_start) == (None, 1)


def test_sparse_first_rows_are_still_data():
    # У первых объектов заполнена часть граф — они не должны уйти в «шапку»
    grid = SheetGrid("s", [
        ["Код", "П1", "П2", "П3", "П4"],
        ["a", 1, None, None, None],
        ["b", None, 2, None, None],
        ["c", 1, 2, 3, 4],
        ["d", 5, 6, 7, 8],
        ["e", 9, 10, 11, 12],
    ])
    layout = detect_layout(grid)
    assert (layout.header_rows, layout.data_start) == ([0], 1)


def test_name_column_wins_over_measurements_despite_duplicates():
    # В названиях есть повтор, в суммах повторов нет — названия всё равно столбец A
    grid = SheetGrid("s", [
        ["Объект", "Сумма"],
        ["Мачта 1", 2.5], ["Мачта 2", 3.5], ["Мачта 2", 4.25], ["Мачта 3", 5.5], ["Мачта 4", 6.75],
    ])
    assert suggest_join([(grid, 1)]) == [0]


def test_row_numbers_are_not_object_names():
    grid = SheetGrid("s", [
        ["№ п/п", "Название", "Высота"],
        [1, "Труба", 10.5], [2, "Мачта", 20.5], [3, "Вышка", 30.5], [4, "Опора", 40.5],
    ])
    assert suggest_join([(grid, 1)]) == [1]


def test_coincidental_numbers_do_not_join_tables():
    # Объекты разные (a… и b…); совпадают только измерения 1, 2, 3 — это не соединение
    first = SheetGrid("a", [["Код", "Вес"], ["a1", 1], ["a2", 2], ["a3", 3]])
    second = SheetGrid("b", [["Код", "Рост"], ["b1", 1], ["b2", 2], ["b3", 3]])
    titles = [{0: "Код", 1: "Вес"}, {0: "Код", 1: "Рост"}]
    assert suggest_join([(first, 1), (second, 1)], titles) == [0, 0]


def test_join_column_can_differ_between_tables():
    first = SheetGrid("a", [["Код", "Высота"], ["k1", 1], ["k2", 2], ["k3", 3]])
    second = SheetGrid("b", [["Вес", "Код"], [5, "k2"], [6, "k3"], [7, "k9"]])
    assert suggest_join([(first, 1), (second, 1)]) == [0, 1]
