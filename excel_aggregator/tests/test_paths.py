"""Unit-тесты: resolve_input_path."""

import pytest

from excel_aggregator.core.errors import PathNotFoundError
from excel_aggregator.core.paths import expand_inputs, resolve_input_path


@pytest.fixture
def sample_file(tmp_path):
    f = tmp_path / "таблица данных.xls"
    f.write_bytes(b"\xd0\xcf\x11\xe0")  # достаточно существования файла
    return f


def test_plain_path(sample_file):
    assert resolve_input_path(str(sample_file)) == sample_file


def test_quoted_paths(sample_file):
    assert resolve_input_path(f'"{sample_file}"') == sample_file
    assert resolve_input_path(f"'{sample_file}'") == sample_file
    assert resolve_input_path(f"«{sample_file}»") == sample_file
    assert resolve_input_path(f'  "{sample_file}"  ') == sample_file


def test_missing_extension(sample_file):
    assert resolve_input_path(str(sample_file.with_suffix(""))) == sample_file


def test_relative_to_base_dir(sample_file):
    base = sample_file.parent
    assert resolve_input_path(sample_file.name, base_dir=base) == sample_file
    # без расширения, относительный
    assert resolve_input_path(sample_file.stem, base_dir=base) == sample_file


def test_not_found(tmp_path):
    with pytest.raises(PathNotFoundError):
        resolve_input_path("нет такого файла.xls", base_dir=tmp_path)


def test_empty_input():
    with pytest.raises(PathNotFoundError):
        resolve_input_path('""')


def _touch(path):
    path.write_bytes(b"\xd0\xcf\x11\xe0")
    return path


def test_expand_folder(tmp_path):
    a = _touch(tmp_path / "а таблица.xls")
    b = _touch(tmp_path / "б таблица.xlsx")
    _touch(tmp_path / "~$б таблица.xlsx")          # временный файл открытой книги
    (tmp_path / "заметки.txt").write_text("x", encoding="utf-8")
    assert expand_inputs(str(tmp_path)) == [a, b]


def test_expand_several_quoted_paths(tmp_path):
    a = _touch(tmp_path / "а таблица.xls")
    b = _touch(tmp_path / "б.xlsx")
    assert expand_inputs(f'"{a}" "{b}"') == [a, b]


def test_expand_single_path_with_spaces(sample_file):
    assert expand_inputs(str(sample_file)) == [sample_file]


def test_expand_empty_folder(tmp_path):
    with pytest.raises(PathNotFoundError):
        expand_inputs(str(tmp_path))
