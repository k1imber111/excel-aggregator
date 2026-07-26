"""Unit-тесты: resolve_input_path."""

import pytest

from excel_aggregator.core.errors import PathNotFoundError
from excel_aggregator.core.paths import resolve_input_path


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
