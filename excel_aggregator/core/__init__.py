"""Core-слой агрегатора Excel-таблиц (без UI)."""

from .aggregator import Aggregator
from .errors import (
    AppError,
    CoordinateParseError,
    CorruptedFileError,
    EmptyKeySetError,
    EncryptedFileError,
    NoMatchError,
    NotExcelFileError,
    OutputWriteError,
    PathNotFoundError,
)
from .exporter import export
from .keys import KeySet, extract_keys
from .layout import TableLayout, column_paths, detect_layout, layout_from
from .matcher import Candidate, find_key_candidates, suggest_join
from .models import (
    Orientation,
    SheetGrid,
    SourceConfig,
    Workbook,
    col_letter,
    parse_coordinate,
)
from .normalize import fold_lookalikes, format_cell, keys_equal, normalize_key
from .paths import expand_inputs, resolve_input_path
from .reader import load_workbook

__all__ = [
    "Aggregator",
    "AppError",
    "Candidate",
    "CoordinateParseError",
    "CorruptedFileError",
    "EmptyKeySetError",
    "EncryptedFileError",
    "KeySet",
    "NoMatchError",
    "NotExcelFileError",
    "Orientation",
    "OutputWriteError",
    "PathNotFoundError",
    "SheetGrid",
    "SourceConfig",
    "TableLayout",
    "Workbook",
    "col_letter",
    "column_paths",
    "detect_layout",
    "expand_inputs",
    "export",
    "extract_keys",
    "find_key_candidates",
    "fold_lookalikes",
    "format_cell",
    "keys_equal",
    "layout_from",
    "load_workbook",
    "normalize_key",
    "parse_coordinate",
    "resolve_input_path",
    "suggest_join",
]
