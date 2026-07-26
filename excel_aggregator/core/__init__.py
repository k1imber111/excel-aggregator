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
from .matcher import Candidate, find_key_candidates
from .models import (
    Orientation,
    SheetGrid,
    SourceConfig,
    Workbook,
    col_letter,
    parse_coordinate,
)
from .normalize import format_cell, keys_equal, normalize_key
from .paths import resolve_input_path
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
    "Workbook",
    "col_letter",
    "export",
    "extract_keys",
    "find_key_candidates",
    "format_cell",
    "keys_equal",
    "load_workbook",
    "normalize_key",
    "parse_coordinate",
    "resolve_input_path",
]
