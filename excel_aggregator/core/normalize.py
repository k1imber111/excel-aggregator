"""Нормализация ключей и значений ячеек."""

from __future__ import annotations

import re

_INT_RE = re.compile(r"-?\d+")
_FLOAT_RE = re.compile(r"-?\d+[.,]\d+")


def _has_leading_zero(digits: str) -> bool:
    """Есть ли ведущий нуль у целой части числа ("040", "04.61")."""
    int_part = digits.lstrip("-").split(".")[0].split(",")[0]
    return len(int_part) > 1 and int_part.startswith("0")


def normalize_key(v) -> str | None:
    """Каноническая строковая форма ключа или None для пустых значений.

    Числа 2, 2.0 и строки "2", " 22 ", "2.0" -> "2" / "22".
    Неразрывные пробелы -> обычные, края обрезаются. Регистр сохраняется.
    Строки с ведущими нулями ("040") не числятся числами и сохраняются как есть.
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer():
            return str(int(f))
        return str(v)

    s = str(v).replace("\xa0", " ").strip()
    if not s:
        return None

    if _INT_RE.fullmatch(s) and not _has_leading_zero(s):
        return str(int(s))
    if _FLOAT_RE.fullmatch(s) and not _has_leading_zero(s):
        f = float(s.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    return s


def keys_equal(a, b) -> bool:
    """Сравнение ключей через normalize_key; для текста — без учёта регистра."""
    ka, kb = normalize_key(a), normalize_key(b)
    if ka is None or kb is None:
        return False
    return ka.casefold() == kb.casefold()


def format_cell(v):
    """Значение для показа в консоли/превью (исходное значение не меняется).

    float с целым значением -> int, None -> пустая строка.
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v
