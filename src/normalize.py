"""Shared identifier normalization for joining Material/Plant/Batch values
across the three uploaded exports.

pandas reads these columns as text (see data_loader.py), but a value can
still arrive as an Excel-numeric-looking string (e.g. "10017617.0") if a
caller builds a DataFrame directly (tests, or a future code path) without
going through data_loader. Centralizing normalization here means every join
key -- inventory filtering, sales-order filtering, allocation aggregation,
FIFO consumption -- treats "10017617" and "10017617.0" as the same material.
"""
from __future__ import annotations


def normalize_identifier(value) -> str | None:
    """Strips surrounding whitespace and a trailing Excel ".0", but never
    touches leading zeros (batch numbers rely on them). Returns None for
    missing values."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith(".0") and text[:-2].lstrip("-").isdigit():
        text = text[:-2]
    return text
