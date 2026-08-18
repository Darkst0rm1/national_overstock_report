"""
Reads the three uploaded exports into pandas DataFrames.

Identifier columns (Material, Plant, Batch, order numbers, etc.) are read
with dtype=str directly at parse time -- fixing dtype *after* pandas has
already inferred int64 would be too late, leading zeros would already be
gone. Whitespace (including non-breaking space, which the ASCII-only \\s
regex class misses) is trimmed with Python's own str.strip(), which is
unicode-aware.
"""
from __future__ import annotations

import io

import pandas as pd

from src import config


def _read_excel(file, text_columns: list[str]) -> pd.DataFrame:
    dtype_map = {c: str for c in text_columns}
    df = pd.read_excel(file, dtype=dtype_map, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
    return df


def load_materials(file) -> pd.DataFrame:
    return _read_excel(file, config.MATERIALS_TEXT_COLUMNS)


def load_sales_order(file) -> pd.DataFrame:
    return _read_excel(file, config.SALES_ORDER_TEXT_COLUMNS)


def load_pricing(file) -> pd.DataFrame:
    return _read_excel(file, config.PRICING_TEXT_COLUMNS)


def to_buffer(uploaded_file) -> io.BytesIO:
    """Streamlit's UploadedFile is already file-like, but reading it twice
    (e.g. once for a preview, once for processing) requires rewinding or
    working off an independent buffer. Centralizing this avoids each call
    site needing to remember to seek(0)."""
    uploaded_file.seek(0)
    buf = io.BytesIO(uploaded_file.read())
    uploaded_file.seek(0)
    return buf
