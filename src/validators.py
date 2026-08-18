"""
Column-presence validation for the three uploaded exports.

Each validator compares the file's actual header row (after trimming
whitespace) against the required column list and returns a list of
human-readable, specific error/warning messages -- never raises, so the
caller can display everything that's wrong at once instead of failing on
the first problem.
"""
from __future__ import annotations

import pandas as pd

from src import config


def _clean_headers(df: pd.DataFrame) -> list[str]:
    return [str(c).strip() for c in df.columns]


def validate_materials(df: pd.DataFrame) -> list[str]:
    headers = set(_clean_headers(df))
    missing = [c for c in config.MATERIALS_REQUIRED_COLUMNS if c not in headers]
    if missing:
        return [
            "Materials export is missing required column(s): " + ", ".join(missing)
        ]
    return []


def validate_sales_order(df: pd.DataFrame) -> list[str]:
    headers = set(_clean_headers(df))
    missing = [c for c in config.OVERSTOCK_COLUMNS if c not in headers]
    if missing:
        return [
            "Open sales-order export is missing required column(s): "
            + ", ".join(missing)
        ]
    return []


def validate_pricing(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    """Returns (blocking_errors, soft_warnings).

    Only "Material" is a hard requirement (it is the join key for every
    lookup into the Price List sheet). The remaining Price List columns are
    optional in the source per spec -- missing ones are reported as a soft
    warning and left blank in the output, never invented.
    """
    headers = set(_clean_headers(df))
    errors = [
        f"Material/pricing export is missing required column: {c}"
        for c in config.PRICE_LIST_HARD_REQUIRED_COLUMNS
        if c not in headers
    ]
    missing_optional = [
        c
        for c in config.PRICE_LIST_COLUMNS
        if c not in headers and c not in config.PRICE_LIST_HARD_REQUIRED_COLUMNS
    ]
    warnings = []
    if missing_optional:
        warnings.append(
            "Material/pricing export does not contain: "
            + ", ".join(missing_optional)
            + ". These Price List column(s) will be left blank."
        )
    return errors, warnings
