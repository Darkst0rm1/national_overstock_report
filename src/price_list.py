"""
Builds the "Price List" sheet from the material/pricing export.

Every target column is copied through under its exact source name when
present; columns the source export doesn't contain are left blank (NaN) for
every row -- never filled from the reference workbook or invented.
"""
from __future__ import annotations

import pandas as pd

from src import config


def build_price_list(pricing_df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=pricing_df.index)
    for col in config.PRICE_LIST_COLUMNS:
        if col in pricing_df.columns:
            out[col] = pricing_df[col]
        else:
            out[col] = pd.NA
    return out
