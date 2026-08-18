"""
Builds the "Overstock" sheet: the complete open sales-order export,
reordered into the exact column order required by the report spec (source
columns already carry these exact names; this only selects/reorders them,
it never renames or drops data).
"""
from __future__ import annotations

import pandas as pd

from src import config


def build_overstock(sales_order_df: pd.DataFrame) -> pd.DataFrame:
    return sales_order_df[list(config.OVERSTOCK_COLUMNS)].copy()
