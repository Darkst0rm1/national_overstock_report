"""
Builds the row data for the "Old report" sheet from the Materials export.

This module only produces the columns that come directly from the Materials
export (Material Group Name, Material, Material Description, Plant, Plant
Name, Batch, Storage Location, Shelf Life Expiration Date, Unrestricted
Stock) plus one internal helper column recording, per row, which Allocation
sheet column letter that row's plant resolves to.

BDM / Sold by KG / Pack / Size / Unique Key / Allocation / Available are all
formula-driven in the reference workbook and are written as live Excel
formulas by excel_writer.py -- not computed here -- so they stay
recalculable exactly like the reference (e.g. if the Price List sheet is
edited later, BDM/Pack/Size update automatically).
"""
from __future__ import annotations

import pandas as pd

from src import config
from src.allocation import AllocationResult, selling_plant_column_letter

DIRECT_COLUMNS = [
    "Material Group Name",
    "Material",
    "Material Description",
    "Plant",
    "Plant Name",
    "Batch",
    "Storage Location",
    "Shelf Life Expiration Date",
    "Unrestricted Stock",
]

ALLOCATION_COLUMN_LETTER_FIELD = "_allocation_column_letter"


def build_old_report(materials_df: pd.DataFrame, allocation: AllocationResult) -> pd.DataFrame:
    out = materials_df[DIRECT_COLUMNS].copy()
    out[ALLOCATION_COLUMN_LETTER_FIELD] = out["Plant"].map(
        lambda p: selling_plant_column_letter(allocation, p)
    )
    return out
