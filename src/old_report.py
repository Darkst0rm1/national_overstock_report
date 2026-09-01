"""Builds the row data for the "Old report" sheet from the Materials export.

This module owns the inventory side of the report: filtering the Materials
export down to the eligible population (src/eligibility.py) and FIFO-
consuming each material/plant's allocated quantity across that population's
batches (src/fifo_allocation.py). It produces every *value* column
(Material, ..., Unrestricted Stock, Unique Key, Allocated, Available).

BDM / Sold by KG / Pack / Size / DSD Price / Deal Price are formula-driven
in the reference workbook and stay that way here -- written as live Excel
formulas by excel_writer.py, not computed in this module -- so they remain
recalculable exactly like the reference (e.g. if the Price List sheet is
edited later, they update automatically). Allocated/Available are the
documented exception: they are plain computed values because the Allocation
sheet is a native PivotTable, and any formula referencing a pivot's own
output cells is unreliable until the pivot's refreshOnLoad completes (see
excel_writer.py's _write_allocation_sheet for the full rationale).
"""
from __future__ import annotations

from datetime import date

import pandas as pd

from src.eligibility import (
    eligible_material_universe,
    filter_eligible_inventory,
    filter_food_service_inventory,
)
from src.fifo_allocation import allocate_fifo_by_batch

DIRECT_COLUMNS = [
    "Material",
    "Material Description",
    "Plant",
    "Plant Name",
    "Batch",
    "Storage Location",
    "Shelf Life Expiration Date",
    "Unrestricted Stock",
]


def build_old_report(
    materials_df: pd.DataFrame,
    pricing_df: pd.DataFrame,
    allocation_totals: dict[tuple[str, str], float],
    report_date: date,
) -> pd.DataFrame:
    eligible_materials = eligible_material_universe(materials_df, pricing_df)
    inventory = filter_eligible_inventory(materials_df, eligible_materials, report_date)
    inventory = inventory[DIRECT_COLUMNS].copy()
    return allocate_fifo_by_batch(inventory, allocation_totals)


def build_food_service_report(
    materials_df: pd.DataFrame,
    pricing_df: pd.DataFrame,
    allocation_totals: dict[tuple[str, str], float],
) -> pd.DataFrame:
    """Row data for the "Foodservice" sheet -- kept for future use, per the
    rule 3 (excluded-BDM) population of build_old_report's inventory filter,
    with no date window (see filter_food_service_inventory). These materials
    were never in scope for the Allocation aggregation, so Allocated will
    naturally be 0 for every row unless allocation_totals happens to carry a
    stray key for one of them."""
    inventory = filter_food_service_inventory(materials_df, pricing_df)
    inventory = inventory[DIRECT_COLUMNS].copy()
    return allocate_fifo_by_batch(inventory, allocation_totals)
