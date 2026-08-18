"""
Builds the "Allocation" sheet's data: Confirmed Quantity (CS) from the open
sales-order export, aggregated by Material (rows) and Plant (columns), with
row/column Grand Totals -- reproducing the reference workbook's PivotTable
output as a plain (static) grid.

A native, refreshable Excel PivotTable is not reproduced here: openpyxl
cannot reliably author the pivot-cache XML a live PivotTable needs, and
attempting it risks the workbook failing to open cleanly in Excel. The
*data* the reference pivot displays (same row/column labels, same layout,
same aggregation, same Grand Totals) is reproduced exactly; only the
"live, click-to-refresh" pivot mechanism is not.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from openpyxl.utils import get_column_letter

from src import config

BLANK_LABEL = "(blank)"


def _blankify(series: pd.Series) -> pd.Series:
    s = series.astype("object")
    s = s.where(s.notna(), BLANK_LABEL)
    s = s.map(lambda v: BLANK_LABEL if isinstance(v, str) and v.strip() == "" else v)
    return s


def _sort_key(value: str):
    # Real values sort ascending first; "(blank)" always sorts last.
    return (1, "") if value == BLANK_LABEL else (0, value)


@dataclass
class AllocationResult:
    materials: list[str]
    plants: list[str]
    grid: dict[str, dict[str, float | None]]
    row_totals: dict[str, float]
    col_totals: dict[str, float]
    grand_total: float
    plant_column_letter: dict[str, str] = field(default_factory=dict)


def build_allocation(sales_order_df: pd.DataFrame) -> AllocationResult:
    work = sales_order_df[list(config.ALLOCATION_SOURCE_COLUMNS)].copy()
    work["Material"] = _blankify(work["Material"])
    work["Plant"] = _blankify(work["Plant"])
    work["Confirmed Quantity (CS)"] = pd.to_numeric(
        work["Confirmed Quantity (CS)"], errors="coerce"
    )

    materials = sorted(work["Material"].unique(), key=_sort_key)
    plants = sorted(work["Plant"].unique(), key=_sort_key)

    pivot = pd.pivot_table(
        work,
        index="Material",
        columns="Plant",
        values="Confirmed Quantity (CS)",
        aggfunc="sum",
    ).reindex(index=materials, columns=plants)

    grid: dict[str, dict[str, float | None]] = {}
    row_totals: dict[str, float] = {}
    for material in materials:
        row = {}
        total = 0.0
        for plant in plants:
            val = pivot.loc[material, plant]
            if pd.isna(val):
                row[plant] = None
            else:
                row[plant] = float(val)
                total += float(val)
        grid[material] = row
        row_totals[material] = total

    col_totals: dict[str, float] = {}
    for plant in plants:
        col_totals[plant] = float(pivot[plant].fillna(0).sum())

    grand_total = float(sum(col_totals.values()))

    # Column A is the Material row-label column, so plant columns start at B.
    plant_column_letter = {
        plant: get_column_letter(2 + idx) for idx, plant in enumerate(plants)
    }

    return AllocationResult(
        materials=materials,
        plants=plants,
        grid=grid,
        row_totals=row_totals,
        col_totals=col_totals,
        grand_total=grand_total,
        plant_column_letter=plant_column_letter,
    )


def selling_plant_column_letter(result: AllocationResult, plant: str | None) -> str | None:
    """Resolves a Materials-export Plant value to the Allocation column
    letter it should be looked up against, applying the city-pairing rule
    (Lineage 3PL plants share their TOL selling plant's column). Returns
    None if that selling plant has no confirmed-order data at all (i.e. its
    Allocation figure is definitionally 0)."""
    if plant is None:
        return None
    selling_plant = config.PLANT_TO_SELLING_PLANT.get(plant, plant)
    return result.plant_column_letter.get(selling_plant)
