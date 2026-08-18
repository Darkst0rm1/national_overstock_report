"""
Writes the final workbook by editing the reference workbook IN PLACE as a
live template, rather than rebuilding one from scratch with openpyxl.

Why: the reference file's "Allocation" sheet is a real, native Excel
PivotTable, and its source range is 'Overstock'!A1:AE1048576 -- i.e. the
pivot reads directly off the Overstock sheet in the same workbook. openpyxl
round-trips that PivotTable's XML intact (verified: load + save reproduces
the same cacheSource, cacheFields, and pivotFields). So instead of trying
to fabricate pivot-cache XML ourselves (unsupported, corruption-prone), we:

  1. Load the reference workbook as a template.
  2. Overwrite the data rows of Overstock / Price List / Old report with
     this run's data, resizing each sheet's row count and cloning the
     template's own cell styles onto any newly-inserted rows -- so every
     font, fill, border, number format, column width, freeze pane, and
     autofilter is inherited automatically instead of hand re-implemented.
  3. Set the PivotTable's cache to refreshOnLoad=1 and force ascending
     sort on the Material/Plant fields, so Excel recomputes a fresh,
     correctly laid-out pivot the moment the file is opened.
  4. Also pre-populate the pivot's own A:G cells with the exact values our
     own aggregation (src/allocation.py) computes -- identical to what
     Excel's live refresh will produce -- so the sheet looks correct even
     before any refresh happens, and rebuild the manually-added "Unique
     Key" helper column (H) to match that same row range.

This keeps the PivotTable genuinely native (field list, drill-down,
right-click Refresh all work), while guaranteeing correct-looking data on
first open regardless of refresh timing.
"""
from __future__ import annotations

from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles.colors import Color
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from src import config
from src.allocation import AllocationResult
from src.old_report import ALLOCATION_COLUMN_LETTER_FIELD

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "assets" / "report_template.xlsx"

# Fixed cacheField positions in the reference workbook's Overstock column
# order (0-indexed) -- these only stay valid because OVERSTOCK_COLUMNS is
# never reordered relative to the reference file.
_MATERIAL_PIVOT_FIELD_INDEX = config.OVERSTOCK_COLUMNS.index("Material")
_PLANT_PIVOT_FIELD_INDEX = config.OVERSTOCK_COLUMNS.index("Plant")


def _clean(value):
    if value is None:
        return None
    if pd.api.types.is_scalar(value) and pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _sync_row_count(ws, old_last_row: int, new_last_row: int, style_template_row: int, n_cols: int):
    """Grows or shrinks a sheet's data rows to exactly new_last_row,
    cloning cell styles from style_template_row onto any newly-inserted
    rows so they stay formatted like the rest of the column."""
    if new_last_row > old_last_row:
        n_new = new_last_row - old_last_row
        ws.insert_rows(old_last_row + 1, amount=n_new)
        for r in range(old_last_row + 1, new_last_row + 1):
            for c in range(1, n_cols + 1):
                src = ws.cell(style_template_row, c)
                dst = ws.cell(r, c)
                dst.font = copy(src.font)
                dst.fill = copy(src.fill)
                dst.border = copy(src.border)
                dst.alignment = copy(src.alignment)
                dst.number_format = src.number_format
    elif new_last_row < old_last_row:
        ws.delete_rows(new_last_row + 1, old_last_row - new_last_row)


# ---------------------------------------------------------------------------
# Overstock
# ---------------------------------------------------------------------------
def _write_overstock_sheet(wb: Workbook, overstock_df: pd.DataFrame):
    ws = wb["Overstock"]
    columns = config.OVERSTOCK_COLUMNS
    old_last_row = ws.max_row
    new_last_row = overstock_df.shape[0] + 1

    _sync_row_count(ws, old_last_row, new_last_row, style_template_row=2, n_cols=len(columns))

    for r_idx, (_, row) in enumerate(overstock_df.iterrows(), start=2):
        for c_idx, col_name in enumerate(columns, start=1):
            ws.cell(r_idx, c_idx, _clean(row[col_name]))

    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{new_last_row}"


# ---------------------------------------------------------------------------
# Price List
# ---------------------------------------------------------------------------
def _write_price_list_sheet(wb: Workbook, price_list_df: pd.DataFrame):
    ws = wb["Price List"]
    columns = config.PRICE_LIST_COLUMNS
    old_last_row = ws.max_row
    new_last_row = price_list_df.shape[0] + 1

    _sync_row_count(ws, old_last_row, new_last_row, style_template_row=2, n_cols=len(columns))

    for r_idx, (_, row) in enumerate(price_list_df.iterrows(), start=2):
        for c_idx, col_name in enumerate(columns, start=1):
            ws.cell(r_idx, c_idx, _clean(row[col_name]))

    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{new_last_row}"


# ---------------------------------------------------------------------------
# Old report
# ---------------------------------------------------------------------------
def _write_old_report_sheet(
    wb: Workbook,
    old_report_df: pd.DataFrame,
    allocation_unique_key_letter: str,
):
    ws = wb["Old report"]
    columns = config.OLD_REPORT_COLUMNS
    old_last_row = ws.max_row
    new_last_row = old_report_df.shape[0] + 1

    _sync_row_count(ws, old_last_row, new_last_row, style_template_row=2, n_cols=len(columns))

    price_list_letter = {
        name: get_column_letter(idx + 1) for idx, name in enumerate(config.PRICE_LIST_COLUMNS)
    }

    for r_offset, (_, src_row) in enumerate(old_report_df.iterrows()):
        r = r_offset + 2

        ws.cell(r, 1, f"=_xlfn.XLOOKUP(C:C,'Price List'!C:C,'Price List'!{price_list_letter['BDM']}:{price_list_letter['BDM']})")
        ws.cell(r, 2, _clean(src_row["Material Group Name"]))
        ws.cell(r, 3, _clean(src_row["Material"]))
        ws.cell(r, 4, _clean(src_row["Material Description"]))
        ws.cell(r, 5, f"=_xlfn.XLOOKUP(C:C,'Price List'!C:C,'Price List'!{price_list_letter['Sold by Wt']}:{price_list_letter['Sold by Wt']})")
        ws.cell(r, 6, f"=_xlfn.XLOOKUP(C:C,'Price List'!C:C,'Price List'!{price_list_letter['Pack']}:{price_list_letter['Pack']})")
        ws.cell(r, 7, f"=_xlfn.XLOOKUP(C:C,'Price List'!C:C,'Price List'!{price_list_letter['Size']}:{price_list_letter['Size']})")
        ws.cell(r, 8, _clean(src_row["Plant"]))
        ws.cell(r, 9, _clean(src_row["Plant Name"]))
        ws.cell(r, 10, _clean(src_row["Batch"]))
        ws.cell(r, 11, _clean(src_row["Storage Location"]))
        ws.cell(r, 12, _clean(src_row["Shelf Life Expiration Date"]))
        ws.cell(r, 13, _clean(src_row["Unrestricted Stock"]))
        ws.cell(r, 14, f'=C{r} & "_" & {config.UNIQUE_KEY_SUFFIX}')

        alloc_col_letter = src_row[ALLOCATION_COLUMN_LETTER_FIELD]
        if alloc_col_letter:
            ws.cell(
                r, 15,
                f"=_xlfn.XLOOKUP(N{r},Allocation!{allocation_unique_key_letter}:{allocation_unique_key_letter},"
                f"Allocation!{alloc_col_letter}:{alloc_col_letter},0)",
            )
        else:
            ws.cell(r, 15, 0)

        ws.cell(r, 16, f"=M{r}-O{r}")

    ws.auto_filter.ref = f"A1:{config.OLD_REPORT_AUTOFILTER_LAST_COLUMN}{new_last_row}"


# ---------------------------------------------------------------------------
# Allocation (real PivotTable, pre-populated + set to refresh live)
# ---------------------------------------------------------------------------
def _write_allocation_sheet(wb: Workbook, allocation: AllocationResult, sales_order_row_count: int) -> str:
    ws = wb["Allocation"]
    pt = ws._pivots[0]

    n_plants = len(allocation.plants)
    grand_total_col_idx = 2 + n_plants
    unique_key_col_idx = grand_total_col_idx + 1
    grand_total_letter = get_column_letter(grand_total_col_idx)
    unique_key_letter = get_column_letter(unique_key_col_idx)

    old_last_row = ws.max_row
    old_last_col = ws.max_column
    new_last_row = 5 + len(allocation.materials)  # header rows 3-4, data starts row 5, +1 grand-total row

    _sync_row_count(ws, old_last_row, new_last_row, style_template_row=5, n_cols=unique_key_col_idx)

    # The template's plant-column count (and therefore where Grand Total /
    # Unique Key land) depends on this run's data and can be narrower than
    # the template's own. Clear any stale trailing columns left over from
    # the template's wider layout so old labels/values don't linger next to
    # the new ones.
    if old_last_col > unique_key_col_idx:
        for r in range(3, max(old_last_row, new_last_row) + 1):
            for c in range(unique_key_col_idx + 1, old_last_col + 1):
                ws.cell(r, c).value = None
    elif grand_total_col_idx > old_last_col:
        # This run has more distinct plants than the template did -- clone
        # an existing pivot-body column's style onto the new columns so
        # they stay visually consistent (thin borders, matching font).
        for r in range(4, new_last_row + 1):
            style_src = ws.cell(r, 2)
            for c in range(old_last_col + 1, grand_total_col_idx + 1):
                dst = ws.cell(r, c)
                dst.font = copy(style_src.font)
                dst.border = copy(style_src.border)
                dst.alignment = copy(style_src.alignment)

    # Header row (row 4): plant labels, Grand Total, Unique Key.
    for idx, plant in enumerate(allocation.plants):
        ws.cell(4, 2 + idx, plant)
    ws.cell(4, grand_total_col_idx, "Grand Total")
    key_header = ws.cell(4, unique_key_col_idx, "Unique Key")
    key_header.font = Font(name="Arial", size=11, bold=True)
    key_header.fill = PatternFill(patternType="solid", fgColor=Color(theme=4, tint=0.7999816888943144))

    row = 5
    for material in allocation.materials:
        ws.cell(row, 1, material)
        for idx, plant in enumerate(allocation.plants):
            ws.cell(row, 2 + idx, allocation.grid[material][plant])
        ws.cell(row, grand_total_col_idx, allocation.row_totals[material])
        ws.cell(row, unique_key_col_idx, f'=A{row} & "_" & {config.UNIQUE_KEY_SUFFIX}')
        row += 1

    ws.cell(row, 1, "Grand Total")
    for idx, plant in enumerate(allocation.plants):
        ws.cell(row, 2 + idx, allocation.col_totals[plant])
    ws.cell(row, grand_total_col_idx, allocation.grand_total)
    ws.cell(row, unique_key_col_idx, f'=A{row} & "_" & {config.UNIQUE_KEY_SUFFIX}')

    # Make this a genuinely live, refreshable PivotTable: force a fresh
    # recompute from the (new) Overstock data on open, sorted ascending so
    # the refreshed layout matches what we just pre-populated above.
    pt.cache.refreshOnLoad = True
    pt.cache.recordCount = sales_order_row_count
    pt.pivotFields[_MATERIAL_PIVOT_FIELD_INDEX].sortType = "ascending"
    pt.pivotFields[_PLANT_PIVOT_FIELD_INDEX].sortType = "ascending"
    pt.location.ref = f"A3:{grand_total_letter}{row}"

    return unique_key_letter


def build_workbook(
    overstock_df: pd.DataFrame,
    price_list_df: pd.DataFrame,
    old_report_df: pd.DataFrame,
    allocation: AllocationResult,
) -> Workbook:
    wb = load_workbook(TEMPLATE_PATH)

    # Allocation must be written first: its column layout (which depends on
    # how many distinct plants are in this run's data) determines the exact
    # cell references Old report's formulas need.
    unique_key_letter = _write_allocation_sheet(wb, allocation, overstock_df.shape[0])
    _write_overstock_sheet(wb, overstock_df)
    _write_price_list_sheet(wb, price_list_df)
    _write_old_report_sheet(wb, old_report_df, unique_key_letter)

    return wb
