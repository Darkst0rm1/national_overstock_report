"""
Writes the final 4-sheet workbook with openpyxl, reproducing the reference
workbook's sheet order, column layout, number formats, fonts/fills/borders,
freeze panes, autofilters, column widths/row heights, and live formulas.

Sheet write order matters: Allocation must be written first because its
column layout (which depends on how many distinct plants are in this run's
sales-order data) determines the exact cell references used by Old report's
Allocation-lookup formulas.
"""
from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

from src import config
from src.allocation import AllocationResult
from src.old_report import ALLOCATION_COLUMN_LETTER_FIELD

_REGULAR_FONT = Font(name=config.FONT_NAME, size=config.FONT_SIZE)
_BOLD_FONT = Font(name=config.FONT_NAME, size=config.FONT_SIZE, bold=True)
_HEADER_FILL = PatternFill(patternType="solid", fgColor=config.HEADER_FILL_RGB)
_THIN = Side(style="thin")
_THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _clean(value):
    if value is None:
        return None
    if pd.api.types.is_scalar(value) and pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _set_widths(ws, widths: dict[str, float]):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


# ---------------------------------------------------------------------------
# Allocation
# ---------------------------------------------------------------------------
def _write_allocation_sheet(wb: Workbook, allocation: AllocationResult):
    ws = wb.create_sheet("Allocation")

    n_plants = len(allocation.plants)
    grand_total_col_idx = 2 + n_plants
    unique_key_col_idx = grand_total_col_idx + 1
    grand_total_letter = get_column_letter(grand_total_col_idx)
    unique_key_letter = get_column_letter(unique_key_col_idx)

    ws["A3"] = "Sum of Confirmed Quantity (CS)"
    ws["A3"].font = _REGULAR_FONT
    ws["B3"] = "Column Labels"
    ws["B3"].font = _REGULAR_FONT

    center = Alignment(horizontal="center", vertical="center")
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.cell(4, 1, "Row Labels").font = _REGULAR_FONT
    ws.cell(4, 1).border = _THIN_BORDER
    for idx, plant in enumerate(allocation.plants):
        cell = ws.cell(4, 2 + idx, plant)
        cell.font = _REGULAR_FONT
        cell.alignment = center
        cell.border = _THIN_BORDER
    gt_header = ws.cell(4, grand_total_col_idx, "Grand Total")
    gt_header.font = _REGULAR_FONT
    gt_header.alignment = center
    gt_header.border = _THIN_BORDER
    key_header = ws.cell(4, unique_key_col_idx, "Unique Key")
    key_header.font = _BOLD_FONT
    key_header.alignment = center_wrap
    key_header.fill = PatternFill(patternType="solid", fgColor=Color(theme=4, tint=0.7999816888943144))

    row = 5
    for material in allocation.materials:
        ws.cell(row, 1, material).font = _REGULAR_FONT
        ws.cell(row, 1).border = _THIN_BORDER
        for idx, plant in enumerate(allocation.plants):
            val = allocation.grid[material][plant]
            cell = ws.cell(row, 2 + idx, val)
            cell.font = _REGULAR_FONT
            cell.border = _THIN_BORDER
        gt_cell = ws.cell(row, grand_total_col_idx, allocation.row_totals[material])
        gt_cell.font = _REGULAR_FONT
        gt_cell.border = _THIN_BORDER
        key_cell = ws.cell(row, unique_key_col_idx, f'=A{row} & "_" & {config.UNIQUE_KEY_SUFFIX}')
        key_cell.font = _REGULAR_FONT
        row += 1

    # Grand Total row
    ws.cell(row, 1, "Grand Total").font = _REGULAR_FONT
    ws.cell(row, 1).border = _THIN_BORDER
    for idx, plant in enumerate(allocation.plants):
        cell = ws.cell(row, 2 + idx, allocation.col_totals[plant])
        cell.font = _REGULAR_FONT
        cell.border = _THIN_BORDER
    gt_cell = ws.cell(row, grand_total_col_idx, allocation.grand_total)
    gt_cell.font = _REGULAR_FONT
    gt_cell.border = _THIN_BORDER
    ws.cell(row, unique_key_col_idx, f'=A{row} & "_" & {config.UNIQUE_KEY_SUFFIX}').font = _REGULAR_FONT

    _set_widths(ws, config.COLUMN_WIDTHS["Allocation"])
    return unique_key_letter


# ---------------------------------------------------------------------------
# Overstock
# ---------------------------------------------------------------------------
def _write_overstock_sheet(wb: Workbook, overstock_df: pd.DataFrame):
    ws = wb.create_sheet("Overstock")
    columns = config.OVERSTOCK_COLUMNS

    for c_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(1, c_idx, col_name)
        cell.font = _BOLD_FONT
        cell.fill = _HEADER_FILL
        align = config.OVERSTOCK_HEADER_ALIGN.get(col_name)
        vertical = "center" if align else None
        cell.alignment = Alignment(horizontal=align, vertical=vertical, wrap_text=True)

    for r_idx, (_, row) in enumerate(overstock_df.iterrows(), start=2):
        for c_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(r_idx, c_idx, _clean(row[col_name]))
            cell.font = _REGULAR_FONT
            fmt = config.OVERSTOCK_NUMBER_FORMATS.get(col_name)
            if fmt:
                cell.number_format = fmt

    ws.row_dimensions[1].height = config.HEADER_ROW_HEIGHT["Overstock"]
    ws.freeze_panes = config.FREEZE_PANES["Overstock"]
    last_row = max(overstock_df.shape[0] + 1, 1)
    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    _set_widths(ws, config.COLUMN_WIDTHS["Overstock"])


# ---------------------------------------------------------------------------
# Price List
# ---------------------------------------------------------------------------
def _write_price_list_sheet(wb: Workbook, price_list_df: pd.DataFrame):
    ws = wb.create_sheet("Price List")
    columns = config.PRICE_LIST_COLUMNS

    for c_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(1, c_idx, col_name)
        cell.font = _BOLD_FONT
        cell.fill = _HEADER_FILL
        align = config.PRICE_LIST_HEADER_ALIGN.get(col_name)
        vertical = "center" if align else None
        cell.alignment = Alignment(horizontal=align, vertical=vertical)

    for r_idx, (_, row) in enumerate(price_list_df.iterrows(), start=2):
        for c_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(r_idx, c_idx, _clean(row[col_name]))
            cell.font = _REGULAR_FONT
            fmt = config.PRICE_LIST_NUMBER_FORMATS.get(col_name)
            if fmt:
                cell.number_format = fmt

    ws.freeze_panes = config.FREEZE_PANES["Price List"]
    last_row = max(price_list_df.shape[0] + 1, 1)
    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    _set_widths(ws, config.COLUMN_WIDTHS["Price List"])


# ---------------------------------------------------------------------------
# Old report
# ---------------------------------------------------------------------------
def _write_old_report_sheet(
    wb: Workbook,
    old_report_df: pd.DataFrame,
    allocation: AllocationResult,
    allocation_unique_key_letter: str,
):
    ws = wb.create_sheet("Old report")
    columns = config.OLD_REPORT_COLUMNS

    price_list_letter = {
        name: get_column_letter(idx + 1) for idx, name in enumerate(config.PRICE_LIST_COLUMNS)
    }

    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(1, c_idx, col_name)
        cell.font = _BOLD_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = header_align

    for r_offset, (_, src_row) in enumerate(old_report_df.iterrows()):
        r = r_offset + 2

        # Formula-driven columns, matching the reference workbook's exact
        # formula structure (whole-column XLOOKUP references).
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
            # Selling plant has no confirmed-order data in this run at all,
            # so the Allocation figure is definitionally 0 (same fallback
            # XLOOKUP itself would use if the lookup failed).
            ws.cell(r, 15, 0)

        ws.cell(r, 16, f"=M{r}-O{r}")

        for c_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(r, c_idx)
            cell.font = _REGULAR_FONT
            fmt = config.OLD_REPORT_NUMBER_FORMATS.get(col_name)
            if fmt:
                cell.number_format = fmt

    ws.row_dimensions[1].height = config.HEADER_ROW_HEIGHT["Old report"]
    ws.freeze_panes = config.FREEZE_PANES["Old report"]
    last_row = max(old_report_df.shape[0] + 1, 1)
    ws.auto_filter.ref = f"A1:{config.OLD_REPORT_AUTOFILTER_LAST_COLUMN}{last_row}"
    _set_widths(ws, config.COLUMN_WIDTHS["Old report"])
    return ws


def build_workbook(
    overstock_df: pd.DataFrame,
    price_list_df: pd.DataFrame,
    old_report_df: pd.DataFrame,
    allocation: AllocationResult,
) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    unique_key_letter = _write_allocation_sheet(wb, allocation)
    _write_overstock_sheet(wb, overstock_df)
    _write_price_list_sheet(wb, price_list_df)
    old_report_ws = _write_old_report_sheet(wb, old_report_df, allocation, unique_key_letter)

    old_report_ws.sheet_view.tabSelected = True
    wb.active = wb.sheetnames.index("Old report")
    return wb
