import pandas as pd

from src import config
from src.allocation import build_allocation
from src.excel_writer import build_workbook
from src.old_report import build_old_report
from src.overstock import build_overstock
from src.price_list import build_price_list


def _tiny_sales_order_df():
    row = {c: None for c in config.OVERSTOCK_COLUMNS}
    row.update(
        {
            "Sales Order": "3000001",
            "Sales Order Item": "10",
            "Material": "10000001",
            "Material Description": "TEST ITEM",
            "Confirmed Quantity (CS)": 5,
            "Plant": "2910",
        }
    )
    return pd.DataFrame([row])


def _tiny_pricing_df():
    return pd.DataFrame([{"Material": "10000001", "BDM": "TEST BDM", "Pack": 1, "Size": "1 EA", "Sold by Wt": 0}])


def _tiny_materials_df():
    return pd.DataFrame(
        [
            {
                "Material Group Name": "TEST GROUP",
                "Material": "10000001",
                "Material Description": "TEST ITEM",
                "Plant": "2910",
                "Plant Name": "TOL Mississauga",
                "Batch": "0001",
                "Storage Location": "1100",
                "Shelf Life Expiration Date": pd.Timestamp("2027-01-01"),
                "Unrestricted Stock": 10,
            }
        ]
    )


def _build():
    sales_order_df = _tiny_sales_order_df()
    allocation = build_allocation(sales_order_df)
    overstock_df = build_overstock(sales_order_df)
    price_list_df = build_price_list(_tiny_pricing_df())
    old_report_df = build_old_report(_tiny_materials_df(), allocation)
    return build_workbook(overstock_df, price_list_df, old_report_df, allocation)


def test_sheet_order_and_names():
    wb = _build()
    assert wb.sheetnames == ["Allocation", "Overstock", "Price List", "Old report"]


def test_pivot_set_to_refresh_live_and_sorted_ascending():
    wb = _build()
    pt = wb["Allocation"]._pivots[0]
    assert pt.cache.refreshOnLoad is True
    assert pt.pivotFields[config.OVERSTOCK_COLUMNS.index("Material")].sortType == "ascending"
    assert pt.pivotFields[config.OVERSTOCK_COLUMNS.index("Plant")].sortType == "ascending"


def test_allocation_sheet_pivot_location_and_cells_are_left_for_excel_to_refresh():
    # Verified against real Excel (COM): pre-editing the pivot's own
    # location.ref/header/data cells before its refreshOnLoad-triggered
    # refresh runs makes that refresh collapse to a single #SPILL! cell
    # instead of resizing cleanly. Excel's native refresh already
    # resizes/repopulates the pivot correctly on its own, so
    # _write_allocation_sheet must leave the pivot's own output area alone
    # and only configure the cache (record count, sort order, refresh flag).
    wb = _build()
    ws = wb["Allocation"]
    pt = ws._pivots[0]
    assert pt.location.ref == "A3:G117"  # untouched template value, not shrunk to this run's size


def test_old_report_allocation_is_a_plain_value_not_a_pivot_formula():
    # Verified against real Excel (COM), including a genuine interactive
    # Ctrl+Alt+F9 full recalculation: ANY formula referencing the Allocation
    # pivot's output cells -- plain cell reference, INDEX/MATCH, or even
    # GETPIVOTDATA -- gets stuck at a stale/error-caught value from before
    # the pivot's refreshOnLoad refresh completes, and nothing short of
    # literally re-entering the formula fixes it. So Allocation is written
    # as a plain number, computed the same way the pivot itself aggregates,
    # which has no such failure mode and can't be misaligned by a later row
    # deletion in "Old report" either -- it just moves/deletes with its row.
    wb = _build()
    ws = wb["Old report"]
    cell = ws.cell(2, 15)
    assert cell.data_type == "n"  # numeric constant, not a formula ("f")
    assert cell.value == 5  # tiny_sales_order_df: Material 10000001 @ 2910 confirmed qty 5


def test_old_report_allocation_applies_city_pairing_for_3pl_plants():
    # 3PL plants (2925/2935) don't take direct orders -- their Allocation
    # figure must come from their city's TOL plant column instead.
    sales_order_df = pd.DataFrame(
        [
            {**{c: None for c in config.OVERSTOCK_COLUMNS}, "Sales Order": "1", "Sales Order Item": "10",
             "Material": "10000001", "Material Description": "X", "Confirmed Quantity (CS)": 9, "Plant": "2920"},
        ]
    )
    materials_df = pd.DataFrame(
        [
            {
                "Material Group Name": "G", "Material": "10000001", "Material Description": "X",
                "Plant": "2925", "Plant Name": "Lineage Calgary", "Batch": "B1", "Storage Location": "1100",
                "Shelf Life Expiration Date": pd.Timestamp("2027-01-01"), "Unrestricted Stock": 10,
            }
        ]
    )
    allocation = build_allocation(sales_order_df)
    overstock_df = build_overstock(sales_order_df)
    price_list_df = build_price_list(_tiny_pricing_df())
    old_report_df = build_old_report(materials_df, allocation)
    wb = build_workbook(overstock_df, price_list_df, old_report_df, allocation)

    ws = wb["Old report"]
    assert ws.cell(2, 15).value == 9  # 2925's own row picks up 2920's confirmed quantity


def test_workbook_reopens_without_error():
    import io

    wb = _build()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    from openpyxl import load_workbook

    reopened = load_workbook(buf)
    assert reopened.sheetnames == ["Allocation", "Overstock", "Price List", "Old report"]
