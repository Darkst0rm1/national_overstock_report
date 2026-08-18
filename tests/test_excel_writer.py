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


def test_allocation_sheet_has_no_stale_trailing_columns():
    wb = _build()
    ws = wb["Allocation"]
    # One plant (2910) -> B=2910, C=Grand Total, D=Unique Key; nothing past D.
    assert ws.cell(4, 5).value is None
    assert ws.cell(5, 5).value is None


def test_old_report_formula_references_correct_dynamic_allocation_column():
    wb = _build()
    ws = wb["Old report"]
    formula = ws.cell(2, 15).value
    assert "Allocation!D:D" in formula  # Unique Key column for a single-plant run
    assert "Allocation!B:B" in formula  # 2910's own column


def test_workbook_reopens_without_error():
    import io

    wb = _build()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    from openpyxl import load_workbook

    reopened = load_workbook(buf)
    assert reopened.sheetnames == ["Allocation", "Overstock", "Price List", "Old report"]
