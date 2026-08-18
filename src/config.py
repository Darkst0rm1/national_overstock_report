"""
Constants that define the National Overstock Report output structure.

Every column list, plant mapping, and number format here was reverse-engineered
directly from the reference workbook
"Overstock National report - Aug 4 2026 (1).xlsx" so the generated report
matches it exactly. Do not add/remove/reorder entries without re-checking the
reference file.
"""

# ---------------------------------------------------------------------------
# Materials export -> required source columns
# ---------------------------------------------------------------------------
# These are copied straight through into the "Old report" sheet (columns
# B, C, D, H, I, J, K, L, M). "Material" also doubles as the join key used to
# look BDM/Sold by KG/Pack/Size up in the Price List sheet.
MATERIALS_REQUIRED_COLUMNS = [
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

# Identifier columns that must never lose leading zeros / be silently
# coerced to numbers -> always read as text.
MATERIALS_TEXT_COLUMNS = ["Material", "Plant", "Batch", "Storage Location"]

# ---------------------------------------------------------------------------
# Open sales-order export -> required source columns / Overstock sheet order
# ---------------------------------------------------------------------------
# This exact order is also the required column order for the "Overstock"
# output sheet (task spec: "Copy the complete open sales-order export into
# this worksheet" in this precise order).
OVERSTOCK_COLUMNS = [
    "Sales Order",
    "Sales Order Item",
    "Key Account #",
    "Sales Order Type",
    "Creation Date",
    "Order Status",
    "Rejection Status",
    "Sold To Party",
    "Ship To Party",
    "Material",
    "Material Description",
    "Order Quantity (CS)",
    "Confirmed Quantity (CS)",
    "Picked Quantity (CS)",
    "Invoice Quantity (CS)",
    "Invoice #",
    "Outbound Delivery #",
    "Confirmed Quantity (KG)",
    "Invoice Item Amount",
    "Invoice Quantity (KG)",
    "Item Net Amount",
    "Item Net Amount (Confirmed)",
    "Order Quantity",
    "Order Quantity (KG)",
    "Picked Quantity (KG)",
    "Plant",
    "Ship To Name",
    "BBD / Shelf Life",
    "Item Value",
    "Requested Delivery Date",
    "Ship to Arrive Date",
]

# Identifier columns in the sales-order export that must stay text.
SALES_ORDER_TEXT_COLUMNS = [
    "Sales Order",
    "Sales Order Item",
    "Key Account #",
    "Sold To Party",
    "Ship To Party",
    "Material",
    "Invoice #",
    "Outbound Delivery #",
    "Plant",
]

# Date columns parsed as datetimes.
SALES_ORDER_DATE_COLUMNS = ["Creation Date", "Requested Delivery Date", "Ship to Arrive Date"]

# Numeric measure columns (everything else is left as text/object).
SALES_ORDER_NUMERIC_COLUMNS = [
    "Order Quantity (CS)",
    "Confirmed Quantity (CS)",
    "Picked Quantity (CS)",
    "Invoice Quantity (CS)",
    "Confirmed Quantity (KG)",
    "Invoice Item Amount",
    "Invoice Quantity (KG)",
    "Item Net Amount",
    "Item Net Amount (Confirmed)",
    "Order Quantity",
    "Order Quantity (KG)",
    "Picked Quantity (KG)",
    "Item Value",
]

# Fields used by the Allocation pivot (must exist in the sales-order export).
ALLOCATION_SOURCE_COLUMNS = ["Material", "Plant", "Confirmed Quantity (CS)"]

# ---------------------------------------------------------------------------
# Material/pricing export -> Price List sheet
# ---------------------------------------------------------------------------
# Exact required output order.
PRICE_LIST_COLUMNS = [
    "BDM",
    "Brand Name",
    "Material",
    "Description",
    "Pack",
    "Size",
    "Pricing Date",
    "1 - Store Door",
    "2 - Wholesale",
    "Plant",
    "Sold by Wt",
]

# Only "Material" is a hard requirement (it is the join key). Every other
# target column is copied through if present under its exact name, and left
# blank otherwise per spec ("If a required field is not present in the
# uploaded source, leave it blank... do not invent information").
PRICE_LIST_HARD_REQUIRED_COLUMNS = ["Material"]

PRICING_TEXT_COLUMNS = ["Material", "Plant"]
PRICING_DATE_COLUMNS = ["Pricing Date"]

# ---------------------------------------------------------------------------
# Old report sheet
# ---------------------------------------------------------------------------
OLD_REPORT_COLUMNS = [
    "BDM",
    "Material Group Name",
    "Material",
    "Material Description",
    "Sold by KG",
    "Pack",
    "Size",
    "Plant",
    "Plant Name",
    "Batch",
    "Storage Location",
    "Shelf Life Expiration Date",
    "Unrestricted Stock",
    "Unique Key",
    "Allocation",
    "Available",
]

# Business rule confirmed against the reference workbook: 3PL storage plants
# (Lineage) don't take direct sales orders, so their "Allocation" figure is
# pulled from the confirmed open-order quantity of the TOL selling plant in
# the same city, not from their own (nonexistent) order data. Any plant not
# listed here maps to itself (i.e. look up its own plant column).
PLANT_TO_SELLING_PLANT = {
    "2910": "2910",  # TOL Mississauga
    "2920": "2920",  # TOL Calgary
    "2925": "2920",  # Lineage Calgary CS -> TOL Calgary
    "2930": "2930",  # TOL Surrey
    "2935": "2930",  # Lineage Surrey CS -> TOL Surrey
}

# Unique Key suffix used by both the Allocation sheet and the Old report
# sheet in the reference workbook. It is a literal constant (not the row's
# actual plant) -- the join between the two sheets is effectively by
# Material only, since the suffix is identical on both sides. Reproduced
# verbatim from the reference file's formulas.
UNIQUE_KEY_SUFFIX = "2910"

# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
# All fonts, fills, borders, number formats, column widths, row heights, and
# freeze panes are inherited automatically from assets/report_template.xlsx
# (a copy of the reference workbook), since excel_writer.py edits that file
# in place rather than rebuilding a workbook from scratch. Only structural
# facts that can't be inherited -- because they depend on this run's row
# count -- are listed here.

# The reference workbook's "Old report" autofilter deliberately stops at
# column M (Unrestricted Stock), excluding Unique Key/Allocation/Available.
OLD_REPORT_AUTOFILTER_LAST_COLUMN = "M"
