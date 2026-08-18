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
# Number formats (copied verbatim from the reference workbook's cells)
# ---------------------------------------------------------------------------
FMT_GENERAL = "General"
FMT_DATE = "mm/dd/yyyy"
FMT_CS_3DP = '0.000\\ "CS"'
FMT_CS_0DP = '0\\ "CS"'
FMT_KG_2DP = '0.00\\ "KG"'
FMT_KG_3DP = '0.000\\ "KG"'
FMT_QTY_3DP_PLAIN = "0.000"
FMT_CURRENCY = '* #,##0.00_0_0_0\\ "CAD";* \\-\\ #,##0.00_0_0_0\\ "CAD"'
FMT_PACK_INT = "0"
FMT_YES_NO = '"Yes";;"No"'

OVERSTOCK_NUMBER_FORMATS = {
    "Creation Date": FMT_DATE,
    "Order Quantity (CS)": FMT_CS_3DP,
    "Confirmed Quantity (CS)": FMT_CS_3DP,
    "Picked Quantity (CS)": FMT_CS_3DP,
    "Invoice Quantity (CS)": FMT_CS_3DP,
    "Confirmed Quantity (KG)": FMT_KG_2DP,
    "Invoice Item Amount": FMT_CURRENCY,
    "Invoice Quantity (KG)": FMT_QTY_3DP_PLAIN,
    "Item Net Amount": FMT_CURRENCY,
    "Item Net Amount (Confirmed)": FMT_CURRENCY,
    "Order Quantity": FMT_CS_3DP,
    "Order Quantity (KG)": FMT_KG_2DP,
    "Picked Quantity (KG)": FMT_KG_3DP,
    "Item Value": FMT_CURRENCY,
    "Requested Delivery Date": FMT_DATE,
}

PRICE_LIST_NUMBER_FORMATS = {
    "Pack": FMT_PACK_INT,
    "Pricing Date": FMT_DATE,
    "1 - Store Door": FMT_CURRENCY,
    "2 - Wholesale": FMT_CURRENCY,
    "Sold by Wt": FMT_YES_NO,
}

OLD_REPORT_NUMBER_FORMATS = {
    "Shelf Life Expiration Date": FMT_DATE,
    "Unrestricted Stock": FMT_CS_0DP,
    "Available": FMT_CS_0DP,
}

# ---------------------------------------------------------------------------
# Layout: column widths, row heights, freeze panes, autofilter ranges
# ---------------------------------------------------------------------------
COLUMN_WIDTHS = {
    "Allocation": {
        "A": 13.625, "B": 11.875, "C": 8.75, "D": 10.0, "E": 9.75,
        "F": 9.625, "G": 13.375, "H": 14.0,
    },
    "Overstock": {
        "A": 10.5, "B": 10.75, "C": 9.0, "D": 6.125, "E": 16.25, "F": 9.5,
        "G": 9.875, "H": 16.25, "I": 11.25, "J": 11.25, "K": 23.25,
        "L": 11.875, "M": 15.25, "N": 12.875, "P": 11.0, "Q": 12.125,
        "R": 13.5, "S": 14.625, "T": 13.625, "U": 16.875, "W": 12.5,
        "AA": 33.875, "AB": 19.25, "AC": 16.0, "AD": 14.75,
    },
    "Price List": {
        "A": 26.375, "B": 18.875, "C": 11.375, "D": 81.625, "E": 7.25,
        "F": 11.375, "G": 15.25, "H": 17.25, "I": 18.0, "J": 8.25, "K": 13.25,
    },
    "Old report": {
        "A": 21.875, "B": 22.25, "C": 10.25, "D": 33.875, "E": 8.5,
        "F": 7.875, "G": 11.875, "H": 11.75, "I": 17.375, "J": 14.5,
        "K": 10.125, "L": 14.75, "M": 13.625, "N": 14.0, "O": 10.125,
    },
}

HEADER_ROW_HEIGHT = {
    "Overstock": 42.0,
    "Old report": 27.95,
}

FREEZE_PANES = {
    "Allocation": None,
    "Overstock": "A2",
    "Price List": "A264",
    "Old report": "A2",
}

# Autofilter ranges are re-derived at write time from actual row/col counts
# except for "Old report", whose reference autofilter stops at column M
# (Unrestricted Stock) and deliberately excludes Unique Key/Allocation/
# Available.
OLD_REPORT_AUTOFILTER_LAST_COLUMN = "M"

HEADER_FILL_RGB = "FFF7F7F7"
FONT_NAME = "Arial"
FONT_SIZE = 11

# Per-header horizontal alignment, copied verbatim from the reference
# workbook (None = Excel's default, left-aligned for text). Every Overstock
# and Old report header wraps text; Price List headers do not.
OVERSTOCK_HEADER_ALIGN = {
    "Sales Order": None,
    "Sales Order Item": "center",
    "Key Account #": None,
    "Sales Order Type": None,
    "Creation Date": "center",
    "Order Status": "center",
    "Rejection Status": "center",
    "Sold To Party": None,
    "Ship To Party": None,
    "Material": None,
    "Material Description": None,
    "Order Quantity (CS)": "center",
    "Confirmed Quantity (CS)": "center",
    "Picked Quantity (CS)": "center",
    "Invoice Quantity (CS)": "center",
    "Invoice #": "center",
    "Outbound Delivery #": "center",
    "Confirmed Quantity (KG)": "center",
    "Invoice Item Amount": "center",
    "Invoice Quantity (KG)": "center",
    "Item Net Amount": "center",
    "Item Net Amount (Confirmed)": "center",
    "Order Quantity": "center",
    "Order Quantity (KG)": "center",
    "Picked Quantity (KG)": "center",
    "Plant": "center",
    "Ship To Name": "left",
    "BBD / Shelf Life": None,
    "Item Value": None,
    "Requested Delivery Date": None,
    "Ship to Arrive Date": None,
}

PRICE_LIST_HEADER_ALIGN = {
    "BDM": None,
    "Brand Name": None,
    "Material": None,
    "Description": None,
    "Pack": "center",
    "Size": "center",
    "Pricing Date": "center",
    "1 - Store Door": "center",
    "2 - Wholesale": "center",
    "Plant": "center",
    "Sold by Wt": "center",
}
