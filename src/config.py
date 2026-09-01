"""
Constants that define the National Overstock Report output structure.

Every column list, plant mapping, and number format here was reverse-engineered
directly from the reference workbook
"Overstock National report - Aug 4 2026 (1).xlsx" so the generated report
matches it exactly. Do not add/remove/reorder entries without re-checking the
reference file.

The eligible-inventory / eligible-sales-order rules below were separately
reverse-engineered against a second pair of reference files for the
2026-08-24 run: "National Stock Overstock Aug 24 2026.xlsx to seperate and
send (1).xlsx" (manually-prepared, authoritative) vs. "National Overstock
Report - 2026-08-24 (1).xlsx" (the dashboard's un-filtered output at the
time). See src/eligibility.py and src/fifo_allocation.py for how they're
applied.
"""

# ---------------------------------------------------------------------------
# Materials export -> required source columns
# ---------------------------------------------------------------------------
# These are copied straight through into the "Old report" sheet (columns
# B, C, D, H, I, J, K, L). "Material" also doubles as the join key used to
# look BDM/Sold by KG/Pack/Size/DSD Price up in the Price List sheet.
MATERIALS_REQUIRED_COLUMNS = [
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
    "Allocated",
    "Available",
    "DSD Price",
    "Deal Price",
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

# ---------------------------------------------------------------------------
# Eligible-inventory / eligible-sales-order rules
# ---------------------------------------------------------------------------
# Reverse-engineered against "National Stock Overstock Aug 24 2026.xlsx to
# seperate and send (1).xlsx" (the manually-prepared, authoritative report)
# vs. "National Overstock Report - 2026-08-24 (1).xlsx" (the dashboard's
# un-filtered output). A material is in scope for the report only if ALL of:
#
#   1. Its Material code does not start with this prefix -- these are
#      packaging/shipper SKUs (e.g. "50016409 BLUE DIAMOND EMPTY CARDBOARD
#      SHIPPER"), never sellable stock.
PACKAGING_MATERIAL_PREFIX = "50"

#   2. Its BDM (looked up from the Material/pricing export by Material,
#      first match) is not one of these reps. Verified against the Aug 24
#      reference: every material handled by these three BDMs was excluded
#      from the manual report's inventory and sales-order sheets with zero
#      exceptions across their full material lists (RANA FOOD SERVICE,
#      JONDAY, FRESCA MEXICAN FOODS, PUTTERS FS, BRICKMANS, and 9 other
#      single-material brands all resolve to just these 3 BDMs), while every
#      *other* BDM appears in both the kept and excluded rows. So exclusion
#      is by BDM (a small, stable, reusable list), never by brand / Material
#      Group Name (an unbounded, brand-specific list that would need
#      updating every time the business adds a new food-service SKU).
EXCLUDED_FOOD_SERVICE_BDMS = {"MIGUEL GUTIERREZ", "Karishma Salian", "JEREMY MAINGOT PR"}

#   3. It has at least one row in the Material/pricing export at all --
#      materials absent from Price List entirely (e.g. 10013786, "C2O COCO
#      WATER GNGR LME TURM") were excluded from the Aug 24 reference despite
#      passing both rules above. (Enforced in src/eligibility.py directly
#      against the Price List DataFrame; no separate constant needed.)

# Inventory rows are further restricted to a rolling shelf-life window
# starting at the report date (see src/eligibility.filter_eligible_inventory
# -- the report date is a Streamlit date_input, never hardcoded). Verified
# exactly against the Aug 24 reference (report date 2026-08-24): all 271
# reference rows have a Shelf Life Expiration Date in [8/24, 11/25]
# inclusive, and that window explains 95 of the un-filtered dashboard's 145
# extra rows with zero false positives/negatives.
INVENTORY_WINDOW_DAYS = 93

# Deal Price = DSD Price x (1 - 75%) = DSD Price x 0.25, per the reference
# workbook's own "=P{r}*(1-0.75)" formula.
DEAL_PRICE_DISCOUNT = 0.75

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
# column L (Unrestricted Stock), excluding Unique Key/Allocated/Available/
# DSD Price/Deal Price.
OLD_REPORT_AUTOFILTER_LAST_COLUMN = "L"
