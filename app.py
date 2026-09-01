"""
National Overstock Report -- Streamlit interface.

Kept intentionally minimal per spec: title, a report-date input, three
uploaders, a Generate button, validation messages, and a download button. No
charts, KPIs, extra pages, or navigation.

The report date drives the eligible-inventory shelf-life window (see
src/config.INVENTORY_WINDOW_DAYS / src/eligibility.py) and must never be
hardcoded to a past run's date -- it defaults to today but is editable so a
late or backdated run still filters against the correct date.
"""
from datetime import date

import streamlit as st

from src.report_builder import generate_report

st.set_page_config(page_title="National Overstock Report", layout="centered")

st.title("National Overstock Report")

report_date = st.date_input("Report date", value=date.today())
materials_file = st.file_uploader("Materials export", type=["xlsx"])
sales_order_file = st.file_uploader("Open sales-order export", type=["xlsx"])
pricing_file = st.file_uploader("Material/pricing export", type=["xlsx"])

if st.button("Generate Report", type="primary"):
    if not (materials_file and sales_order_file and pricing_file):
        st.error("Please upload all three files: Materials export, Open sales-order export, and Material/pricing export.")
    else:
        with st.spinner("Building report..."):
            result = generate_report(materials_file, sales_order_file, pricing_file, report_date)

        for warning in result.warnings:
            st.warning(warning)

        if not result.success:
            for error in result.errors:
                st.error(error)
        else:
            st.success("Report generated successfully.")
            st.session_state["report_bytes"] = result.workbook_bytes

if "report_bytes" in st.session_state:
    st.download_button(
        "Download National Overstock Report",
        data=st.session_state["report_bytes"],
        file_name=f"National Overstock Report - {date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
