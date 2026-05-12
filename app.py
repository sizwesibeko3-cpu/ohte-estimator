import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

st.title("OHTE Tender Estimator")

# =========================
# RATE LIBRARY
# =========================
rates = {
    "Cantilever": {"A":22000, "B":35000, "S":55000},
    "64Kn Steel Mast": {"A":150000, "B":245000, "S":310000},
    "85Kn Steel Mast": {"A":190000, "B":315000, "S":395000},
    "Concrete Mast 64Kn": {"A":130000, "B":175000, "S":225000},
    "Rail Bond": {"A":2000, "B":2800, "S":4200},
    "Foundation 64Kn": {"A":70000, "B":90000, "S":115000},
    "Lattice Bridge": {"A":800000, "B":1350000, "S":1900000}
}

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader("Upload BOQ Excel (Item, Unit, Qty)", type=["xlsx"])

if uploaded_file:
    boq = pd.read_excel(uploaded_file)
    boq.columns = [c.strip().title() for c in boq.columns]

    st.write("### BOQ Preview")
    st.dataframe(boq)

    # =========================
    # STRATEGY SELECTOR
    # =========================
    mode = st.selectbox("Select Pricing Mode", ["A (Aggressive)", "B (Balanced)", "S (Safe)"])
    mode = mode[0]

    # =========================
    # PRICING ENGINE
    # =========================
    def get_rate(item):
        item = str(item).strip().title()
        if item in rates:
            return rates[item][mode]
        return 0

    boq["Rate"] = boq["Item"].apply(get_rate)
    boq["Total"] = boq["Qty"] * boq["Rate"]

    st.write("### Priced BOQ")
    st.dataframe(boq)

    total = boq["Total"].sum()
    st.success(f"Total Tender Value: {total:,.2f}")

    # =========================
    # COMPETITOR SIMULATION
    # =========================
    comp = boq.copy()
    comp["Aggressive"] = comp["Qty"] * comp["Item"].apply(lambda x: rates.get(x.title(),{"A":0})["A"])
    comp["Balanced"] = comp["Qty"] * comp["Item"].apply(lambda x: rates.get(x.title(),{"B":0})["B"])
    comp["Tier-1"] = comp["Qty"] * comp["Item"].apply(lambda x: rates.get(x.title(),{"S":0})["S"])

    st.write("### Competitor Comparison")
    st.dataframe(comp)

    # =========================
    # EXPORT EXCEL
    # =========================
    output = BytesIO()
    wb = Workbook()

    ws = wb.active
    ws.title = "Priced_BOQ"
    ws.append(list(boq.columns))
    for row in boq.values:
        ws.append(list(row))

    ws2 = wb.create_sheet("Competitor")
    ws2.append(list(comp.columns))
    for row in comp.values:
        ws2.append(list(row))

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Priced BOQ Excel",
        data=output,
        file_name="OHTE_Estimator_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
