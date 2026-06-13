import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Page config
st.set_page_config(
    page_title="SLMG Inventory Hub",
    page_icon="banner_bg.png",
    layout="wide"
)

# Load data
@st.cache_data
def load_data():
    master = pd.read_excel("Master Stock.xlsx")
    risk = pd.read_excel("Near Expiry iNVENTORY_.xlsx")
    master.columns = master.columns.str.strip()
    risk.columns = risk.columns.str.strip()
    return master, risk

master, risk = load_data()

# ---- Clean up date columns ----
date_cols_master = ["MFG Date", "EXP Date", "DOD Date"]
for col in date_cols_master:
    if col in master.columns:
        master[col] = pd.to_datetime(master[col], errors="coerce").dt.date

date_cols_risk = ["MFG Date", "EXP Date", "BBD Date", "DOD Date", "BBD/Expiry"]
for col in date_cols_risk:
    if col in risk.columns:
        risk[col] = pd.to_datetime(risk[col], errors="coerce").dt.date

# Adjust top padding so heading is visible
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- Data prep ----
master["Shelflife in Days"] = pd.to_numeric(master["Shelflife in Days"], errors="coerce").fillna(0).astype(int)
master["Total"] = pd.to_numeric(master["Total"], errors="coerce").fillna(0).astype(int)

# Unified shelf life status logic
master["SL Status"] = "Safe"
master.loc[master["Shelflife in Days"] < 30, "SL Status"] = "Critical"
master.loc[(master["Shelflife in Days"] >= 31) & (master["Shelflife in Days"] <= 90), "SL Status"] = "Warning"

risk["Quantity"] = pd.to_numeric(risk["Quantity"], errors="coerce").fillna(0).astype(int)
risk["Consumed inventory"] = pd.to_numeric(risk["Consumed inventory"], errors="coerce").fillna(0).astype(int)
risk["Days to BBD"] = pd.to_numeric(risk["Days to BBD"], errors="coerce").fillna(0).astype(int)
risk["Days to SBD"] = pd.to_numeric(risk["Days to SBD"], errors="coerce").fillna(0).astype(int)

# Unified BBD status logic
risk["BBD Status"] = "Safe"
risk.loc[risk["Days to BBD"] < 30, "BBD Status"] = "Critical"
risk.loc[(risk["Days to BBD"] >= 31) & (risk["Days to BBD"] <= 90), "BBD Status"] = "Warning"

def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Heading
st.markdown("<h2 style='text-align:center; font-family:Georgia; font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)
tab1, tab2 = st.tabs(["Stock Overview", "Risk Stock Overview"])
with tab1:
    st.header("Stock Overview")

    # Filters
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    with col1: selected_sites = st.multiselect("Site", sorted(master["Site"].dropna().unique()))
    with col2:
        wh_source = master.copy()
        if selected_sites: wh_source = wh_source[wh_source["Site"].isin(selected_sites)]
        selected_warehouses = st.multiselect("Warehouse", sorted(wh_source["Warehouse"].dropna().unique()))
    with col3: selected_skus = st.multiselect("SKU", sorted(master["Item Description"].dropna().unique()))
    with col4: selected_brands = st.multiselect("Brand", sorted(master["Brand"].dropna().unique()))
    with col5: selected_categories = st.multiselect("Category", sorted(master["Category"].dropna().unique()))
    with col6: selected_pack = st.multiselect("Pack Size", sorted(master["Pack Size"].dropna().unique()))
    with col7: selected_sl = st.multiselect("Shelf Life", ["Critical", "Warning", "Safe"])
    with col8: search_text = st.text_input("Search")

    filtered = master.copy()
    if selected_sites: filtered = filtered[filtered["Site"].isin(selected_sites)]
    if selected_warehouses: filtered = filtered[filtered["Warehouse"].isin(selected_warehouses)]
    if selected_skus: filtered = filtered[filtered["Item Description"].isin(selected_skus)]
    if selected_brands: filtered = filtered[filtered["Brand"].isin(selected_brands)]
    if selected_categories: filtered = filtered[filtered["Category"].isin(selected_categories)]
    if selected_pack: filtered = filtered[filtered["Pack Size"].isin(selected_pack)]
    if selected_sl: filtered = filtered[filtered["SL Status"].isin(selected_sl)]
    if search_text:
        mask = filtered.astype(str).apply(lambda x: x.str.contains(search_text, case=False, na=False)).any(axis=1)
        filtered = filtered[mask]

    st.divider()

    # KPI cards
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Total Stock", f"{int(filtered['Total'].sum()):,}")
    c2.metric("Unique SKUs", int(filtered["Item Description"].nunique()))
    c3.metric("Sites", int(filtered["Site"].nunique()))
    c4.metric("Warehouses", int(filtered["Warehouse"].nunique()))
    c5.metric("Critical SL (<30)", int(filtered[filtered["Shelflife in Days"] < 30]["Total"].sum()))
    c6.metric("Warning SL (31-90)", int(filtered[(filtered["Shelflife in Days"] >= 31) & (filtered["Shelflife in Days"] <= 90)]["Total"].sum()))
    c7.metric("Safe SL (>90)", int(filtered[filtered["Shelflife in Days"] > 90]["Total"].sum()))

    st.divider()

    # Stock by site chart
    stock_site = filtered.groupby("Site")["Total"].sum().reset_index().sort_values("Total", ascending=False)
    fig = px.bar(stock_site, x="Site", y="Total", text="Total", title="Stock By Site")
    st.plotly_chart(fig, use_container_width=True, key="stock_by_site")

    # Top SKUs
    left, right = st.columns(2)
    with left:
        st.subheader("Top 5 SKUs By Inventory")
        top_inventory = filtered.groupby("Item Description")["Total"].sum().reset_index().sort_values("Total", ascending=False).head(5)
        top_inventory["Total"] = top_inventory["Total"].astype(int)
        st.dataframe(top_inventory, hide_index=True, use_container_width=True)
    with right:
        st.subheader("Top 5 SKUs — Least Shelf Life")
        least_sl = filtered[["Item Description", "Shelflife in Days", "Total"]].sort_values("Shelflife in Days").head(5)
        least_sl["Shelflife in Days"] = least_sl["Shelflife in Days"].astype(int)
        least_sl["Total"] = least_sl["Total"].astype(int)
        st.dataframe(least_sl, hide_index=True, use_container_width=True)

    st.divider()
    detail = filtered.copy()
    # Format Shelflife column as percentage string if present
    if "Shelflife" in detail.columns:
        detail["Shelflife"] = (detail["Shelflife"].astype(float) * 100).round(0).astype(int).astype(str) + "%"
    for col in detail.select_dtypes(include=["int64","float64"]).columns:
        if col != "Shelflife":
            detail[col] = detail[col].astype(int)
    st.subheader("Detail Table")
    st.dataframe(detail, hide_index=True, use_container_width=True, height=500)
    st.download_button("Export to Excel", export_excel(detail), file_name="Stock_Overview.xlsx")
with tab2:
    st.header("Risk Stock Overview")

    # Filters
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: selected_sites = st.multiselect("Site", sorted(risk["Unit"].dropna().unique()), key="risk_site")
    with col2:
        wh_source = risk.copy()
        if selected_sites: wh_source = wh_source[wh_source["Unit"].isin(selected_sites)]
        selected_wh = st.multiselect("Warehouse", sorted(wh_source["Warehouse"].dropna().unique()), key="risk_wh")
    with col3: selected_sku = st.multiselect("SKU", sorted(risk["SKU"].dropna().unique()), key="risk_sku")
    with col4: selected_expiry = st.multiselect("Expiry Status", ["Critical", "Warning", "Safe"], key="risk_exp")
    with col5: search_risk = st.text_input("Search", key="risk_search")

    # Apply filters
    rf = risk.copy()
    if selected_sites: rf = rf[rf["Unit"].isin(selected_sites)]
    if selected_wh: rf = rf[rf["Warehouse"].isin(selected_wh)]
    if selected_sku: rf = rf[rf["SKU"].isin(selected_sku)]
    if selected_expiry: rf = rf[rf["BBD Status"].isin(selected_expiry)]
    if search_risk:
        mask = rf.astype(str).apply(lambda x: x.str.contains(search_risk, case=False, na=False)).any(axis=1)
        rf = rf[mask]

    # If no rows left, show message
    if rf.empty:
        st.warning("No data available for the selected filters.")
    else:
        # KPI CARDS
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Qty", f"{int(rf['Quantity'].sum()):,}")
        c2.metric("Critical BBD (<30)", f"{int(rf[rf['Days to BBD'] < 30]['Quantity'].sum()):,}")
        c3.metric("Warning BBD (31-90)", f"{int(rf[(rf['Days to BBD'] >= 31) & (rf['Days to BBD'] <= 90)]['Quantity'].sum()):,}")
        c4.metric("Safe BBD (>90)", f"{int(rf[rf['Days to BBD'] > 90]['Quantity'].sum()):,}")
        c5.metric("Critical SBD (<30)", f"{int(rf[rf['Days to SBD'] < 30]['Quantity'].sum()):,}")

        st.divider()

        # Top 5 At-Risk SKUs
        left, right = st.columns(2)
        with left:
            st.subheader("Top 5 At-Risk BBD SKUs")
            top_bbd = rf[["SKU", "Quantity", "Days to BBD"]].sort_values("Days to BBD").head(5)
            top_bbd["Quantity"] = top_bbd["Quantity"].astype(int)
            top_bbd["Days to BBD"] = top_bbd["Days to BBD"].astype(int)
            st.dataframe(top_bbd, hide_index=True, use_container_width=True)
        with right:
            st.subheader("Top 5 At-Risk SBD SKUs")
            top_sbd = rf[["SKU", "Quantity", "Days to SBD"]].sort_values("Days to SBD").head(5)
            top_sbd["Quantity"] = top_sbd["Quantity"].astype(int)
            top_sbd["Days to SBD"] = top_sbd["Days to SBD"].astype(int)
            st.dataframe(top_sbd, hide_index=True, use_container_width=True)

        st.divider()

        # Plant Level Breakdown
        plant_data = []
        for unit in sorted(rf["Unit"].dropna().unique()):
            temp = rf[rf["Unit"] == unit]
            leakage = temp[temp["Warehouse"].astype(str).str.endswith("_101")]["Quantity"].sum()
            plant_data.append({
                "UNIT": unit,
                "TOTAL QTY": int(temp["Quantity"].sum()),
                "ITEMS": int(temp["SKU"].nunique()),
                "CRITICAL BBD": int(temp[temp["Days to BBD"] < 30]["Quantity"].sum()),
                "WARNING BBD": int(temp[(temp["Days to BBD"] >= 31) & (temp["Days to BBD"] <= 90)]["Quantity"].sum()),
                "SAFE BBD": int(temp[temp["Days to BBD"] > 90]["Quantity"].sum()),
                "CONSUMED INV": int(temp["Consumed inventory"].sum()),
                "LEAKAGE/BREAKAGE WH": int(leakage)
            })
        plant_df = pd.DataFrame(plant_data)
        st.subheader("Plant Level Breakdown")
        st.dataframe(plant_df, hide_index=True, use_container_width=True)

        # Detail Table
        st.subheader("Detail Table")
        for col in rf.select_dtypes(include=["int64","float64"]).columns:
            rf[col] = rf[col].astype(int)
        st.dataframe(rf, hide_index=True, use_container_width=True, height=500)
        st.download_button("Export to Excel", export_excel(rf), file_name="Risk_Overview.xlsx")
