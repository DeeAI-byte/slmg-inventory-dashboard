import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ✅ Page config always comes right after imports
st.set_page_config(page_title="SLMG Inventory Hub", page_icon="banner_bg.png", layout="wide")

# -----------------------------
# Load all four datasets
# -----------------------------
@st.cache_data
def load_data():
    master = pd.read_excel("Master Stock.xlsx")
    risk = pd.read_excel("Near Expiry iNVENTORY_.xlsx")
    distributor = pd.read_excel("DBR.xlsx")
    secondary = pd.read_excel("Secondary.xlsx")   # ✅ added Secondary

    # Strip whitespace from headers
    for df in [master, risk, distributor, secondary]:
        df.columns = df.columns.str.strip()

    # Clean numeric columns in secondary
    if "QTY" in secondary.columns:
        secondary["QTY"] = pd.to_numeric(secondary["QTY"], errors="coerce").fillna(0).astype(int)
    if "NetRevenue" in secondary.columns:
        secondary["NetRevenue"] = pd.to_numeric(secondary["NetRevenue"], errors="coerce").fillna(0).astype(int)

    return master, risk, distributor, secondary

# ✅ Unpack all four datasets right after defining load_data
master, risk, distributor, secondary = load_data()

# -----------------------------
# Date cleanup starts below
# -----------------------------

# Date cleanup
for col in ["MFG Date","EXP Date","DOD Date"]:
    if col in master.columns: master[col] = pd.to_datetime(master[col], errors="coerce").dt.date
for col in ["MFG Date","EXP Date","BBD/Expiry","DOD Date"]:
    if col in risk.columns: risk[col] = pd.to_datetime(risk[col], errors="coerce").dt.date
for col in ["MFG Date","EXP Date","BBD/Expiry"]:
    if col in distributor.columns: distributor[col] = pd.to_datetime(distributor[col], errors="coerce").dt.date

st.markdown("<style>.block-container {padding-top: 1.5rem;}</style>", unsafe_allow_html=True)

# Data prep
master["Quantity"] = pd.to_numeric(master["Quantity"], errors="coerce").fillna(0).astype(int)
if "Shelflife" in master.columns:
    # Convert Excel values to whole numbers (e.g. 78 instead of 0.78)
    master["Shelflife"] = (pd.to_numeric(master["Shelflife"], errors="coerce") * 100).round().astype(int)

    # Status logic uses the numeric Shelflife directly
    master["SL Status"] = "Safe"
    master.loc[master["Shelflife"] < 30, "SL Status"] = "Critical"
    master.loc[(master["Shelflife"] >= 31) & (master["Shelflife"] <= 90), "SL Status"] = "Warning"

risk["Quantity"] = pd.to_numeric(risk["Quantity"], errors="coerce").fillna(0).astype(int)
risk["Consumed inventory"] = pd.to_numeric(risk["Consumed inventory"], errors="coerce").fillna(0).astype(int)
for col in ["Days to BBD","Days to SBD","Days to DOD","Days to LBD"]:
    if col in risk.columns: risk[col] = pd.to_numeric(risk[col], errors="coerce").fillna(0).astype(int)
risk["BBD Status"] = "Safe"
risk.loc[risk["Days to BBD"] < 30, "BBD Status"] = "Critical"
risk.loc[(risk["Days to BBD"] >= 31) & (risk["Days to BBD"] <= 90), "BBD Status"] = "Warning"

distributor["Quantity"] = pd.to_numeric(distributor["Quantity"], errors="coerce").fillna(0).astype(int)
distributor["Days to BBD"] = pd.to_numeric(distributor["Days to BBD"], errors="coerce").fillna(0).astype(int)
distributor["BBD Status"] = "Safe"
distributor.loc[distributor["Days to BBD"] < 30, "BBD Status"] = "Critical"
distributor.loc[(distributor["Days to BBD"] >= 31) & (distributor["Days to BBD"] <= 90), "BBD Status"] = "Warning"

def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.markdown("<h2 style='text-align:center; font-family:Georgia; font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)
tab1, tab2, tab3, tab4 = st.tabs(["Stock Overview", "Risk Stock Overview", "Distributor Stock Overview", "Secondary Sales Overview"])

with tab1:
    st.header("Stock Overview")

    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    with col1: selected_sites = st.multiselect("Site", sorted(master["Site"].dropna().unique()), key="stock_site")
    with col2:
        wh_source = master.copy()
        if selected_sites: wh_source = wh_source[wh_source["Site"].isin(selected_sites)]
        selected_warehouses = st.multiselect("Warehouse", sorted(wh_source["Warehouse"].dropna().unique()), key="stock_wh")
    with col3: selected_skus = st.multiselect("SKU", sorted(master["SKU"].dropna().unique()), key="stock_sku")
    with col4: selected_brands = st.multiselect("Brand", sorted(master["Brand"].dropna().unique()), key="stock_brand")
    with col5: selected_categories = st.multiselect("Category", sorted(master["Category"].dropna().unique()), key="stock_cat")
    with col6: selected_pack = st.multiselect("Pack Size", sorted(master["Pack Size"].dropna().unique()), key="stock_pack")
    with col7: selected_sl = st.multiselect("Shelf Life", ["Critical","Warning","Safe"], key="stock_sl")
    with col8: search_text = st.text_input("Search", key="stock_search")

    filtered = master.copy()
    if selected_sites: filtered = filtered[filtered["Site"].isin(selected_sites)]
    if selected_warehouses: filtered = filtered[filtered["Warehouse"].isin(selected_warehouses)]
    if selected_skus: filtered = filtered[filtered["SKU"].isin(selected_skus)]
    if selected_brands: filtered = filtered[filtered["Brand"].isin(selected_brands)]
    if selected_categories: filtered = filtered[filtered["Category"].isin(selected_categories)]
    if selected_pack: filtered = filtered[filtered["Pack Size"].isin(selected_pack)]
    if selected_sl: filtered = filtered[filtered["SL Status"].isin(selected_sl)]
    if search_text:
        mask = filtered.astype(str).apply(lambda x: x.str.contains(search_text, case=False, na=False)).any(axis=1)
        filtered = filtered[mask]

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Total Stock", f"{int(filtered['Quantity'].sum()):,}")
    c2.metric("Unique SKUs", int(filtered["SKU"].nunique()))
    c3.metric("Sites", int(filtered["Site"].nunique()))
    c4.metric("Warehouses", int(filtered["Warehouse"].nunique()))
    c5.metric("Critical SL (<30%)", int(filtered[filtered["Shelflife"] < 30]["Quantity"].sum()))
    c6.metric("Warning SL (31-90%)", int(filtered[(filtered["Shelflife"] >= 31) & (filtered["Shelflife"] <= 90)]["Quantity"].sum()))
    c7.metric("Safe SL (>=90%)", int(filtered[filtered["Shelflife"] >= 90]["Quantity"].sum()))

    stock_site = filtered.groupby("Site")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
    fig = px.bar(stock_site, x="Site", y="Quantity", text="Quantity", title="Stock By Site")
    st.plotly_chart(fig, use_container_width=True)

    left,right = st.columns(2)
    with left:
        st.subheader("Top 5 SKUs By Inventory")
        top_inventory = filtered.groupby("SKU")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False).head(5)
        st.dataframe(top_inventory, hide_index=True, use_container_width=True)
    with right:
        st.subheader("Top 5 SKUs — Lowest Shelf Life %")
        least_sl = filtered[["SKU","Shelflife","Quantity"]].sort_values("Shelflife").head(5)
        st.dataframe(least_sl, hide_index=True, use_container_width=True)

    st.subheader("Detail Table")
    st.dataframe(filtered, hide_index=True, use_container_width=True, height=500)
    st.download_button("Export to Excel", export_excel(filtered), file_name="Stock_Overview.xlsx")
with tab2:
    st.header("Risk Stock Overview")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: selected_sites = st.multiselect("Site", sorted(risk["Unit"].dropna().unique()), key="risk_site")
    with col2:
        wh_source = risk.copy()
        if selected_sites: wh_source = wh_source[risk["Unit"].isin(selected_sites)]
        selected_wh = st.multiselect("Warehouse", sorted(wh_source["Warehouse"].dropna().unique()), key="risk_wh")
    with col3: selected_sku = st.multiselect("SKU", sorted(risk["SKU"].dropna().unique()), key="risk_sku")
    with col4: selected_expiry = st.multiselect("Expiry Status", ["Critical","Warning","Safe"], key="risk_exp")
    with col5: search_risk = st.text_input("Search", key="risk_search")

    rf = risk.copy()
    if selected_sites: rf = rf[rf["Unit"].isin(selected_sites)]
    if selected_wh: rf = rf[rf["Warehouse"].isin(selected_wh)]
    if selected_sku: rf = rf[rf["SKU"].isin(selected_sku)]
    if selected_expiry: rf = rf[rf["BBD Status"].isin(selected_expiry)]
    if search_risk:
        mask = rf.astype(str).apply(lambda x: x.str.contains(search_risk, case=False, na=False)).any(axis=1)
        rf = rf[mask]

    if rf.empty:
        st.warning("No data available for the selected filters.")
    else:
        # KPI Cards
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Qty", f"{int(rf['Quantity'].sum()):,}")
        c2.metric("Critical BBD (<30 days)", int(rf[rf["Days to BBD"] < 30]["Quantity"].sum()))
        c3.metric("Warning BBD (31-90 days)", int(rf[(rf["Days to BBD"] >= 31) & (rf["Days to BBD"] <= 90)]["Quantity"].sum()))
        c4.metric("Safe BBD (>90 days)", int(rf[rf["Days to BBD"] > 90]["Quantity"].sum()))
        c5.metric("Critical SBD (<30 days)", int(rf[rf["Days to SBD"] < 30]["Quantity"].sum()))
        c6.metric("Critical LBD (<30 days)", int(rf[rf["Days to LBD"] < 30]["Quantity"].sum()))

        st.divider()

        # Top 5 At-Risk SKUs
        left, right = st.columns(2)
        with left:
            st.subheader("Top 5 At-Risk BBD SKUs")
            top_bbd = rf[["SKU","Quantity","Days to BBD"]].sort_values("Days to BBD").head(5)
            st.dataframe(top_bbd, hide_index=True, use_container_width=True)
        with right:
            st.subheader("Top 5 At-Risk SBD SKUs")
            top_sbd = rf[["SKU","Quantity","Days to SBD"]].sort_values("Days to SBD").head(5)
            st.dataframe(top_sbd, hide_index=True, use_container_width=True)

        st.divider()

        # Plant Level Breakdown
        breakdown = []
        for unit in sorted(rf["Unit"].dropna().unique()):
            temp = rf[rf["Unit"] == unit]
            leakage = temp[temp["Warehouse"].astype(str).str.endswith("_101")]["Quantity"].sum()
            breakdown.append({
                "UNIT": unit,
                "TOTAL QTY": int(temp["Quantity"].sum()),
                "ITEMS": int(temp["SKU"].nunique()),
                "CRITICAL BBD": int(temp[temp["Days to BBD"] < 30]["Quantity"].sum()),
                "WARNING BBD": int(temp[(temp["Days to BBD"] >= 31) & (temp["Days to BBD"] <= 90)]["Quantity"].sum()),
                "SAFE BBD": int(temp[temp["Days to BBD"] > 90]["Quantity"].sum()),
                "CRITICAL SBD": int(temp[temp["Days to SBD"] < 30]["Quantity"].sum()),
                "CONSUMED INV": int(temp["Consumed inventory"].sum()),
                "LEAKAGE/BREAKAGE WH": int(leakage)
            })
        plant_df = pd.DataFrame(breakdown)
        st.subheader("Plant Level Breakdown")
        st.dataframe(plant_df, hide_index=True, use_container_width=True)

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(rf, hide_index=True, use_container_width=True, height=500)
        st.download_button("Export to Excel", export_excel(rf), file_name="Risk_Overview.xlsx")
with tab3:
    st.header("Distributor Stock Overview")

    # Filters in one line (District, Distributor, SKU, Brand, Pack Size, Expiry Status, Search)
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    with col1: selected_districts = st.multiselect("District", sorted(distributor["District"].dropna().unique()), key="dist_district")
    with col2:
        dist_source = distributor.copy()
        if selected_districts: dist_source = dist_source[dist_source["District"].isin(selected_districts)]
        selected_distributors = st.multiselect("Distributor", sorted(dist_source["Distributor"].dropna().unique()), key="dist_distributor")
    with col3: selected_sku = st.multiselect("SKU", sorted(distributor["SKU"].dropna().unique()), key="dist_sku")
    with col4: selected_brand = st.multiselect("Brand", sorted(distributor["Brand"].dropna().unique()), key="dist_brand")
    with col5: selected_pack = st.multiselect("Pack Size", sorted(distributor["Pack Size"].dropna().unique()), key="dist_pack")
    with col6: selected_expiry = st.multiselect("Expiry Status", ["Critical","Warning","Safe"], key="dist_exp")
    with col7: search_dist = st.text_input("Search", key="dist_search")

    # Apply filters
    df = distributor.copy()
    if selected_districts: df = df[df["District"].isin(selected_districts)]
    if selected_distributors: df = df[df["Distributor"].isin(selected_distributors)]
    if selected_sku: df = df[df["SKU"].isin(selected_sku)]
    if selected_brand: df = df[df["Brand"].isin(selected_brand)]
    if selected_pack: df = df[df["Pack Size"].isin(selected_pack)]
    if selected_expiry: df = df[df["BBD Status"].isin(selected_expiry)]
    if search_dist:
        mask = df.astype(str).apply(lambda x: x.str.contains(search_dist, case=False, na=False)).any(axis=1)
        df = df[mask]

    if df.empty:
        st.warning("No data available for the selected filters.")
    else:
        # KPI Cards
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        c1.metric("Total Qty", f"{int(df['Quantity'].sum()):,}")
        c2.metric("Total Distributors", int(df["Distributor"].nunique()))
        c3.metric("Unique SKUs", int(df["SKU"].nunique()))
        c4.metric("Districts", int(df["District"].nunique()))
        c5.metric("Critical BBD (<30 days)", int(df[df["Days to BBD"] < 30]["Quantity"].sum()))
        c6.metric("Warning BBD (31-90 days)", int(df[(df["Days to BBD"] >= 31) & (df["Days to BBD"] <= 90)]["Quantity"].sum()))
        c7.metric("Safe BBD (>90 days)", int(df[df["Days to BBD"] > 90]["Quantity"].sum()))

        st.divider()

        # Stock by District chart
        stock_district = df.groupby("District")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
        fig_dist = px.bar(stock_district, x="District", y="Quantity", text="Quantity", title="Stock by District")
        st.plotly_chart(fig_dist, use_container_width=True)

        st.divider()

        # Stock by Brand and Stock by Pack Size charts side by side
        colA, colB = st.columns(2)

        with colA:
            stock_brand = df.groupby("Brand")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_brand = px.bar(stock_brand, x="Brand", y="Quantity", text="Quantity", title="Stock by Brand")
            st.plotly_chart(fig_brand, use_container_width=True)

        with colB:
            stock_pack = df.groupby("Pack Size")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_pack = px.bar(stock_pack, x="Pack Size", y="Quantity", text="Quantity", title="Stock by Pack Size")
            st.plotly_chart(fig_pack, use_container_width=True)

        st.divider()

        # District Level Breakdown
        breakdown = []
        for district in sorted(df["District"].dropna().unique()):
            temp = df[df["District"] == district]
            for distributor_name in sorted(temp["Distributor"].dropna().unique()):
                sub = temp[temp["Distributor"] == distributor_name]
                breakdown.append({
                    "District": district,
                    "Distributor": distributor_name,
                    "Total Qty": int(sub["Quantity"].sum()),
                    "Items": int(sub["SKU"].nunique()),
                    "Critical BBD": int(sub[sub["Days to BBD"] < 30]["Quantity"].sum()),
                    "Warning BBD": int(sub[(sub["Days to BBD"] >= 31) & (sub["Days to BBD"] <= 90)]["Quantity"].sum()),
                    "Safe BBD": int(sub[sub["Days to BBD"] > 90]["Quantity"].sum())
                })
        breakdown_df = pd.DataFrame(breakdown)
        st.subheader("District Level Breakdown")
        st.dataframe(breakdown_df, hide_index=True, use_container_width=True)

        st.divider()
        st.subheader("Detail Table")
        for col in df.select_dtypes(include=["int64","float64"]).columns:
            df[col] = df[col].astype(int)
        # Detail table includes Brand and Pack Size automatically
        st.dataframe(df, hide_index=True, use_container_width=True, height=500)
        st.download_button("Export to Excel", export_excel(df), file_name="Distributor_Overview.xlsx")
with tab4:
    st.header("Secondary Sales Overview")
    
     # ✅ use the already-loaded dataframe
    df = secondary.copy()

    # Filters in one line
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    with col1: selected_district = st.multiselect("District", sorted(secondary["District"].dropna().unique()), key="sec_district")
    with col2: selected_sm = st.multiselect("SM", sorted(secondary["SM"].dropna().unique()), key="sec_sm")
    with col3: selected_asm = st.multiselect("ASM", sorted(secondary["ASM"].dropna().unique()), key="sec_asm")
    with col4: selected_route = st.multiselect("Route", sorted(secondary["Route"].dropna().unique()), key="sec_route")
    with col5: selected_distributor = st.multiselect("Distributor", sorted(secondary["Distributor"].dropna().unique()), key="sec_distributor")
    with col6: selected_brand = st.multiselect("Brand", sorted(secondary["Brand"].dropna().unique()), key="sec_brand")
    with col7: selected_category = st.multiselect("Category", sorted(secondary["Category"].dropna().unique()), key="sec_category")
    with col8: selected_pack = st.multiselect("Pack Size", sorted(secondary["Pack Size"].dropna().unique()), key="sec_pack")
    search_sec = st.text_input("Search", key="sec_search")

    # Apply filters
    df = secondary.copy()
    if selected_district: df = df[df["District"].isin(selected_district)]
    if selected_sm: df = df[df["SM"].isin(selected_sm)]
    if selected_asm: df = df[df["ASM"].isin(selected_asm)]
    if selected_route: df = df[df["Route"].isin(selected_route)]
    if selected_distributor: df = df[df["Distributor"].isin(selected_distributor)]
    if selected_brand: df = df[df["Brand"].isin(selected_brand)]
    if selected_category: df = df[df["Category"].isin(selected_category)]
    if selected_pack: df = df[df["Pack Size"].isin(selected_pack)]
    if search_sec:
        mask = df.astype(str).apply(lambda x: x.str.contains(search_sec, case=False, na=False)).any(axis=1)
        df = df[mask]

    # Output section
    if df.empty:
        st.warning("No data available for the selected filters.")
    else:
        # KPI Cards
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Outlets", int(df["Outlet"].nunique()))
        c2.metric("Total Secondary Sales Volume (QTY)", f"{int(df['QTY'].sum()):,}")
        c3.metric("Total Secondary Sales Revenue", f"{int(df['NetRevenue'].sum()):,}")
        c4.metric("Unique SKUs", int(df["ITEMCODE"].nunique()))

        st.divider()

        # Charts
        colA, colB = st.columns(2)
        with colA:
            asm_perf = df.groupby("ASM")[["QTY","NetRevenue"]].sum().reset_index().sort_values("QTY", ascending=False)
            fig_asm = px.bar(asm_perf, x="ASM", y="QTY", text="QTY", title="ASM Performance (Volume)")
            st.plotly_chart(fig_asm, width=800)
        with colB:
            brand_perf = df.groupby("Brand")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_brand = px.bar(brand_perf, x="Brand", y="QTY", text="QTY", title="Brand Performance")
            st.plotly_chart(fig_brand, width=800)

        st.divider()
        colC, colD = st.columns(2)
        with colC:
            cat_perf = df.groupby("Category")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_cat = px.bar(cat_perf, x="Category", y="QTY", text="QTY", title="Category Performance")
            st.plotly_chart(fig_cat, width=800)
        with colD:
            pack_perf = df.groupby("Pack Size")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_pack = px.bar(pack_perf, x="Pack Size", y="QTY", text="QTY", title="Pack Size Performance")
            st.plotly_chart(fig_pack, width=800)

        st.divider()
        vpo_contrib = df.groupby("VPO")["QTY"].sum().reset_index()
        fig_vpo = px.pie(vpo_contrib, names="VPO", values="QTY", title="VPO Contribution")
        st.plotly_chart(fig_vpo, width=800)

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(df.head(1000), hide_index=True, use_container_width=True, height=500)
        st.download_button("Export to Excel", export_excel(df), file_name="Secondary_Sales_Overview.xlsx")
