import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="SLMG Inventory Hub", page_icon="banner_bg.png", layout="wide")

# -----------------------------
# Loader function
# -----------------------------
def load_file(file_path: str):
    if file_path.endswith(".parquet"):
        return pd.read_parquet(file_path)
    elif file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    else:
        return pd.read_excel(file_path)

# -----------------------------
# Memory optimization helper
# -----------------------------
def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """Convert low-cardinality string columns to category and downcast numerics.
    This can cut memory usage by 80-90% on typical wide, repetitive datasets.
    Uses pd.api.types helpers instead of dtype == object comparisons, which
    is safer across pandas 2.x and 3.x (pandas 3.0 changed string internals)."""
    if df.empty:
        return df

    n = len(df)
    for col in df.columns:
        try:
            if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col]):
                nunique = df[col].nunique()
                if n > 0 and nunique / n < 0.5:
                    df[col] = df[col].astype("category")
            elif pd.api.types.is_integer_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], downcast="integer")
            elif pd.api.types.is_float_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], downcast="float")
        except Exception:
            pass  # skip any column that causes issues
    return df

# -----------------------------
# Load all datasets via dictionary
# -----------------------------
@st.cache_data
def load_data():
    files = {
        "master": "Master Stock.xlsx",
        "risk": "Near Expiry iNVENTORY_.xlsx",
        "distributor": "DBR.xlsx",
        "secondary": "Secondary.parquet"
    }
    datasets = {}
    for key, path in files.items():
        try:
            df = load_file(path)
            df.columns = df.columns.str.strip()
            datasets[key] = df
        except Exception:
            # If file missing, create empty DataFrame with expected columns
            if key == "secondary":
                datasets[key] = pd.DataFrame(columns=[
                    "District","SM","ASM","Route","Distributor","Outlet",
                    "Brand","Category","Pack Size","ITEMCODE","QTY","NetRevenue",
                    "VPO","CustomerHierarchy"
                ])
            else:
                datasets[key] = pd.DataFrame()

    # Clean numeric columns in secondary (handle Qty vs QTY column name variants)
    sec = datasets["secondary"]
    qty_key = "Qty" if "Qty" in sec.columns else "QTY"
    if qty_key in sec.columns:
        sec[qty_key] = pd.to_numeric(sec[qty_key], errors="coerce").fillna(0).astype(int)
    if "NetRevenue" in sec.columns:
        sec["NetRevenue"] = pd.to_numeric(sec["NetRevenue"], errors="coerce").fillna(0).astype(int)

    # Apply memory optimization to every dataset (runs once, cached)
    for key in datasets:
        datasets[key] = optimize_dtypes(datasets[key])

    return datasets

datasets = load_data()
master, risk, distributor, secondary = (
    datasets["master"],
    datasets["risk"],
    datasets["distributor"],
    datasets["secondary"]
)

# -----------------------------
# Date cleanup
# -----------------------------
for col in ["MFG Date","EXP Date","DOD Date"]:
    if col in master.columns:
        master[col] = pd.to_datetime(master[col], errors="coerce", dayfirst=True).dt.date

for col in ["MFG Date","EXP Date","BBD/Expiry","DOD Date"]:
    if col in risk.columns:
        risk[col] = pd.to_datetime(risk[col], errors="coerce", dayfirst=True).dt.date

# Create BBD Status for risk from "Days to BBD" (this was missing, which
# caused a KeyError whenever the Expiry Status filter was used)
if "Days to BBD" in risk.columns:
    risk["Days to BBD"] = pd.to_numeric(risk["Days to BBD"], errors="coerce")
    risk["BBD Status"] = "Safe"
    risk.loc[risk["Days to BBD"] < 30, "BBD Status"] = "Critical"
    risk.loc[(risk["Days to BBD"] >= 31) & (risk["Days to BBD"] <= 90), "BBD Status"] = "Warning"

# Round numeric columns to whole numbers (Quantity, Consumed inventory,
# Days to DOD/LBD/SBD/BBD, etc. come in from Excel with long decimals).
# Uses include="number" instead of explicit int64/float64 dtype names,
# since optimize_dtypes() downcasts to float32/int32 etc. earlier, and
# those wouldn't match a literal int64/float64 check.
for col in risk.select_dtypes(include="number").columns:
    risk[col] = pd.to_numeric(risk[col], errors="coerce").fillna(0).round().astype(int)

for col in ["MFG Date","BBD/Expiry"]:
    if col in distributor.columns:
        distributor[col] = pd.to_datetime(distributor[col], errors="coerce", dayfirst=True).dt.date

# 👇 Handle EXP Date separately (outside the loop)
if "EXP Date" in distributor.columns:
    # Keep raw Excel values visible
    distributor["EXP Date"] = distributor["EXP Date"].astype(str).str.strip()

    # Calculate Days to BBD safely
    distributor["Days to BBD"] = (
        pd.to_datetime(distributor["EXP Date"], errors="coerce", dayfirst=True) - pd.to_datetime("today")
    ).dt.days

    # Create BBD Status from Days to BBD
    distributor["BBD Status"] = "Safe"
    distributor.loc[distributor["Days to BBD"] < 30, "BBD Status"] = "Critical"
    distributor.loc[(distributor["Days to BBD"] >= 31) & (distributor["Days to BBD"] <= 90), "BBD Status"] = "Warning"

st.markdown("<style>.block-container {padding-top: 1.5rem;}</style>", unsafe_allow_html=True)

master["Quantity"] = pd.to_numeric(master["Quantity"], errors="coerce").fillna(0).astype(int)
if "Shelflife" in master.columns:
    master["Shelflife"] = (pd.to_numeric(master["Shelflife"], errors="coerce") * 100).round().astype(int)
    master["SL Status"] = "Safe"
    master.loc[master["Shelflife"] < 30, "SL Status"] = "Critical"
    master.loc[(master["Shelflife"] >= 31) & (master["Shelflife"] <= 90), "SL Status"] = "Warning"

def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def lazy_export_section(df, file_name: str, button_key: str):
    """Only builds the Excel file in memory when the user explicitly asks for it,
    instead of regenerating it on every rerun of every tab."""
    if st.button(f"Prepare Export ({file_name})", key=f"{button_key}_prepare"):
        st.session_state[f"{button_key}_bytes"] = export_excel(df)

    cached_bytes = st.session_state.get(f"{button_key}_bytes")
    if cached_bytes is not None:
        st.download_button(
            "Download Excel",
            cached_bytes,
            file_name=file_name,
            key=f"{button_key}_download"
        )

st.markdown("<h2 style='text-align:center; font-family:Georgia; font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)
tab1, tab2, tab3, tab4 = st.tabs(["Stock Overview", "Risk Stock Overview", "Distributor Stock Overview", "Secondary Sales Overview"])
with tab1:
    st.header("Stock Overview")

    # Filters
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

    # ✅ Initialize filtered before use
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

    # KPIs
    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Total Stock", f"{int(filtered['Quantity'].sum()):,}")
    c2.metric("Unique SKUs", int(filtered["SKU"].nunique()))
    c3.metric("Sites", int(filtered["Site"].nunique()))
    c4.metric("Warehouses", int(filtered["Warehouse"].nunique()))
    c5.metric("Critical SL (<30%)", int(filtered[filtered["Shelflife"] < 30]["Quantity"].sum()))
    c6.metric("Warning SL (31-90%)", int(filtered[(filtered["Shelflife"] >= 31) & (filtered["Shelflife"] <= 90)]["Quantity"].sum()))
    c7.metric("Safe SL (>=90%)", int(filtered[filtered["Shelflife"] >= 90]["Quantity"].sum()))

    # Chart
    stock_site = filtered.groupby("Site")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
    fig = px.bar(stock_site, x="Site", y="Quantity", text="Quantity", title="Stock By Site")
    fig.update_traces(texttemplate='%{text:,}', textposition='outside')
    fig.update_yaxes(tickformat="d")
    st.plotly_chart(fig, width="stretch")

    # Tables
    left,right = st.columns(2)
    with left:
        st.subheader("Top 5 SKUs By Inventory")
        top_inventory = filtered.groupby("SKU")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False).head(5)
        st.dataframe(top_inventory, hide_index=True, width="stretch")
    with right:
        st.subheader("Top 5 SKUs — Lowest Shelf Life %")
        least_sl = filtered[["SKU","Shelflife","Quantity"]].sort_values("Shelflife").head(5)
        st.dataframe(least_sl, hide_index=True, width="stretch")

    st.subheader("Detail Table")
    st.dataframe(filtered, hide_index=True, width="stretch", height=500)
    lazy_export_section(filtered, "Stock_Overview.xlsx", "stock")
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
    with col4: selected_expiry = st.multiselect("Expiry Status", ["Critical","Warning","Safe"], key="risk_exp")
    with col5: search_risk = st.text_input("Search", key="risk_search")

    # ✅ Initialize rf before use
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
            st.dataframe(top_bbd, hide_index=True, width="stretch")
        with right:
            st.subheader("Top 5 At-Risk SBD SKUs")
            top_sbd = rf[["SKU","Quantity","Days to SBD"]].sort_values("Days to SBD").head(5)
            st.dataframe(top_sbd, hide_index=True, width="stretch")

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
        st.dataframe(plant_df, hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(rf, hide_index=True, width="stretch", height=500)
        lazy_export_section(rf, "Risk_Overview.xlsx", "risk")
with tab3:
    st.header("Distributor Stock Overview")

    # Filters
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

    # ✅ Initialize df before use
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

        # Charts with integer formatting
        stock_district = df.groupby("District")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
        fig_dist = px.bar(stock_district, x="District", y="Quantity",
                          text=stock_district["Quantity"].astype(int),
                          title="Stock by District")
        fig_dist.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_dist.update_yaxes(tickformat="d")
        st.plotly_chart(fig_dist, width="stretch")

        colA, colB = st.columns(2)
        with colA:
            stock_brand = df.groupby("Brand")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_brand = px.bar(stock_brand, x="Brand", y="Quantity",
                               text=stock_brand["Quantity"].astype(int),
                               title="Stock by Brand")
            fig_brand.update_traces(texttemplate='%{text:,}', textposition='outside')
            fig_brand.update_yaxes(tickformat="d")
            st.plotly_chart(fig_brand, width="stretch")
        with colB:
            stock_pack = df.groupby("Pack Size")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_pack = px.bar(stock_pack, x="Pack Size", y="Quantity",
                              text=stock_pack["Quantity"].astype(int),
                              title="Stock by Pack Size")
            fig_pack.update_traces(texttemplate='%{text:,}', textposition='outside')
            fig_pack.update_yaxes(tickformat="d")
            st.plotly_chart(fig_pack, width="stretch")

        st.divider()
        st.subheader("District Level Breakdown")
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
        st.dataframe(breakdown_df, hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        for col in df.select_dtypes(include=["int64","float64"]).columns:
            df[col] = df[col].fillna(0).astype(int)  # safer cast
        st.dataframe(df, hide_index=True, width="stretch", height=500)
        lazy_export_section(df, "Distributor_Overview.xlsx", "dist")
with tab4:
    st.header("Secondary Sales Overview")

    brand_filter_col = "Brands" if "Brands" in secondary.columns else "Brand"

    # Row 1: 8 cascading dimension filters
    cascade_columns = [
        ("District", "sec_district"),
        ("SM", "sec_sm"),
        ("ASM", "sec_asm"),
        ("Route", "sec_route"),
        ("Distributor", "sec_distributor"),
        (brand_filter_col, "sec_brand"),
        ("Category", "sec_category"),
        ("Pack Size", "sec_pack"),
    ]

    cols = st.columns(8)
    source = secondary.copy()
    for (col_name, widget_key), col in zip(cascade_columns, cols):
        with col:
            options = sorted(source[col_name].dropna().unique())
            selected_values = st.multiselect(col_name, options, key=widget_key)
        if selected_values:
            source = source[source[col_name].isin(selected_values)]

    # Row 2: Month, Date filters + Search (compact layout)
    f1, f2, f3, f4 = st.columns([1, 2, 2, 3])
    with f1:
        month_options = sorted(source["Month"].dropna().unique()) if "Month" in source.columns else []
        selected_months = st.multiselect("Month", month_options, key="sec_month")
    with f2:
        date_options = sorted(source["Date"].dropna().unique()) if "Date" in source.columns else []
        selected_dates = st.multiselect("Date", date_options, key="sec_date")
    with f3:
        search_sec = st.text_input("Search", key="sec_search")
    with f4:
        pass  # intentional spacer to keep search compact

    if selected_months:
        source = source[source["Month"].isin(selected_months)]
    if selected_dates:
        source = source[source["Date"].isin(selected_dates)]

    df = source.copy()
    if search_sec:
        mask = df.astype(str).apply(lambda x: x.str.contains(search_sec, case=False, na=False)).any(axis=1)
        df = df[mask]

    # ✅ Always show KPIs/charts (zeros if empty)
    # Handle column name variants between parquet versions
    qty_col = "Qty" if "Qty" in df.columns else "QTY"
    outlet_col = "Outlet Code" if "Outlet Code" in df.columns else "Outlet"

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Outlets", int(df[outlet_col].nunique()))
    c2.metric("Total Secondary Sales Volume (QTY)", f"{int(df[qty_col].sum()):,}")
    c3.metric("Total Secondary Sales Revenue", f"{int(df['NetRevenue'].sum()):,}")

    # NSR (Revenue per Unit) — whole number
    total_qty = df[qty_col].sum()
    total_rev = df["NetRevenue"].sum()
    nsr = round(total_rev / total_qty) if total_qty > 0 else 0
    c4.metric("NSR (Revenue per Unit)", f"{int(nsr):,}")

    # Avg IPS = average number of unique IPS items per outlet
    # Uses nunique() on the IPS column grouped by outlet — matches pivot table logic
    if "IPS" in df.columns and not df.empty:
        avg_ips = round(df.groupby(outlet_col)["IPS"].nunique().mean())
    else:
        avg_ips = 0
    c5.metric("Avg IPS (Items per Store)", int(avg_ips))

    st.divider()

    # Charts
    brand_col = "Brands" if "Brands" in df.columns else "Brand"
    colA, colB = st.columns(2)
    with colA:
        asm_perf = df.groupby("ASM")[[qty_col,"NetRevenue"]].sum().reset_index().sort_values(qty_col, ascending=False)
        asm_perf[qty_col] = asm_perf[qty_col].round().astype(int)
        fig_asm = px.bar(asm_perf, x="ASM", y=qty_col, text=qty_col, title="ASM Performance (Volume)")
        fig_asm.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_asm.update_yaxes(tickformat="d")
        st.plotly_chart(fig_asm, width="stretch")
    with colB:
        brand_perf = df.groupby(brand_col)[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        brand_perf[qty_col] = brand_perf[qty_col].round().astype(int)
        fig_brand = px.bar(brand_perf, x=brand_col, y=qty_col, text=qty_col, title="Brand Performance")
        fig_brand.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_brand.update_yaxes(tickformat="d")
        st.plotly_chart(fig_brand, width="stretch")

    st.divider()
    colC, colD = st.columns(2)
    with colC:
        cat_perf = df.groupby("Category")[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        cat_perf[qty_col] = cat_perf[qty_col].round().astype(int)
        fig_cat = px.bar(cat_perf, x="Category", y=qty_col, text=qty_col, title="Category Performance")
        fig_cat.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_cat.update_yaxes(tickformat="d")
        st.plotly_chart(fig_cat, width="stretch")
    with colD:
        pack_perf = df.groupby("Pack Size")[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        pack_perf[qty_col] = pack_perf[qty_col].round().astype(int)
        fig_pack = px.bar(pack_perf, x="Pack Size", y=qty_col, text=qty_col, title="Pack Size Performance")
        fig_pack.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_pack.update_yaxes(tickformat="d")
        st.plotly_chart(fig_pack, width="stretch")

    st.divider()
    colE, colF = st.columns(2)
    with colE:
        vpo_contrib = df.groupby("VPO")[qty_col].sum().reset_index()
        vpo_contrib[qty_col] = vpo_contrib[qty_col].round().astype(int)
        fig_vpo = px.pie(vpo_contrib, names="VPO", values=qty_col, title="VPO Contribution")
        fig_vpo.update_traces(texttemplate='%{percent:.1%}', textinfo='percent+label')
        st.plotly_chart(fig_vpo, width="stretch")
    with colF:
        cust_contrib = df.groupby("CustomerHierarchy")[qty_col].sum().reset_index()
        cust_contrib[qty_col] = cust_contrib[qty_col].round().astype(int)
        fig_cust = px.pie(cust_contrib, names="CustomerHierarchy", values=qty_col, title="CustomerHierarchy Contribution")
        fig_cust.update_traces(texttemplate='%{percent:.1%}', textinfo='percent+label')
        st.plotly_chart(fig_cust, width="stretch")

    st.divider()
    st.subheader("Detail Table")
    st.dataframe(df.head(1000), hide_index=True, width="stretch", height=500)
    lazy_export_section(df, "Secondary_Sales_Overview.xlsx", "sec")
