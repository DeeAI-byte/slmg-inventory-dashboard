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
        df = load_file(path)
        df.columns = df.columns.str.strip()
        datasets[key] = df

    # Clean numeric columns in secondary
    sec = datasets["secondary"]
    if "QTY" in sec.columns:
        sec["QTY"] = pd.to_numeric(sec["QTY"], errors="coerce").fillna(0).astype(int)
    if "NetRevenue" in sec.columns:
        sec["NetRevenue"] = pd.to_numeric(sec["NetRevenue"], errors="coerce").fillna(0).astype(int)

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
    if col in master.columns: master[col] = pd.to_datetime(master[col], errors="coerce").dt.date
for col in ["MFG Date","EXP Date","BBD/Expiry","DOD Date"]:
    if col in risk.columns: risk[col] = pd.to_datetime(risk[col], errors="coerce").dt.date
for col in ["MFG Date","EXP Date","BBD/Expiry"]:
    if col in distributor.columns: distributor[col] = pd.to_datetime(distributor[col], errors="coerce").dt.date

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

st.markdown("<h2 style='text-align:center; font-family:Georgia; font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)
tab1, tab2, tab3, tab4 = st.tabs(["Stock Overview", "Risk Stock Overview", "Distributor Stock Overview", "Secondary Sales Overview"])

with tab1:
    st.header("Stock Overview")
    # ... filters code unchanged ...

    stock_site = filtered.groupby("Site")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
    fig = px.bar(stock_site, x="Site", y="Quantity", text="Quantity", title="Stock By Site")
    st.plotly_chart(fig, width="stretch")

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
    st.download_button("Export to Excel", export_excel(filtered), file_name="Stock_Overview.xlsx")
with tab2:
    st.header("Risk Stock Overview")

    # ... filters code unchanged ...

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
        plant_df = pd.DataFrame(breakdown)
        st.subheader("Plant Level Breakdown")
        st.dataframe(plant_df, hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(rf, hide_index=True, width="stretch", height=500)
        st.download_button("Export to Excel", export_excel(rf), file_name="Risk_Overview.xlsx")
with tab3:
    st.header("Distributor Stock Overview")

    # ... filters code unchanged ...

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
        st.plotly_chart(fig_dist, width="stretch")

        st.divider()

        # Stock by Brand and Pack Size charts
        colA, colB = st.columns(2)
        with colA:
            stock_brand = df.groupby("Brand")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_brand = px.bar(stock_brand, x="Brand", y="Quantity", text="Quantity", title="Stock by Brand")
            st.plotly_chart(fig_brand, width="stretch")
        with colB:
            stock_pack = df.groupby("Pack Size")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_pack = px.bar(stock_pack, x="Pack Size", y="Quantity", text="Quantity", title="Stock by Pack Size")
            st.plotly_chart(fig_pack, width="stretch")

        st.divider()
        st.subheader("District Level Breakdown")
        st.dataframe(breakdown_df, hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(df, hide_index=True, width="stretch", height=500)
        st.download_button("Export to Excel", export_excel(df), file_name="Distributor_Overview.xlsx")

with tab4:
    st.header("Secondary Sales Overview")

    # ... filters code unchanged ...

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
            st.plotly_chart(fig_asm, width="stretch")
        with colB:
            brand_perf = df.groupby("Brand")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_brand = px.bar(brand_perf, x="Brand", y="QTY", text="QTY", title="Brand Performance")
            st.plotly_chart(fig_brand, width="stretch")

        st.divider()
        colC, colD = st.columns(2)
        with colC:
            cat_perf = df.groupby("Category")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_cat = px.bar(cat_perf, x="Category", y="QTY", text="QTY", title="Category Performance")
            st.plotly_chart(fig_cat, width="stretch")
        with colD:
            pack_perf = df.groupby("Pack Size")["QTY"].sum().reset_index().sort_values("QTY", ascending=False)
            fig_pack = px.bar(pack_perf, x="Pack Size", y="QTY", text="QTY", title="Pack Size Performance")
            st.plotly_chart(fig_pack, width="stretch")

        st.divider()
        # Pie charts side by side
        colE, colF = st.columns(2)
        with colE:
            vpo_contrib = df.groupby("VPO")["QTY"].sum().reset_index()
            fig_vpo = px.pie(vpo_contrib, names="VPO", values="QTY", title="VPO Contribution")
            st.plotly_chart(fig_vpo, width="stretch")
        with colF:
            cust_contrib = df.groupby("CustomerHierarchy")["QTY"].sum().reset_index()
            fig_cust = px.pie(cust_contrib, names="CustomerHierarchy", values="QTY", title="CustomerHierarchy Contribution")
            st.plotly_chart(fig_cust, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(df.head(1000), hide_index=True, width="stretch", height=500)
        st.download_button("Export to Excel", export_excel(df), file_name="Secondary_Sales_Overview.xlsx")
