import streamlit as st
import pandas as pd
import plotly.express as px
# from io import BytesIO  # removed: unused import to reduce memory footprint
import tempfile
import os
import calendar

st.set_page_config(page_title="SLMG Inventory Hub", page_icon="banner_bg.png", layout="wide")

# -----------------------------
# Loader
# -----------------------------
def load_file(file_path: str):
    if file_path.endswith(".parquet"):
        return pd.read_parquet(file_path)
    elif file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    else:
        return pd.read_excel(file_path)

# -----------------------------
# Memory optimization (Excel files only)
# -----------------------------
def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    n = len(df)
    for col in df.columns:
        try:
            if df[col].dtype == object:
                if n > 0 and df[col].nunique() / n < 0.5:
                    df[col] = df[col].astype("category")
            elif df[col].dtype == "int64":
                df[col] = pd.to_numeric(df[col], downcast="integer")
            elif df[col].dtype == "float64":
                df[col] = pd.to_numeric(df[col], downcast="float")
        except Exception:
            pass
    return df

# -----------------------------
# Export helper
# -----------------------------
def export_excel_to_tempfile(df, file_name: str) -> str:
    """Write the dataframe to a temporary Excel file and return its path.
    This avoids storing large byte arrays in session_state; only the file path is kept.
    Caller is responsible for cleaning up the file if desired.
    """
    # Use suffix from provided file_name
    suffix = os.path.splitext(file_name)[1] or '.xlsx'
    fd, path = tempfile.mkstemp(suffix=suffix, prefix='export_')
    os.close(fd)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return path


def lazy_export_section(df, file_name: str, button_key: str):
    # When preparing export, write to a temp file and store the path in session_state
    if st.button(f"Prepare Export ({file_name})", key=f"{button_key}_prepare"):
        # If a previous temp file exists, remove it first
        prev = st.session_state.get(f"{button_key}_file")
        try:
            if prev and os.path.exists(prev):
                os.remove(prev)
        except Exception:
            pass
        st.session_state[f"{button_key}_file"] = export_excel_to_tempfile(df, file_name)

    file_path = st.session_state.get(f"{button_key}_file")
    if file_path is not None and os.path.exists(file_path):
        # Read bytes on demand for download; do not keep bytes in session_state
        with open(file_path, 'rb') as f:
            data = f.read()
        st.download_button("Download Excel", data, file_name=file_name, key=f"{button_key}_download")
        # After embedding the bytes into the download button, remove the temp file and session_state reference
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass
        try:
            del st.session_state[f"{button_key}_file"]
        except Exception:
            pass

# -----------------------------
# Load & prep all data — runs ONCE, cached forever
# -----------------------------
@st.cache_data
def load_data():
    files = {
        "master":      "Master Stock.xlsx",
        "risk":        "Near Expiry iNVENTORY_.xlsx",
        "distributor": "DBR.xlsx",
        "secondary":   "Secondary.parquet"
    }
    datasets = {}
    for key, path in files.items():
        try:
            df = load_file(path)
            df.columns = df.columns.str.strip()
            datasets[key] = df
        except Exception:
            datasets[key] = pd.DataFrame()

    # ── MASTER ──────────────────────────────────────────────────────────
    m = datasets["master"]
    if not m.empty:
        for col in ["MFG Date", "EXP Date", "DOD Date"]:
            if col in m.columns:
                # keep as datetime64[ns] normalized to date (midnight) to avoid Python object dtype
                m[col] = pd.to_datetime(m[col], errors="coerce", dayfirst=True).dt.normalize()
        if "Quantity" in m.columns:
            tmp_qty = pd.to_numeric(m["Quantity"], errors="coerce").fillna(0)
            m["Quantity"] = pd.to_numeric(tmp_qty, downcast="integer")
        if "Shelflife" in m.columns:
            tmp_sl = (pd.to_numeric(m["Shelflife"], errors="coerce") * 100).round().fillna(0)
            m["Shelflife"] = pd.to_numeric(tmp_sl, downcast="integer")
            m["SL Status"] = "Safe"
            m.loc[m["Shelflife"] < 30, "SL Status"] = "Critical"
            m.loc[(m["Shelflife"] >= 31) & (m["Shelflife"] <= 90), "SL Status"] = "Warning"
        datasets["master"] = optimize_dtypes(m)

    # ── RISK ─────────────────────────────────────────────────────────────
    r = datasets["risk"]
    if not r.empty:
        for col in ["MFG Date", "EXP Date", "BBD/Expiry", "DOD Date"]:
            if col in r.columns:
                r[col] = pd.to_datetime(r[col], errors="coerce", dayfirst=True).dt.normalize()
        if "Days to BBD" in r.columns:
            r["Days to BBD"] = pd.to_numeric(r["Days to BBD"], errors="coerce")
            r["BBD Status"] = "Safe"
            r.loc[r["Days to BBD"] < 30, "BBD Status"] = "Critical"
            r.loc[(r["Days to BBD"] >= 31) & (r["Days to BBD"] <= 90), "BBD Status"] = "Warning"
        for col in r.select_dtypes(include="number").columns:
            tmp = pd.to_numeric(r[col], errors="coerce").fillna(0).round()
            r[col] = pd.to_numeric(tmp, downcast="integer")
        datasets["risk"] = optimize_dtypes(r)

    # ── DISTRIBUTOR ───────────────────────────────────────────────────────
    d = datasets["distributor"]
    if not d.empty:
        for col in ["MFG Date", "BBD/Expiry"]:
            if col in d.columns:
                d[col] = pd.to_datetime(d[col], errors="coerce", dayfirst=True).dt.normalize()
        if "EXP Date" in d.columns:
            d["EXP Date"] = d["EXP Date"].astype(str).str.strip()
            d["Days to BBD"] = (
                pd.to_datetime(d["EXP Date"], errors="coerce", dayfirst=True) - pd.to_datetime("today")
            ).dt.days
            d["BBD Status"] = "Safe"
            d.loc[d["Days to BBD"] < 30, "BBD Status"] = "Critical"
            d.loc[(d["Days to BBD"] >= 31) & (d["Days to BBD"] <= 90), "BBD Status"] = "Warning"
        datasets["distributor"] = optimize_dtypes(d)

    # ── SECONDARY (parquet) — apply dtype optimization too and prepare search text
    sec = datasets["secondary"]
    if not sec.empty:
        qty_key = "Qty" if "Qty" in sec.columns else "QTY"
        if qty_key in sec.columns:
            tmp_qty = pd.to_numeric(sec[qty_key], errors="coerce").fillna(0)
            sec[qty_key] = pd.to_numeric(tmp_qty, downcast="integer")
        if "NetRevenue" in sec.columns:
            tmp_rev = pd.to_numeric(sec["NetRevenue"], errors="coerce").fillna(0)
            sec["NetRevenue"] = pd.to_numeric(tmp_rev, downcast="integer")
        # Apply dtype optimizations (convert low-cardinality strings to category, downcast nums)
        sec = optimize_dtypes(sec)
        datasets["secondary"] = sec

    # Do NOT build a persistent _search_text to avoid large persistent memory usage.
    # Search text will be built lazily per filtered dataframe when the user performs a search.

    return datasets

datasets    = load_data()
master      = datasets["master"]
risk        = datasets["risk"]
distributor = datasets["distributor"]
secondary   = datasets["secondary"]

from functools import lru_cache

# Caching helpers for expensive aggregations based on filtered index fingerprint.
# These helpers accept an index tuple (tuple of ints) to identify which rows are in the filtered set.
@lru_cache(maxsize=128)
def cached_groupby_sum(idx_tuple, dataset_name, group_col, qty_col):
    """Return groupby sum DataFrame for the requested dataset slice identified by idx_tuple.
    idx_tuple: tuple of index labels included in the filtered DataFrame
    dataset_name: one of 'secondary','master','risk','distributor'
    """
    ds = {'secondary': secondary, 'master': master, 'risk': risk, 'distributor': distributor}[dataset_name]
    if not idx_tuple:
        # empty selection
        return pd.DataFrame(columns=[group_col, qty_col])
    sub = ds.loc[list(idx_tuple)]
    if group_col not in sub.columns or qty_col not in sub.columns:
        return pd.DataFrame(columns=[group_col, qty_col])
    res = sub.groupby(group_col, observed=True)[qty_col].sum().reset_index()
    return res

# Cache unique option lists for filter widgets
@lru_cache(maxsize=64)
def cached_unique_options(idx_tuple, dataset_name, column_name):
    ds = {'secondary': secondary, 'master': master, 'risk': risk, 'distributor': distributor}[dataset_name]
    if not idx_tuple:
        arr = ds[column_name].dropna().unique() if column_name in ds.columns else []
    else:
        sub = ds.loc[list(idx_tuple)]
        arr = sub[column_name].dropna().unique() if column_name in sub.columns else []
    try:
        return sorted(arr)
    except Exception:
        return list(arr)

# -----------------------------
# UI
# -----------------------------
st.markdown("<style>.block-container{padding-top:1.5rem;}</style>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align:center;font-family:Georgia;font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["Stock Overview", "Risk Stock Overview", "Distributor Stock Overview", "Secondary Sales Overview"])

# ═════════════════════════════════════════════════════════════════════==
# TAB 1 — Stock Overview
# ═════════════════════════════════════════════════════════════════════==
with tab1:
    st.header("Stock Overview")

    col1,col2,col3,col4,col5,col6,col7,col8 = st.columns(8)
    with col1: sel_sites   = st.multiselect("Site",      sorted(master["Site"].dropna().unique()),      key="stock_site")
    with col2:
        wh_opts = master.loc[master["Site"].isin(sel_sites), "Warehouse"] if sel_sites else master["Warehouse"]
        sel_wh  = st.multiselect("Warehouse", sorted(wh_opts.dropna().unique()), key="stock_wh")
    with col3: sel_skus    = st.multiselect("SKU",       sorted(master["SKU"].dropna().unique()),       key="stock_sku")
    with col4: sel_brands  = st.multiselect("Brand",     sorted(master["Brand"].dropna().unique()),     key="stock_brand")
    with col5: sel_cats    = st.multiselect("Category",  sorted(master["Category"].dropna().unique()),  key="stock_cat")
    with col6: sel_pack    = st.multiselect("Pack Size", sorted(master["Pack Size"].dropna().unique()), key="stock_pack")
    with col7: sel_sl      = st.multiselect("Shelf Life",["Critical","Warning","Safe"],                 key="stock_sl")
    with col8: search_text = st.text_input("Search", key="stock_search")

    fmask = pd.Series(True, index=master.index)
    if sel_sites:  fmask &= master["Site"].isin(sel_sites)
    if sel_wh:     fmask &= master["Warehouse"].isin(sel_wh)
    if sel_skus:   fmask &= master["SKU"].isin(sel_skus)
    if sel_brands: fmask &= master["Brand"].isin(sel_brands)
    if sel_cats:   fmask &= master["Category"].isin(sel_cats)
    if sel_pack:   fmask &= master["Pack Size"].isin(sel_pack)
    if sel_sl:     fmask &= master["SL Status"].isin(sel_sl)
    filtered = master[fmask]
    if search_text:
        # Build searchable text lazily for the filtered dataframe (single combined string per row)
        text_cols = filtered.select_dtypes(include=["object","category"]).columns.tolist()
        if text_cols:
            search_series = filtered[text_cols].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            filtered = filtered[search_series.str.contains(search_text.lower(), na=False)]
        else:
            filtered = filtered.head(0)

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    total_stock_qty = int(filtered['Quantity'].sum())
    c1.metric("Total Stock",        f"{total_stock_qty:,}")
    c2.metric("Unique SKUs",        int(filtered["SKU"].nunique()))
    c3.metric("Sites",              int(filtered["Site"].nunique()))
    c4.metric("Warehouses",         int(filtered["Warehouse"].nunique()))
    # Compute shelf-life buckets once
    sl_critical = filtered[filtered["Shelflife"] < 30]["Quantity"].sum()
    sl_warning = filtered[(filtered["Shelflife"] >= 31) & (filtered["Shelflife"] <= 90)]["Quantity"].sum()
    sl_safe = filtered[filtered["Shelflife"] >= 90]["Quantity"].sum()
    c5.metric("Critical SL (<30%)", int(sl_critical))
    c6.metric("Warning SL (31-90%)", int(sl_warning))
    c7.metric("Safe SL (>=90%)",    int(sl_safe))

    stock_site = filtered.groupby("Site", observed=True)["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
    stock_site_text = pd.to_numeric(stock_site["Quantity"].round(), downcast='integer')
    fig = px.bar(stock_site, x="Site", y="Quantity", text=stock_site_text, title="Stock By Site")
    def _style_bar(fig):
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_yaxes(tickformat="d")
        return fig
    def _style_pie(fig):
        fig.update_traces(texttemplate='%{percent:.1%}', textinfo='percent+label')
        return fig
    fig = _style_bar(fig)
    st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)
    with left:
        st.subheader("Top 5 SKUs By Inventory")
        top_skus = filtered.groupby("SKU", observed=True)["Quantity"].sum().reset_index()
        top5 = top_skus.nlargest(5, "Quantity") if not top_skus.empty else top_skus
        st.dataframe(top5, hide_index=True, use_container_width=True)
    with right:
        st.subheader("Top 5 SKUs — Lowest Shelf Life %")
        if "Shelflife" in filtered.columns:
            low5 = filtered[["SKU","Shelflife","Quantity"]].nsmallest(5, "Shelflife")
        else:
            low5 = filtered[["SKU","Quantity"]].head(0)
        st.dataframe(low5, hide_index=True, use_container_width=True)

    st.subheader("Detail Table")
    st.dataframe(filtered, hide_index=True, use_container_width=True, height=500)
    lazy_export_section(filtered, "Stock_Overview.xlsx", "stock")

# ═════════════════════════════════════════════════════════════════════==
# TAB 2 — Risk Stock Overview
# ═════════════════════════════════════════════════════════════════════==
with tab2:
    st.header("Risk Stock Overview")

    col1,col2,col3,col4,col5 = st.columns(5)
    with col1: sel_sites  = st.multiselect("Site",          sorted(risk["Unit"].dropna().unique()),      key="risk_site")
    with col2:
        wh_opts = risk.loc[risk["Unit"].isin(sel_sites), "Warehouse"] if sel_sites else risk["Warehouse"]
        sel_wh  = st.multiselect("Warehouse", sorted(wh_opts.dropna().unique()), key="risk_wh")
    with col3: sel_sku    = st.multiselect("SKU",           sorted(risk["SKU"].dropna().unique()),       key="risk_sku")
    with col4: sel_expiry = st.multiselect("Expiry Status", ["Critical","Warning","Safe"],               key="risk_exp")
    with col5: search_risk= st.text_input("Search", key="risk_search")

    rmask = pd.Series(True, index=risk.index)
    if sel_sites:  rmask &= risk["Unit"].isin(sel_sites)
    if sel_wh:     rmask &= risk["Warehouse"].isin(sel_wh)
    if sel_sku:    rmask &= risk["SKU"].isin(sel_sku)
    if sel_expiry: rmask &= risk["BBD Status"].isin(sel_expiry)
    rf = risk[rmask]
    if search_risk:
        text_cols = rf.select_dtypes(include=["object","category"]).columns.tolist()
        if text_cols:
            search_series = rf[text_cols].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            rf = rf[search_series.str.contains(search_risk.lower(), na=False)]
        else:
            rf = rf.head(0)

    if rf.empty:
        st.warning("No data available for the selected filters.")
    else:
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Total Qty",              f"{int(rf['Quantity'].sum()):,}")
        c2.metric("Critical BBD (<30d)",    int(rf[rf["Days to BBD"] < 30]["Quantity"].sum()))
        c3.metric("Warning BBD (31-90d)",   int(rf[(rf["Days to BBD"] >= 31) & (rf["Days to BBD"] <= 90)]["Quantity"].sum()))
        c4.metric("Safe BBD (>90d)",        int(rf[rf["Days to BBD"] > 90]["Quantity"].sum()))
        c5.metric("Critical SBD (<30d)",    int(rf[rf["Days to SBD"] < 30]["Quantity"].sum()) if "Days to SBD" in rf else 0)
        c6.metric("Critical LBD (<30d)",    int(rf[rf["Days to LBD"] < 30]["Quantity"].sum()) if "Days to LBD" in rf else 0)

        st.divider()
        left, right = st.columns(2)
        with left:
            st.subheader("Top 5 At-Risk BBD SKUs")
            top5_bbd = rf.nsmallest(5, 'Days to BBD')[['SKU','Quantity','Days to BBD']] if 'Days to BBD' in rf.columns else rf[['SKU','Quantity']].head(0)
            st.dataframe(top5_bbd, hide_index=True, use_container_width=True)
        with right:
            st.subheader("Top 5 At-Risk SBD SKUs")
            if "Days to SBD" in rf.columns:
                top5_sbd = rf.nsmallest(5, 'Days to SBD')[['SKU','Quantity','Days to SBD']]
            else:
                top5_sbd = rf[['SKU','Quantity']].head(0)
            st.dataframe(top5_sbd, hide_index=True, use_container_width=True)

        st.divider()
        # Plant Level Breakdown using groupby aggregations (avoids repeated slicing)
        groups = rf.groupby("Unit", observed=True)
        total_qty = groups["Quantity"].sum()
        items = groups["SKU"].nunique()
        critical_bbd = rf[rf["Days to BBD"] < 30].groupby("Unit", observed=True)["Quantity"].sum() if "Days to BBD" in rf.columns else pd.Series(dtype="int")
        warning_bbd = rf[(rf["Days to BBD"] >= 31) & (rf["Days to BBD"] <= 90)].groupby("Unit", observed=True)["Quantity"].sum() if "Days to BBD" in rf.columns else pd.Series(dtype="int")
        safe_bbd = rf[rf["Days to BBD"] > 90].groupby("Unit", observed=True)["Quantity"].sum() if "Days to BBD" in rf.columns else pd.Series(dtype="int")
        leakage = rf[rf["Warehouse"].str.endswith("_101", na=False)].groupby("Unit", observed=True)["Quantity"].sum()

        breakdown_df = pd.DataFrame({
            "UNIT": total_qty.index,
            "TOTAL QTY": total_qty.values,
            "ITEMS": items.reindex(total_qty.index).values,
            "CRITICAL BBD": pd.to_numeric(critical_bbd.reindex(total_qty.index).fillna(0), downcast='integer').values,
            "WARNING BBD": pd.to_numeric(warning_bbd.reindex(total_qty.index).fillna(0), downcast='integer').values,
            "SAFE BBD": pd.to_numeric(safe_bbd.reindex(total_qty.index).fillna(0), downcast='integer').values,
            "LEAKAGE WH": pd.to_numeric(leakage.reindex(total_qty.index).fillna(0), downcast='integer').values,
        })
        # Optional columns
        if "Days to SBD" in rf.columns:
            critical_sbd = rf[rf["Days to SBD"] < 30].groupby("Unit", observed=True)["Quantity"].sum()
            breakdown_df["CRITICAL SBD"] = pd.to_numeric(critical_sbd.reindex(total_qty.index).fillna(0), downcast='integer').values
        if "Consumed inventory" in rf.columns:
            consumed = groups["Consumed inventory"].sum()
            breakdown_df["CONSUMED INV"] = pd.to_numeric(consumed.reindex(total_qty.index).fillna(0), downcast='integer').values

        st.subheader("Plant Level Breakdown")
        st.dataframe(breakdown_df, hide_index=True, use_container_width=True)

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(rf, hide_index=True, use_container_width=True, height=500)
        lazy_export_section(rf, "Risk_Overview.xlsx", "risk")

# ═════════════════════════════════════════════════════════════════════==
# TAB 3 — Distributor Stock Overview
# ═════════════════════════════════════════════════════════════════════==
with tab3:
    st.header("Distributor Stock Overview")

    col1,col2,col3,col4,col5,col6,col7 = st.columns(7)
    with col1: sel_dist    = st.multiselect("District",      sorted(distributor["District"].dropna().unique()),  key="dist_district")
    with col2:
        d_opts = distributor.loc[distributor["District"].isin(sel_dist), "Distributor"] if sel_dist else distributor["Distributor"]
        sel_db = st.multiselect("Distributor", sorted(d_opts.dropna().unique()), key="dist_distributor")
    with col3: sel_sku     = st.multiselect("SKU",           sorted(distributor["SKU"].dropna().unique()),       key="dist_sku")
    with col4: sel_brand   = st.multiselect("Brand",         sorted(distributor["Brand"].dropna().unique()),     key="dist_brand")
    with col5: sel_pack    = st.multiselect("Pack Size",     sorted(distributor["Pack Size"].dropna().unique()), key="dist_pack")
    with col6: sel_expiry  = st.multiselect("Expiry Status", ["Critical","Warning","Safe"],                      key="dist_exp")
    with col7: search_dist = st.text_input("Search", key="dist_search")

    dmask = pd.Series(True, index=distributor.index)
    if sel_dist:   dmask &= distributor["District"].isin(sel_dist)
    if sel_db:     dmask &= distributor["Distributor"].isin(sel_db)
    if sel_sku:    dmask &= distributor["SKU"].isin(sel_sku)
    if sel_brand:  dmask &= distributor["Brand"].isin(sel_brand)
    if sel_pack:   dmask &= distributor["Pack Size"].isin(sel_pack)
    if sel_expiry: dmask &= distributor["BBD Status"].isin(sel_expiry)
    df = distributor[dmask]
    if search_dist:
        text_cols = df.select_dtypes(include=["object","category"]).columns.tolist()
        if text_cols:
            search_series = df[text_cols].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            df = df[search_series.str.contains(search_dist.lower(), na=False)]
        else:
            df = df.head(0)

    if df.empty:
        st.warning("No data available for the selected filters.")
    else:
        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
        c1.metric("Total Qty",             f"{int(df['Quantity'].sum()):,}")
        c2.metric("Total Distributors",    int(df["Distributor"].nunique()))
        c3.metric("Unique SKUs",           int(df["SKU"].nunique()))
        c4.metric("Districts",             int(df["District"].nunique()))
        c5.metric("Critical BBD (<30d)",   int(df[df["Days to BBD"] < 30]["Quantity"].sum()))
        c6.metric("Warning BBD (31-90d)",  int(df[(df["Days to BBD"] >= 31) & (df["Days to BBD"] <= 90)]["Quantity"].sum()))
        c7.metric("Safe BBD (>90d)",       int(df[df["Days to BBD"] > 90]["Quantity"].sum()))

        st.divider()
        stock_district = df.groupby("District", observed=True)["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
        stock_district_text = pd.to_numeric(stock_district["Quantity"].round(), downcast='integer')
        fig_dist = px.bar(stock_district, x="District", y="Quantity", text=stock_district_text, title="Stock by District")
        fig_dist = _style_bar(fig_dist)
        st.plotly_chart(fig_dist, use_container_width=True)

        colA, colB = st.columns(2)
        with colA:
            sb = df.groupby("Brand", observed=True)["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            sb_text = pd.to_numeric(sb["Quantity"].round(), downcast='integer')
            fig_b = px.bar(sb, x="Brand", y="Quantity", text=sb_text, title="Stock by Brand")
            fig_b = _style_bar(fig_b)
            st.plotly_chart(fig_b, use_container_width=True)
        with colB:
            sp = df.groupby("Pack Size", observed=True)["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            sp_text = pd.to_numeric(sp["Quantity"].round(), downcast='integer')
            fig_p = px.bar(sp, x="Pack Size", y="Quantity", text=sp_text, title="Stock by Pack Size")
            fig_p = _style_bar(fig_p)
            st.plotly_chart(fig_p, use_container_width=True)

        st.divider()
        st.subheader("District Level Breakdown")
        # District-Distributor breakdown using groupby to avoid nested loops
        grp = df.groupby(["District","Distributor"], observed=True)
        total_qty = grp["Quantity"].sum()
        items = grp["SKU"].nunique()
        critical = df[df["Days to BBD"] < 30].groupby(["District","Distributor"], observed=True)["Quantity"].sum() if "Days to BBD" in df.columns else pd.Series(dtype="int")
        warning = df[(df["Days to BBD"] >= 31) & (df["Days to BBD"] <= 90)].groupby(["District","Distributor"], observed=True)["Quantity"].sum() if "Days to BBD" in df.columns else pd.Series(dtype="int")
        safe = df[df["Days to BBD"] > 90].groupby(["District","Distributor"], observed=True)["Quantity"].sum() if "Days to BBD" in df.columns else pd.Series(dtype="int")

        breakdown_df = total_qty.reset_index().rename(columns={"Quantity":"Total Qty"})
        breakdown_df["Items"] = items.values
        breakdown_df["Critical BBD"] = pd.to_numeric(critical.reindex(total_qty.index).fillna(0), downcast='integer').values
        breakdown_df["Warning BBD"] = pd.to_numeric(warning.reindex(total_qty.index).fillna(0), downcast='integer').values
        breakdown_df["Safe BBD"] = pd.to_numeric(safe.reindex(total_qty.index).fillna(0), downcast='integer').values

        st.dataframe(breakdown_df, hide_index=True, use_container_width=True)

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(df, hide_index=True, use_container_width=True, height=500)
        lazy_export_section(df, "Distributor_Overview.xlsx", "dist")

# ═════════════════════════════════════════════════════════════════════==
# TAB 4 — Secondary Sales Overview
# ═════════════════════════════════════════════════════════════════════==
with tab4:
    st.header("Secondary Sales Overview")

    qty_col    = "Qty"       if "Qty"       in secondary.columns else "QTY"
    outlet_col = "Outlet Code" if "Outlet Code" in secondary.columns else "Outlet"
    brand_col  = "Brands"    if "Brands"    in secondary.columns else "Brand"

    # Row 1: 8 cascading dimension filters (mask-based, zero copies)
    cascade_cols = [
        ("District",    "sec_district"),
        ("SM",          "sec_sm"),
        ("ASM",         "sec_asm"),
        ("Route",       "sec_route"),
        ("Distributor", "sec_distributor"),
        (brand_col,     "sec_brand"),
        ("Category",    "sec_category"),
        ("Pack Size",   "sec_pack"),
    ]
    cols = st.columns(8)
    mask = pd.Series(True, index=secondary.index)
    for (col_name, widget_key), col in zip(cascade_cols, cols):
        with col:
            # Cache the options for the current mask to avoid repeated full-table scans
            idx_tuple_for_opts = tuple(secondary.loc[mask].index)
            options = cached_unique_options(idx_tuple_for_opts, 'secondary', col_name)
            selected = st.multiselect(col_name, options, key=widget_key)
        if selected:
            mask &= secondary[col_name].isin(selected)

    # Month filter (month name only) and Search on same line. Month width equals one filter; search takes remaining space.
    date_candidates = [c for c in secondary.columns if any(k in c.lower() for k in ("date","invoice","trans","doc","created"))]
    month_selected = None
    if date_candidates:
        date_col = date_candidates[0]
        # Build month options lazily from the current mask to respect preceding filters
        try:
            months_ser = pd.to_datetime(secondary.loc[mask, date_col], errors='coerce').dt.month_name()
            month_options_raw = months_ser.dropna().unique()
            # order months Jan..Dec
            month_order = list(calendar.month_name)[1:]
            month_options = [m for m in month_order if m in month_options_raw]
        except Exception:
            month_options = []
    else:
        month_options = []

    # Layout: one small column for Month (same width as other filters) and a large column for Search
    col_month, col_search = st.columns([1,7])
    with col_month:
        if month_options:
            month_selected = st.multiselect("Month", month_options, key="sec_month")
            if month_selected and date_candidates:
                # Apply month filter to mask (month name only)
                month_ser_full = pd.to_datetime(secondary[date_col], errors='coerce').dt.month_name()
                mask &= month_ser_full.isin(month_selected)
    with col_search:
        search_sec = st.text_input("Search", key="sec_search")

    # Apply mask — avoid an unnecessary reset_index copy
    df = secondary.loc[mask]
    if search_sec:
        text_cols = df.select_dtypes(include=["object","category"]).columns.tolist()
        if text_cols:
            search_series = df[text_cols].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            df = df[search_series.str.contains(search_sec.lower(), na=False)]
        else:
            df = df.head(0)

    # KPIs
    total_qty = df[qty_col].sum() if qty_col in df.columns else 0
    total_rev = df["NetRevenue"].sum() if "NetRevenue" in df.columns else 0
    nsr       = round(total_rev / total_qty) if total_qty > 0 else 0
    avg_ips   = round(df.groupby(outlet_col, observed=True)["IPS"].nunique().mean()) if ("IPS" in df.columns and not df.empty) else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Outlets",                  int(df[outlet_col].nunique()) if outlet_col in df.columns else 0)
    c2.metric("Total Volume (QTY)",             f"{int(total_qty):,}")
    c3.metric("Total Revenue",                  f"{int(total_rev):,}")
    c4.metric("NSR (Revenue per Unit)",         f"{int(nsr):,}")
    c5.metric("Avg IPS (Items per Store)",      int(avg_ips))

    st.divider()

    colA, colB = st.columns(2)
    # compute index tuple once for all groupby cache lookups
    idx_tuple = tuple(df.index)
    with colA:
        ap = cached_groupby_sum(idx_tuple, 'secondary', 'ASM', qty_col)
        ap = ap.sort_values(qty_col, ascending=False)
        ap[qty_col] = pd.to_numeric(ap[qty_col].round(), downcast='integer')
        fig_a = px.bar(ap, x="ASM", y=qty_col, text=qty_col, title="ASM Performance (Volume)")
        fig_a = _style_bar(fig_a)
        st.plotly_chart(fig_a, use_container_width=True)
    with colB:
        bp = cached_groupby_sum(idx_tuple, 'secondary', brand_col, qty_col)
        bp = bp.sort_values(qty_col, ascending=False)
        bp[qty_col] = pd.to_numeric(bp[qty_col].round(), downcast='integer')
        fig_b = px.bar(bp, x=brand_col, y=qty_col, text=qty_col, title="Brand Performance")
        fig_b = _style_bar(fig_b)
        st.plotly_chart(fig_b, use_container_width=True)

    st.divider()
    colC, colD = st.columns(2)
    with colC:
        cp = cached_groupby_sum(idx_tuple, 'secondary', 'Category', qty_col)
        cp = cp.sort_values(qty_col, ascending=False)
        cp[qty_col] = pd.to_numeric(cp[qty_col].round(), downcast='integer')
        fig_c = px.bar(cp, x="Category", y=qty_col, text=qty_col, title="Category Performance")
        fig_c = _style_bar(fig_c)
        st.plotly_chart(fig_c, use_container_width=True)
    with colD:
        pp = cached_groupby_sum(idx_tuple, 'secondary', 'Pack Size', qty_col)
        pp = pp.sort_values(qty_col, ascending=False)
        pp[qty_col] = pd.to_numeric(pp[qty_col].round(), downcast='integer')
        fig_p = px.bar(pp, x="Pack Size", y=qty_col, text=qty_col, title="Pack Size Performance")
        fig_p = _style_bar(fig_p)
        st.plotly_chart(fig_p, use_container_width=True)

    st.divider()
    colE, colF = st.columns(2)
    with colE:
        vp = cached_groupby_sum(idx_tuple, 'secondary', 'VPO', qty_col)
        vp[qty_col] = pd.to_numeric(vp[qty_col].round(), downcast='integer')
        fig_v = px.pie(vp, names="VPO", values=qty_col, title="VPO Contribution")
        fig_v = _style_pie(fig_v)
        st.plotly_chart(fig_v, use_container_width=True)
    with colF:
        chp = cached_groupby_sum(idx_tuple, 'secondary', 'CustomerHierarchy', qty_col)
        chp[qty_col] = pd.to_numeric(chp[qty_col].round(), downcast='integer')
        fig_ch = px.pie(chp, names="CustomerHierarchy", values=qty_col, title="CustomerHierarchy Contribution")
        fig_ch = _style_pie(fig_ch)
        st.plotly_chart(fig_ch, use_container_width=True)

    st.divider()
    st.subheader("Detail Table")
    st.dataframe(df.head(1000), hide_index=True, use_container_width=True, height=500)
    lazy_export_section(df, "Secondary_Sales_Overview.xlsx", "sec")
