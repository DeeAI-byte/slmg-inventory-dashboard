import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

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
def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def lazy_export_section(df, file_name: str, button_key: str):
    if st.button(f"Prepare Export ({file_name})", key=f"{button_key}_prepare"):
        st.session_state[f"{button_key}_bytes"] = export_excel(df)
    cached_bytes = st.session_state.get(f"{button_key}_bytes")
    if cached_bytes is not None:
        st.download_button("Download Excel", cached_bytes, file_name=file_name, key=f"{button_key}_download")

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
                m[col] = pd.to_datetime(m[col], errors="coerce", dayfirst=True).dt.date
        m["Quantity"] = pd.to_numeric(m.get("Quantity", 0), errors="coerce").fillna(0).astype(int)
        if "Shelflife" in m.columns:
            m["Shelflife"] = (pd.to_numeric(m["Shelflife"], errors="coerce") * 100).round().astype(int)
            m["SL Status"] = "Safe"
            m.loc[m["Shelflife"] < 30, "SL Status"] = "Critical"
            m.loc[(m["Shelflife"] >= 31) & (m["Shelflife"] <= 90), "SL Status"] = "Warning"
        datasets["master"] = optimize_dtypes(m)

    # ── RISK ─────────────────────────────────────────────────────────────
    r = datasets["risk"]
    if not r.empty:
        for col in ["MFG Date", "EXP Date", "BBD/Expiry", "DOD Date"]:
            if col in r.columns:
                r[col] = pd.to_datetime(r[col], errors="coerce", dayfirst=True).dt.date
        if "Days to BBD" in r.columns:
            r["Days to BBD"] = pd.to_numeric(r["Days to BBD"], errors="coerce")
            r["BBD Status"] = "Safe"
            r.loc[r["Days to BBD"] < 30, "BBD Status"] = "Critical"
            r.loc[(r["Days to BBD"] >= 31) & (r["Days to BBD"] <= 90), "BBD Status"] = "Warning"
        for col in r.select_dtypes(include="number").columns:
            r[col] = pd.to_numeric(r[col], errors="coerce").fillna(0).round().astype(int)
        datasets["risk"] = optimize_dtypes(r)

    # ── DISTRIBUTOR ───────────────────────────────────────────────────────
    d = datasets["distributor"]
    if not d.empty:
        for col in ["MFG Date", "BBD/Expiry"]:
            if col in d.columns:
                d[col] = pd.to_datetime(d[col], errors="coerce", dayfirst=True).dt.date
        if "EXP Date" in d.columns:
            d["EXP Date"] = d["EXP Date"].astype(str).str.strip()
            d["Days to BBD"] = (
                pd.to_datetime(d["EXP Date"], errors="coerce", dayfirst=True) - pd.to_datetime("today")
            ).dt.days
            d["BBD Status"] = "Safe"
            d.loc[d["Days to BBD"] < 30, "BBD Status"] = "Critical"
            d.loc[(d["Days to BBD"] >= 31) & (d["Days to BBD"] <= 90), "BBD Status"] = "Warning"
        datasets["distributor"] = optimize_dtypes(d)

    # ── SECONDARY (parquet — no optimize_dtypes, already efficient) ────────
    sec = datasets["secondary"]
    if not sec.empty:
        qty_key = "Qty" if "Qty" in sec.columns else "QTY"
        if qty_key in sec.columns:
            sec[qty_key] = pd.to_numeric(sec[qty_key], errors="coerce").fillna(0).astype(int)
        if "NetRevenue" in sec.columns:
            sec["NetRevenue"] = pd.to_numeric(sec["NetRevenue"], errors="coerce").fillna(0).astype(int)
        datasets["secondary"] = sec

    return datasets

datasets    = load_data()
master      = datasets["master"]
risk        = datasets["risk"]
distributor = datasets["distributor"]
secondary   = datasets["secondary"]

# -----------------------------
# UI
# -----------------------------
st.markdown("<style>.block-container{padding-top:1.5rem;}</style>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align:center;font-family:Georgia;font-size:32px;'>Coca‑Cola | SLMG Beverages</h2>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["Stock Overview", "Risk Stock Overview", "Distributor Stock Overview", "Secondary Sales Overview"])

# ═══════════════════════════════════════════════════════════════════════
# TAB 1 — Stock Overview
# ═══════════════════════════════════════════════════════════════════════
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
        sm = pd.Series(False, index=filtered.index)
        for c in filtered.select_dtypes(include=["object","category"]).columns:
            sm |= filtered[c].astype(str).str.contains(search_text, case=False, na=False)
        filtered = filtered[sm]

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Total Stock",        f"{int(filtered['Quantity'].sum()):,}")
    c2.metric("Unique SKUs",        int(filtered["SKU"].nunique()))
    c3.metric("Sites",              int(filtered["Site"].nunique()))
    c4.metric("Warehouses",         int(filtered["Warehouse"].nunique()))
    c5.metric("Critical SL (<30%)", int(filtered[filtered["Shelflife"] < 30]["Quantity"].sum()))
    c6.metric("Warning SL (31-90%)",int(filtered[(filtered["Shelflife"] >= 31) & (filtered["Shelflife"] <= 90)]["Quantity"].sum()))
    c7.metric("Safe SL (>=90%)",    int(filtered[filtered["Shelflife"] >= 90]["Quantity"].sum()))

    stock_site = filtered.groupby("Site")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
    fig = px.bar(stock_site, x="Site", y="Quantity", text="Quantity", title="Stock By Site")
    fig.update_traces(texttemplate='%{text:,}', textposition='outside')
    fig.update_yaxes(tickformat="d")
    st.plotly_chart(fig, width="stretch")

    left, right = st.columns(2)
    with left:
        st.subheader("Top 5 SKUs By Inventory")
        st.dataframe(filtered.groupby("SKU")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False).head(5), hide_index=True, width="stretch")
    with right:
        st.subheader("Top 5 SKUs — Lowest Shelf Life %")
        st.dataframe(filtered[["SKU","Shelflife","Quantity"]].sort_values("Shelflife").head(5), hide_index=True, width="stretch")

    st.subheader("Detail Table")
    st.dataframe(filtered, hide_index=True, width="stretch", height=500)
    lazy_export_section(filtered, "Stock_Overview.xlsx", "stock")

# ═══════════════════════════════════════════════════════════════════════
# TAB 2 — Risk Stock Overview
# ═══════════════════════════════════════════════════════════════════════
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
        sm = pd.Series(False, index=rf.index)
        for c in rf.select_dtypes(include=["object","category"]).columns:
            sm |= rf[c].astype(str).str.contains(search_risk, case=False, na=False)
        rf = rf[sm]

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
            st.dataframe(rf[["SKU","Quantity","Days to BBD"]].sort_values("Days to BBD").head(5), hide_index=True, width="stretch")
        with right:
            st.subheader("Top 5 At-Risk SBD SKUs")
            if "Days to SBD" in rf.columns:
                st.dataframe(rf[["SKU","Quantity","Days to SBD"]].sort_values("Days to SBD").head(5), hide_index=True, width="stretch")

        st.divider()
        breakdown = []
        for unit in sorted(rf["Unit"].dropna().unique()):
            temp = rf[rf["Unit"] == unit]
            leakage = temp[temp["Warehouse"].astype(str).str.endswith("_101")]["Quantity"].sum()
            row = {"UNIT": unit, "TOTAL QTY": int(temp["Quantity"].sum()), "ITEMS": int(temp["SKU"].nunique()),
                   "CRITICAL BBD": int(temp[temp["Days to BBD"] < 30]["Quantity"].sum()),
                   "WARNING BBD":  int(temp[(temp["Days to BBD"] >= 31) & (temp["Days to BBD"] <= 90)]["Quantity"].sum()),
                   "SAFE BBD":     int(temp[temp["Days to BBD"] > 90]["Quantity"].sum()),
                   "LEAKAGE WH":   int(leakage)}
            if "Days to SBD" in temp.columns:
                row["CRITICAL SBD"] = int(temp[temp["Days to SBD"] < 30]["Quantity"].sum())
            if "Consumed inventory" in temp.columns:
                row["CONSUMED INV"] = int(temp["Consumed inventory"].sum())
            breakdown.append(row)
        st.subheader("Plant Level Breakdown")
        st.dataframe(pd.DataFrame(breakdown), hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(rf, hide_index=True, width="stretch", height=500)
        lazy_export_section(rf, "Risk_Overview.xlsx", "risk")

# ═══════════════════════════════════════════════════════════════════════
# TAB 3 — Distributor Stock Overview
# ═══════════════════════════════════════════════════════════════════════
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
        sm = pd.Series(False, index=df.index)
        for c in df.select_dtypes(include=["object","category"]).columns:
            sm |= df[c].astype(str).str.contains(search_dist, case=False, na=False)
        df = df[sm]

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
        stock_district = df.groupby("District")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
        fig_dist = px.bar(stock_district, x="District", y="Quantity", text=stock_district["Quantity"].astype(int), title="Stock by District")
        fig_dist.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_dist.update_yaxes(tickformat="d")
        st.plotly_chart(fig_dist, width="stretch")

        colA, colB = st.columns(2)
        with colA:
            sb = df.groupby("Brand")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_b = px.bar(sb, x="Brand", y="Quantity", text=sb["Quantity"].astype(int), title="Stock by Brand")
            fig_b.update_traces(texttemplate='%{text:,}', textposition='outside')
            fig_b.update_yaxes(tickformat="d")
            st.plotly_chart(fig_b, width="stretch")
        with colB:
            sp = df.groupby("Pack Size")["Quantity"].sum().reset_index().sort_values("Quantity", ascending=False)
            fig_p = px.bar(sp, x="Pack Size", y="Quantity", text=sp["Quantity"].astype(int), title="Stock by Pack Size")
            fig_p.update_traces(texttemplate='%{text:,}', textposition='outside')
            fig_p.update_yaxes(tickformat="d")
            st.plotly_chart(fig_p, width="stretch")

        st.divider()
        st.subheader("District Level Breakdown")
        breakdown = []
        for district in sorted(df["District"].dropna().unique()):
            temp = df[df["District"] == district]
            for dist_name in sorted(temp["Distributor"].dropna().unique()):
                sub = temp[temp["Distributor"] == dist_name]
                breakdown.append({"District": district, "Distributor": dist_name,
                    "Total Qty":   int(sub["Quantity"].sum()),
                    "Items":       int(sub["SKU"].nunique()),
                    "Critical BBD":int(sub[sub["Days to BBD"] < 30]["Quantity"].sum()),
                    "Warning BBD": int(sub[(sub["Days to BBD"] >= 31) & (sub["Days to BBD"] <= 90)]["Quantity"].sum()),
                    "Safe BBD":    int(sub[sub["Days to BBD"] > 90]["Quantity"].sum())})
        st.dataframe(pd.DataFrame(breakdown), hide_index=True, width="stretch")

        st.divider()
        st.subheader("Detail Table")
        st.dataframe(df, hide_index=True, width="stretch", height=500)
        lazy_export_section(df, "Distributor_Overview.xlsx", "dist")

# ═══════════════════════════════════════════════════════════════════════
# TAB 4 — Secondary Sales Overview
# ═══════════════════════════════════════════════════════════════════════
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
            options = sorted(secondary.loc[mask, col_name].dropna().unique())
            selected = st.multiselect(col_name, options, key=widget_key)
        if selected:
            mask &= secondary[col_name].isin(selected)

    # Row 2: Month + Search
    f1, f2, f3 = st.columns([1, 2, 5])
    with f1:
        month_opts = sorted(secondary.loc[mask, "Month"].dropna().unique()) if "Month" in secondary.columns else []
        sel_months = st.multiselect("Month", month_opts, key="sec_month")
    with f2:
        search_sec = st.text_input("Search", key="sec_search")
    with f3:
        pass

    if sel_months:
        mask &= secondary["Month"].isin(sel_months)

    # Single filtered copy — applied once
    df = secondary[mask].reset_index(drop=True)
    if search_sec:
        sm = pd.Series(False, index=df.index)
        for c in df.select_dtypes(include=["object","category"]).columns:
            sm |= df[c].astype(str).str.contains(search_sec, case=False, na=False)
        df = df[sm]

    # KPIs
    total_qty = df[qty_col].sum() if qty_col in df.columns else 0
    total_rev = df["NetRevenue"].sum() if "NetRevenue" in df.columns else 0
    nsr       = round(total_rev / total_qty) if total_qty > 0 else 0
    avg_ips   = round(df.groupby(outlet_col)["IPS"].nunique().mean()) if ("IPS" in df.columns and not df.empty) else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Outlets",                  int(df[outlet_col].nunique()) if outlet_col in df.columns else 0)
    c2.metric("Total Volume (QTY)",             f"{int(total_qty):,}")
    c3.metric("Total Revenue",                  f"{int(total_rev):,}")
    c4.metric("NSR (Revenue per Unit)",         f"{int(nsr):,}")
    c5.metric("Avg IPS (Items per Store)",      int(avg_ips))

    st.divider()

    colA, colB = st.columns(2)
    with colA:
        ap = df.groupby("ASM")[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        ap[qty_col] = ap[qty_col].round().astype(int)
        fig_a = px.bar(ap, x="ASM", y=qty_col, text=qty_col, title="ASM Performance (Volume)")
        fig_a.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_a.update_yaxes(tickformat="d")
        st.plotly_chart(fig_a, width="stretch")
    with colB:
        bp = df.groupby(brand_col)[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        bp[qty_col] = bp[qty_col].round().astype(int)
        fig_b = px.bar(bp, x=brand_col, y=qty_col, text=qty_col, title="Brand Performance")
        fig_b.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_b.update_yaxes(tickformat="d")
        st.plotly_chart(fig_b, width="stretch")

    st.divider()
    colC, colD = st.columns(2)
    with colC:
        cp = df.groupby("Category")[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        cp[qty_col] = cp[qty_col].round().astype(int)
        fig_c = px.bar(cp, x="Category", y=qty_col, text=qty_col, title="Category Performance")
        fig_c.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_c.update_yaxes(tickformat="d")
        st.plotly_chart(fig_c, width="stretch")
    with colD:
        pp = df.groupby("Pack Size")[qty_col].sum().reset_index().sort_values(qty_col, ascending=False)
        pp[qty_col] = pp[qty_col].round().astype(int)
        fig_p = px.bar(pp, x="Pack Size", y=qty_col, text=qty_col, title="Pack Size Performance")
        fig_p.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_p.update_yaxes(tickformat="d")
        st.plotly_chart(fig_p, width="stretch")

    st.divider()
    colE, colF = st.columns(2)
    with colE:
        vp = df.groupby("VPO")[qty_col].sum().reset_index()
        vp[qty_col] = vp[qty_col].round().astype(int)
        fig_v = px.pie(vp, names="VPO", values=qty_col, title="VPO Contribution")
        fig_v.update_traces(texttemplate='%{percent:.1%}', textinfo='percent+label')
        st.plotly_chart(fig_v, width="stretch")
    with colF:
        chp = df.groupby("CustomerHierarchy")[qty_col].sum().reset_index()
        chp[qty_col] = chp[qty_col].round().astype(int)
        fig_ch = px.pie(chp, names="CustomerHierarchy", values=qty_col, title="CustomerHierarchy Contribution")
        fig_ch.update_traces(texttemplate='%{percent:.1%}', textinfo='percent+label')
        st.plotly_chart(fig_ch, width="stretch")

    st.divider()
    st.subheader("Detail Table")
    st.dataframe(df.head(1000), hide_index=True, width="stretch", height=500)
    lazy_export_section(df, "Secondary_Sales_Overview.xlsx", "sec")
