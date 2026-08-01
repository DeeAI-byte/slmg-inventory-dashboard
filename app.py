"""
SLMG Inventory Hub — optimized for Streamlit Community Cloud
Python 3.11 | pandas 2.2.3 | pyarrow 16 | streamlit 1.45.1
"""
import gc
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="SLMG Inventory Hub", layout="wide")


def make_bar(data: pd.DataFrame, x: str, y: str, title: str):
    fig = px.bar(data, x=x, y=y, text=y, title=title)
    fig.update_traces(texttemplate="%{text:,}", textposition="outside")
    fig.update_yaxes(tickformat="d")
    fig.update_layout(margin=dict(t=40, b=10))
    return fig


def make_pie(data: pd.DataFrame, names: str, values: str, title: str):
    fig = px.pie(data, names=names, values=values, title=title)
    fig.update_traces(textinfo="percent+label", texttemplate="%{label}<br>%{percent:.1%}")
    fig.update_layout(margin=dict(t=40, b=10))
    return fig


def export_section(df: pd.DataFrame, file_name: str, key: str) -> None:
    if st.button("Prepare Export", key=f"{key}_prep"):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        st.download_button(
            "⬇ Download Excel",
            buf.getvalue(),
            file_name=file_name,
            key=f"{key}_dl",
        )
        del buf
        gc.collect()


def _optimize(df: pd.DataFrame) -> pd.DataFrame:
    n = len(df)
    if n == 0:
        return df

    for col in df.columns:
        try:
            dtype = df[col].dtype
            if dtype == object:
                if df[col].nunique() / n < 0.5:
                    df[col] = df[col].astype("category")
            elif dtype == "int64":
                df[col] = pd.to_numeric(df[col], downcast="integer")
            elif dtype == "float64":
                df[col] = pd.to_numeric(df[col], downcast="float")
        except Exception:
            pass
    return df


def _clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def _display_options(df: pd.DataFrame, column: str) -> list[str]:
    if df.empty or column not in df.columns:
        return []
    return sorted({str(value) for value in df[column].dropna().unique()})


def _selected_mask(series: pd.Series, selections: list[str]) -> pd.Series:
    if not selections:
        return pd.Series(True, index=series.index)

    lookup = {str(value): value for value in series.dropna().unique()}
    values = [lookup[value] for value in selections if value in lookup]
    return series.isin(values)


def _search_mask(df: pd.DataFrame, columns: list[str], query: str) -> pd.Series:
    """Search one column at a time so no full-table text column is retained."""
    mask = pd.Series(False, index=df.index)
    needle = query.lower()
    for column in columns:
        if column in df.columns:
            mask |= (df[column].astype("string").str.lower()
                     .str.contains(needle, na=False))
    return mask


def _to_int32(series: pd.Series) -> pd.Series:
    return (
        pd.to_numeric(series, errors="coerce")
        .replace([float("inf"), float("-inf")], 0)
        .fillna(0)
        .round()
        .clip(-2_147_483_648, 2_147_483_647)
        .astype("int32")
    )


def _require_columns(df: pd.DataFrame, columns: list[str], dashboard: str) -> bool:
    missing = [column for column in columns if column not in df.columns]
    if missing:
        st.error(f"{dashboard} cannot be rendered; missing columns: {', '.join(missing)}")
        return False
    return True


@st.cache_resource(show_spinner="Loading data…")
def load_data():

    def safe_read(path: str) -> pd.DataFrame:
        try:
            if path.endswith(".parquet"):
                df = pd.read_parquet(path)
            elif path.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path)
            return _clean_column_names(df)
        except Exception as exc:
            raise RuntimeError(f"Unable to load {path}") from exc

    master = safe_read("Master Stock.xlsx")
    risk = safe_read("Near Expiry iNVENTORY_.xlsx")
    distributor = safe_read("DBR.xlsx")

    try:
        import duckdb as _ddb

        _con = _ddb.connect("secondary.duckdb", read_only=True)
        _arrow = _con.execute("SELECT * FROM secondary").fetch_arrow_table()
        secondary = _arrow.to_pandas(strings_to_categorical=True)
        secondary = _clean_column_names(secondary)
        _con.close()
        del _con, _ddb, _arrow
        gc.collect()
    except Exception as exc:
        raise RuntimeError("Unable to load secondary.duckdb table 'secondary'") from exc

    if not master.empty:
        for c in ["MFG Date", "EXP Date", "DOD Date"]:
            if c in master.columns:
                master[c] = pd.to_datetime(master[c], errors="coerce", dayfirst=True)

        if "Quantity" in master.columns:
            master["Quantity"] = _to_int32(master["Quantity"])

        if "Shelflife" in master.columns:
            sl = (
                (pd.to_numeric(master["Shelflife"], errors="coerce") * 100)
                .round()
                .replace([float("inf"), float("-inf")], 0)
                .fillna(0)
                .clip(-32_768, 32_767)
            )
            master["Shelflife"] = sl.astype("int16")
            master["SL Status"] = pd.cut(
                sl,
                bins=[-1, 30, 90, 100000],
                labels=["Critical", "Warning", "Safe"],
                right=False,
            ).astype("category")

        scols = [
            c for c in ["SKU", "Brand", "Category", "Pack Size", "Site", "Warehouse"]
            if c in master.columns
        ]
        master["_search"] = (
            master[scols].astype("string").fillna("").agg(" ".join, axis=1).str.lower()
        )
        master = _optimize(master)

    if not risk.empty:
        for c in ["MFG Date", "EXP Date", "BBD/Expiry", "DOD Date"]:
            if c in risk.columns:
                risk[c] = pd.to_datetime(risk[c], errors="coerce", dayfirst=True)

        if "Days to BBD" in risk.columns:
            bbd = (
                pd.to_numeric(risk["Days to BBD"], errors="coerce")
                .replace([float("inf"), float("-inf")], 0)
                .fillna(0)
            )
            risk["Days to BBD"] = _to_int32(bbd)
            risk["BBD Status"] = pd.cut(
                bbd,
                bins=[-99999, 30, 90, 99999],
                labels=["Critical", "Warning", "Safe"],
                right=False,
            ).astype("category")

        for c in risk.select_dtypes(include="number").columns:
            risk[c] = _to_int32(risk[c])

        scols = [c for c in ["SKU", "Unit", "Warehouse", "Batch"] if c in risk.columns]
        risk["_search"] = (
            risk[scols].astype("string").fillna("").agg(" ".join, axis=1).str.lower()
        )
        risk = _optimize(risk)

    if not distributor.empty:
        if "BBD/Expiry" in distributor.columns and "EXP Date" not in distributor.columns:
            distributor.rename(columns={"BBD/Expiry": "EXP Date"}, inplace=True)

        for c in ["MFG Date", "EXP Date"]:
            if c in distributor.columns:
                distributor[c] = pd.to_datetime(
                    distributor[c], errors="coerce", dayfirst=True
                )

        if "EXP Date" in distributor.columns:
            distributor["EXP Date"] = distributor["EXP Date"].astype(str).str.strip()
            bbd = (
                pd.to_datetime(distributor["EXP Date"], errors="coerce", dayfirst=True)
                - pd.Timestamp("today")
            ).dt.days.fillna(0)
            distributor["Days to BBD"] = _to_int32(bbd)
            distributor["BBD Status"] = pd.cut(
                bbd,
                bins=[-99999, 30, 90, 99999],
                labels=["Critical", "Warning", "Safe"],
                right=False,
            ).astype("category")

        if "Quantity" in distributor.columns:
            distributor["Quantity"] = _to_int32(distributor["Quantity"])

        scols = [
            c for c in ["SKU", "Brand", "Category", "Pack Size", "District", "Distributor"]
            if c in distributor.columns
        ]
        distributor["_search"] = (
            distributor[scols].astype("string").fillna("").agg(" ".join, axis=1).str.lower()
        )
        distributor = _optimize(distributor)

    if not secondary.empty:
        if "Qty" in secondary.columns and "QTY" not in secondary.columns:
            secondary.rename(columns={"Qty": "QTY"}, inplace=True)
        if "Brands" in secondary.columns and "Brand" not in secondary.columns:
            secondary.rename(columns={"Brands": "Brand"}, inplace=True)
        if "Outlet Code" in secondary.columns and "Outlet" not in secondary.columns:
            secondary.rename(columns={"Outlet Code": "Outlet"}, inplace=True)

        for c in ["QTY", "NetRevenue"]:
            if c not in secondary.columns:
                secondary[c] = 0
            secondary[c] = _to_int32(secondary[c])

        for c in [
            "District", "SM", "ASM", "STL", "Route", "Distributor", "Brand",
            "Category", "Pack Size", "VPO", "Channel", "SubChannel", "Month",
            "CustomerHierarchy", "IPS",
        ]:
            if c in secondary.columns and not isinstance(secondary[c].dtype, pd.CategoricalDtype):
                secondary[c] = secondary[c].astype("category")

    filter_opts = {
        "m": {
            c: _display_options(master, c)
            for c in ["Site", "Warehouse", "SKU", "Brand", "Category", "Pack Size"]
        },
        "r": {
            c: _display_options(risk, c)
            for c in ["Unit", "Warehouse", "SKU"]
        },
        "d": {
            c: _display_options(distributor, c)
            for c in ["District", "Distributor", "SKU", "Brand", "Pack Size"]
        },
        "s": {
            c: _display_options(secondary, c)
            for c in [
                "District", "SM", "ASM", "Route", "Distributor",
                "Brand", "Category", "Pack Size", "Month",
            ]
        },
    }

    gc.collect()
    return master, risk, distributor, secondary, filter_opts


master, risk, distributor, secondary, OPTS = load_data()

st.markdown("<style>.block-container{padding-top:1.5rem;}</style>", unsafe_allow_html=True)
st.markdown(
    "<h2 style='text-align:center;font-family:Georgia;font-size:32px;'>"
    "Coca‑Cola | SLMG Beverages</h2>",
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs([
    "Stock Overview",
    "Risk Stock Overview",
    "Distributor Stock Overview",
    "Secondary Sales Overview",
])


@st.fragment
def render_stock():
    st.header("Stock Overview")

    if not _require_columns(
        master,
        [
            "Site", "Warehouse", "SKU", "Brand", "Category", "Pack Size",
            "Quantity", "Shelflife", "SL Status", "_search",
        ],
        "Stock Overview",
    ):
        return

    c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
    with c1:
        sel_sites = st.multiselect("Site", OPTS["m"]["Site"], key="t1_site")
    with c2:
        wh = (
            _display_options(
                master.loc[_selected_mask(master["Site"], sel_sites)],
                "Warehouse",
            )
            if sel_sites else OPTS["m"]["Warehouse"]
        )
        sel_wh = st.multiselect("Warehouse", wh, key="t1_wh")
    with c3:
        sel_skus = st.multiselect("SKU", OPTS["m"]["SKU"], key="t1_sku")
    with c4:
        sel_brands = st.multiselect("Brand", OPTS["m"]["Brand"], key="t1_brand")
    with c5:
        sel_cats = st.multiselect("Category", OPTS["m"]["Category"], key="t1_cat")
    with c6:
        sel_pack = st.multiselect("Pack Size", OPTS["m"]["Pack Size"], key="t1_pack")
    with c7:
        sel_sl = st.multiselect(
            "Shelf Life",
            ["Critical", "Warning", "Safe"],
            key="t1_sl",
        )
    with c8:
        search = st.text_input("Search", key="t1_search")

    mask = pd.Series(True, index=master.index)
    if sel_sites:
        mask &= _selected_mask(master["Site"], sel_sites)
    if sel_wh:
        mask &= _selected_mask(master["Warehouse"], sel_wh)
    if sel_skus:
        mask &= _selected_mask(master["SKU"], sel_skus)
    if sel_brands:
        mask &= _selected_mask(master["Brand"], sel_brands)
    if sel_cats:
        mask &= _selected_mask(master["Category"], sel_cats)
    if sel_pack:
        mask &= _selected_mask(master["Pack Size"], sel_pack)
    if sel_sl:
        mask &= master["SL Status"].isin(sel_sl)
    if search:
        mask &= master["_search"].str.contains(search.lower(), na=False)

    f = master.loc[mask]

    qty = int(f["Quantity"].sum())
    sl_s = f["Shelflife"]
    q_s = f["Quantity"]
    crit = int(q_s[sl_s < 30].sum())
    warn = int(q_s[(sl_s >= 31) & (sl_s <= 90)].sum())
    safe = int(q_s[sl_s >= 90].sum())

    kc = st.columns(7)
    kc[0].metric("Total Stock", f"{qty:,}")
    kc[1].metric("Unique SKUs", int(f["SKU"].nunique()))
    kc[2].metric("Sites", int(f["Site"].nunique()))
    kc[3].metric("Warehouses", int(f["Warehouse"].nunique()))
    kc[4].metric("Critical SL (<30%)", f"{crit:,}")
    kc[5].metric("Warning SL (31-90%)", f"{warn:,}")
    kc[6].metric("Safe SL (>=90%)", f"{safe:,}")

    site_sum = (
        f.groupby("Site", observed=True)["Quantity"]
        .sum()
        .reset_index()
        .sort_values("Quantity", ascending=False)
    )
    st.plotly_chart(
        make_bar(site_sum, "Site", "Quantity", "Stock By Site"),
        use_container_width=True,
    )

    lc, rc = st.columns(2)
    with lc:
        st.subheader("Top 5 SKUs By Inventory")
        top5 = (
            f.groupby("SKU", observed=True)["Quantity"]
            .sum()
            .nlargest(5)
            .reset_index()
        )
        st.dataframe(top5, hide_index=True, use_container_width=True)

    with rc:
        st.subheader("Top 5 SKUs — Lowest Shelf Life %")
        bot5 = f[["SKU", "Shelflife", "Quantity"]].nsmallest(5, "Shelflife")
        st.dataframe(bot5, hide_index=True, use_container_width=True)

    st.subheader("Detail Table")
    disp = f.drop(columns=["_search"], errors="ignore")
    st.dataframe(disp.head(1000), hide_index=True, use_container_width=True, height=500)
    export_section(disp, "Stock_Overview.xlsx", "stock")


with tab1:
    render_stock()


@st.fragment
def render_risk():
    st.header("Risk Stock Overview")

    if not _require_columns(
        risk,
        [
            "Unit", "Warehouse", "SKU", "Quantity", "Days to BBD",
            "BBD Status", "_search",
        ],
        "Risk Stock Overview",
    ):
        return

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        sel_sites = st.multiselect("Site", OPTS["r"]["Unit"], key="t2_site")
    with c2:
        wh = (
            _display_options(
                risk.loc[_selected_mask(risk["Unit"], sel_sites)],
                "Warehouse",
            )
            if sel_sites else OPTS["r"]["Warehouse"]
        )
        sel_wh = st.multiselect("Warehouse", wh, key="t2_wh")
    with c3:
        sel_sku = st.multiselect("SKU", OPTS["r"]["SKU"], key="t2_sku")
    with c4:
        sel_expiry = st.multiselect(
            "Expiry Status",
            ["Critical", "Warning", "Safe"],
            key="t2_exp",
        )
    with c5:
        search = st.text_input("Search", key="t2_search")

    mask = pd.Series(True, index=risk.index)
    if sel_sites:
        mask &= _selected_mask(risk["Unit"], sel_sites)
    if sel_wh:
        mask &= _selected_mask(risk["Warehouse"], sel_wh)
    if sel_sku:
        mask &= _selected_mask(risk["SKU"], sel_sku)
    if sel_expiry:
        mask &= risk["BBD Status"].isin(sel_expiry)
    if search:
        mask &= risk["_search"].str.contains(search.lower(), na=False)

    rf = risk.loc[mask]

    if rf.empty:
        st.warning("No data available for the selected filters.")
        return

    bbd = rf["Days to BBD"]
    qty = rf["Quantity"]

    kc = st.columns(6)
    kc[0].metric("Total Qty", f"{int(qty.sum()):,}")
    kc[1].metric("Critical BBD (<30d)", f"{int(qty[bbd < 30].sum()):,}")
    kc[2].metric(
        "Warning BBD (31-90d)",
        f"{int(qty[(bbd >= 31) & (bbd <= 90)].sum()):,}",
    )
    kc[3].metric("Safe BBD (>90d)", f"{int(qty[bbd > 90].sum()):,}")
    kc[4].metric(
        "Critical SBD (<30d)",
        f"{int(qty[rf['Days to SBD'] < 30].sum()):,}"
        if "Days to SBD" in rf.columns else "N/A",
    )
    kc[5].metric(
        "Critical LBD (<30d)",
        f"{int(qty[rf['Days to LBD'] < 30].sum()):,}"
        if "Days to LBD" in rf.columns else "N/A",
    )

    st.divider()
    lc, rc = st.columns(2)
    with lc:
        st.subheader("Top 5 At-Risk BBD SKUs")
        st.dataframe(
            rf[["SKU", "Quantity", "Days to BBD"]].nsmallest(5, "Days to BBD"),
            hide_index=True,
            use_container_width=True,
        )
    with rc:
        if "Days to SBD" in rf.columns:
            st.subheader("Top 5 At-Risk SBD SKUs")
            st.dataframe(
                rf[["SKU", "Quantity", "Days to SBD"]].nsmallest(5, "Days to SBD"),
                hide_index=True,
                use_container_width=True,
            )

    st.divider()
    rf2 = rf.assign(
        crit_qty=rf["Quantity"].where(bbd < 30, 0),
        warn_qty=rf["Quantity"].where((bbd >= 31) & (bbd <= 90), 0),
        safe_qty=rf["Quantity"].where(bbd > 90, 0),
        leak_qty=rf["Quantity"].where(
            rf["Warehouse"].astype(str).str.endswith("_101"),
            0,
        ),
    )

    if "Days to SBD" in rf2.columns:
        rf2["crit_sbd_qty"] = rf2["Quantity"].where(rf2["Days to SBD"] < 30, 0)

    agg = (
        rf2.groupby("Unit", observed=True)
        .agg(
            **{
                "TOTAL QTY": ("Quantity", "sum"),
                "ITEMS": ("SKU", "nunique"),
                "CRITICAL BBD": ("crit_qty", "sum"),
                "WARNING BBD": ("warn_qty", "sum"),
                "SAFE BBD": ("safe_qty", "sum"),
                "LEAKAGE WH": ("leak_qty", "sum"),
                **(
                    {"CRITICAL SBD": ("crit_sbd_qty", "sum")}
                    if "crit_sbd_qty" in rf2.columns else {}
                ),
            }
        )
        .reset_index()
        .rename(columns={"Unit": "UNIT"})
    )

    if "Consumed inventory" in rf2.columns:
        agg["CONSUMED INV"] = (
            rf2.groupby("Unit", observed=True)["Consumed inventory"].sum().values
        )

    st.subheader("Plant Level Breakdown")
    st.dataframe(agg, hide_index=True, use_container_width=True)

    st.divider()
    st.subheader("Detail Table")
    disp = rf.drop(columns=["_search"], errors="ignore")
    st.dataframe(disp.head(1000), hide_index=True, use_container_width=True, height=500)
    export_section(disp, "Risk_Overview.xlsx", "risk")


with tab2:
    render_risk()


@st.fragment
def render_distributor():
    st.header("Distributor Stock Overview")

    if not _require_columns(
        distributor,
        [
            "District", "Distributor", "SKU", "Brand", "Pack Size", "Quantity",
            "Days to BBD", "BBD Status", "_search",
        ],
        "Distributor Stock Overview",
    ):
        return

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1:
        sel_dist = st.multiselect("District", OPTS["d"]["District"], key="t3_dist")
    with c2:
        db = (
            _display_options(
                distributor.loc[
                    _selected_mask(distributor["District"], sel_dist)
                ],
                "Distributor",
            )
            if sel_dist else OPTS["d"]["Distributor"]
        )
        sel_db = st.multiselect("Distributor", db, key="t3_db")
    with c3:
        sel_sku = st.multiselect("SKU", OPTS["d"]["SKU"], key="t3_sku")
    with c4:
        sel_brand = st.multiselect("Brand", OPTS["d"]["Brand"], key="t3_brand")
    with c5:
        sel_pack = st.multiselect("Pack Size", OPTS["d"]["Pack Size"], key="t3_pack")
    with c6:
        sel_expiry = st.multiselect(
            "Expiry Status",
            ["Critical", "Warning", "Safe"],
            key="t3_exp",
        )
    with c7:
        search = st.text_input("Search", key="t3_search")

    mask = pd.Series(True, index=distributor.index)
    if sel_dist:
        mask &= _selected_mask(distributor["District"], sel_dist)
    if sel_db:
        mask &= _selected_mask(distributor["Distributor"], sel_db)
    if sel_sku:
        mask &= _selected_mask(distributor["SKU"], sel_sku)
    if sel_brand:
        mask &= _selected_mask(distributor["Brand"], sel_brand)
    if sel_pack:
        mask &= _selected_mask(distributor["Pack Size"], sel_pack)
    if sel_expiry:
        mask &= distributor["BBD Status"].isin(sel_expiry)
    if search:
        mask &= distributor["_search"].str.contains(search.lower(), na=False)

    df = distributor.loc[mask]

    if df.empty:
        st.warning("No data available for the selected filters.")
        return

    bbd = df["Days to BBD"]
    qty = df["Quantity"]

    kc = st.columns(7)
    kc[0].metric("Total Qty", f"{int(qty.sum()):,}")
    kc[1].metric("Total Distributors", int(df["Distributor"].nunique()))
    kc[2].metric("Unique SKUs", int(df["SKU"].nunique()))
    kc[3].metric("Districts", int(df["District"].nunique()))
    kc[4].metric("Critical BBD (<30d)", f"{int(qty[bbd < 30].sum()):,}")
    kc[5].metric(
        "Warning BBD (31-90d)",
        f"{int(qty[(bbd >= 31) & (bbd <= 90)].sum()):,}",
    )
    kc[6].metric("Safe BBD (>90d)", f"{int(qty[bbd > 90].sum()):,}")

    st.divider()
    dist_sum = (
        df.groupby("District", observed=True)["Quantity"]
        .sum()
        .reset_index()
        .sort_values("Quantity", ascending=False)
    )
    dist_sum["Quantity"] = dist_sum["Quantity"].astype(int)
    st.plotly_chart(
        make_bar(dist_sum, "District", "Quantity", "Stock by District"),
        use_container_width=True,
    )

    ca, cb = st.columns(2)
    with ca:
        brand_s = (
            df.groupby("Brand", observed=True)["Quantity"]
            .sum()
            .reset_index()
            .sort_values("Quantity", ascending=False)
        )
        brand_s["Quantity"] = brand_s["Quantity"].astype(int)
        st.plotly_chart(
            make_bar(brand_s, "Brand", "Quantity", "Stock by Brand"),
            use_container_width=True,
        )

    with cb:
        pack_s = (
            df.groupby("Pack Size", observed=True)["Quantity"]
            .sum()
            .reset_index()
            .sort_values("Quantity", ascending=False)
        )
        pack_s["Quantity"] = pack_s["Quantity"].astype(int)
        st.plotly_chart(
            make_bar(pack_s, "Pack Size", "Quantity", "Stock by Pack Size"),
            use_container_width=True,
        )

    st.divider()
    st.subheader("District Level Breakdown")

    df2 = df.assign(
        crit_qty=qty.where(bbd < 30, 0),
        warn_qty=qty.where((bbd >= 31) & (bbd <= 90), 0),
        safe_qty=qty.where(bbd > 90, 0),
    )
    breakdown = (
        df2.groupby(["District", "Distributor"], observed=True)
        .agg(
            **{
                "Total Qty": ("Quantity", "sum"),
                "Items": ("SKU", "nunique"),
                "Critical BBD": ("crit_qty", "sum"),
                "Warning BBD": ("warn_qty", "sum"),
                "Safe BBD": ("safe_qty", "sum"),
            }
        )
        .reset_index()
    )
    st.dataframe(breakdown, hide_index=True, use_container_width=True)

    st.divider()
    st.subheader("Detail Table")
    disp = df.drop(columns=["_search"], errors="ignore")
    st.dataframe(disp.head(1000), hide_index=True, use_container_width=True, height=500)
    export_section(disp, "Distributor_Overview.xlsx", "dist")


with tab3:
    render_distributor()


@st.fragment
def render_secondary():
    st.header("Secondary Sales Overview")

    if not _require_columns(
        secondary,
        [
            "District", "SM", "ASM", "Route", "Distributor", "Brand", "Category",
            "Pack Size", "Month", "Outlet", "VPO", "CustomerHierarchy", "QTY",
            "NetRevenue",
        ],
        "Secondary Sales Overview",
    ):
        return

    cascade = [
        ("District", "s4_district"),
        ("SM", "s4_sm"),
        ("ASM", "s4_asm"),
        ("Route", "s4_route"),
        ("Distributor", "s4_dist"),
        ("Brand", "s4_brand"),
        ("Category", "s4_cat"),
        ("Pack Size", "s4_pack"),
    ]

    cols = st.columns(8)
    mask = pd.Series(True, index=secondary.index)

    for (col_name, wkey), col in zip(cascade, cols):
        with col:
            opts = _display_options(secondary.loc[mask], col_name)
            sel = st.multiselect(col_name, opts, key=wkey)
        if sel:
            mask &= _selected_mask(secondary[col_name], sel)

    m_col, s_col = st.columns([1, 3])
    with m_col:
        month_opts = _display_options(secondary.loc[mask], "Month")
        sel_month = st.multiselect("Month", month_opts, key="s4_month")
    with s_col:
        search = st.text_input("Search", key="s4_search")

    if sel_month:
        mask &= _selected_mask(secondary["Month"], sel_month)
    if search:
        mask &= _search_mask(
            secondary,
            ["District", "SM", "ASM", "Route", "Distributor", "Brand",
             "Category", "Pack Size", "Outlet", "ITEMNAME"],
            search,
        )

    df = secondary.loc[mask]

    if df.empty:
        st.info("No data for the selected filters.")
        return

    total_qty = int(df["QTY"].sum())
    total_rev = int(df["NetRevenue"].sum())
    nsr = round(total_rev / total_qty) if total_qty else 0

    avg_ips = 0
    if "IPS" in df.columns:
        ips_mean = df.groupby("Outlet", observed=True)["IPS"].nunique().mean()
        avg_ips = round(ips_mean) if pd.notna(ips_mean) else 0

    kc = st.columns(5)
    kc[0].metric("Total Outlets", int(df["Outlet"].nunique()))
    kc[1].metric("Total Volume (QTY)", f"{total_qty:,}")
    kc[2].metric("Total Revenue", f"{total_rev:,}")
    kc[3].metric("NSR (Revenue per Unit)", f"{int(nsr):,}")
    kc[4].metric("Avg IPS (Items/Store)", int(avg_ips))

    st.divider()

    def _agg(col):
        return df.groupby(col, observed=True)["QTY"].sum().reset_index()

    ca, cb = st.columns(2)
    with ca:
        asm_s = _agg("ASM").sort_values("QTY", ascending=False)
        asm_s["QTY"] = asm_s["QTY"].astype(int)
        st.plotly_chart(
            make_bar(asm_s, "ASM", "QTY", "ASM Performance (Volume)"),
            use_container_width=True,
        )

    with cb:
        br_s = _agg("Brand").sort_values("QTY", ascending=False)
        br_s["QTY"] = br_s["QTY"].astype(int)
        st.plotly_chart(
            make_bar(br_s, "Brand", "QTY", "Brand Performance"),
            use_container_width=True,
        )

    st.divider()
    cc, cd = st.columns(2)
    with cc:
        cat_s = _agg("Category").sort_values("QTY", ascending=False)
        cat_s["QTY"] = cat_s["QTY"].astype(int)
        st.plotly_chart(
            make_bar(cat_s, "Category", "QTY", "Category Performance"),
            use_container_width=True,
        )

    with cd:
        pk_s = _agg("Pack Size").sort_values("QTY", ascending=False)
        pk_s["QTY"] = pk_s["QTY"].astype(int)
        st.plotly_chart(
            make_bar(pk_s, "Pack Size", "QTY", "Pack Size Performance"),
            use_container_width=True,
        )

    st.divider()
    ce, cf = st.columns(2)
    with ce:
        vpo_s = _agg("VPO")
        vpo_s["QTY"] = vpo_s["QTY"].astype(int)
        st.plotly_chart(
            make_pie(vpo_s, "VPO", "QTY", "VPO Contribution"),
            use_container_width=True,
        )

    with cf:
        ch_s = _agg("CustomerHierarchy")
        ch_s["QTY"] = ch_s["QTY"].astype(int)
        st.plotly_chart(
            make_pie(
                ch_s,
                "CustomerHierarchy",
                "QTY",
                "CustomerHierarchy Contribution",
            ),
            use_container_width=True,
        )

    st.divider()
    st.subheader("Detail Table")
    disp = df.drop(columns=["_search"], errors="ignore")
    st.dataframe(disp.head(1000), hide_index=True, use_container_width=True, height=500)
    export_section(disp, "Secondary_Sales_Overview.xlsx", "sec")


with tab4:
    render_secondary()
