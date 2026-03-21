import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import defaultdict
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Flipkart Reconciliation – Ashirwad Garments",
    layout="wide",
    page_icon="🧾",
)

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.3rem; font-weight: 700; }
.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

st.title("🧾 Flipkart Reconciliation Tool")
st.caption("Ashirwad Garments — auto-calculate Flipkart charges & reconcile orders")

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📂 Upload Files")
    order_file   = st.file_uploader("1️⃣  Order CSV  (Flipkart export)", type=["csv"])
    charges_file = st.file_uploader("2️⃣  Data Excel (3-sheet workbook)", type=["xlsx"])
    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)",  value=5,     min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)",       value=18,    min_value=0, step=1) / 100
    tds_rate  = st.number_input("TDS rate (%)",             value=0.095, min_value=0.0, step=0.001, format="%.3f") / 100
    tcs_rate  = st.number_input("TCS rate (%)",             value=0.477, min_value=0.0, step=0.001, format="%.3f") / 100
    st.caption("TDS & TCS are applied on **Selling Price** (Invoice − GT Charge)")
    st.markdown("---")
    st.markdown("""
**Excel sheet layout:**
- **Sheet 1** – Charges Description
- **Sheet 2** – Category Description (SKU → sub-category)
- **Sheet 3** – Price We Need (SKU → PWN)
""")

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════
CAT_MAP = {
    "kurta":       "Kurta",
    "top":         "Top",
    "blouse":      "Blouse",
    "trouser":     "Pant",
    "dress":       "Dresses",
    "night_suit":  "Nightsuit Sets",
    "nightsuit":   "Nightsuit Sets",
    "shirt":       "Men's shirt",
    "kaftan":      "Kaftan",
    "ethnic_set":  "Co-ords Set",
}

SIZE_EXPAND = {
    "L-XL":       ["L",   "XL"],
    "S-M":        ["S",   "M"],
    "XXL-3XL":    ["XXL", "3XL"],
    "F-S/XXL":    ["F"],
    "F-3xl/5xl":  ["F"],
    "XS-S":       ["XS",  "S"],
    "M-L":        ["M",   "L"],
    "XL-XXL":     ["XL",  "XXL"],
    "3XL-4XL":    ["3XL", "4XL"],
    "5XL-6XL":    ["5XL", "6XL"],
    "7XL-8XL":    ["7XL", "8XL"],
    "4XL-5XL":    ["4XL", "5XL"],
    "2XL-3XL":    ["2XL", "3XL"],
    "XS-S-M":     ["XS",  "S",  "M"],
    "L-XL-XXL":   ["L",   "XL", "XXL"],
}

ORDER_TO_PRICE_SIZE: dict = defaultdict(list)
for _ps, _os_list in SIZE_EXPAND.items():
    for _os in _os_list:
        ORDER_TO_PRICE_SIZE[_os.upper()].append(_ps)

# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_cat(raw: str) -> str:
    if pd.isna(raw):
        return ""
    return CAT_MAP.get(str(raw).strip().lower(), str(raw).strip().title())


def lookup_pwn(sku: str, pwn_dict: dict) -> tuple:
    """Returns (pwn_value, match_method)."""
    key = sku.strip().upper()
    val = pwn_dict.get(key)
    if pd.notna(val):
        return float(val), "direct"
    parts = key.rsplit("-", 1)
    if len(parts) == 2:
        base, size = parts
        for price_size in ORDER_TO_PRICE_SIZE.get(size, []):
            val = pwn_dict.get(f"{base}-{price_size.upper()}")
            if pd.notna(val):
                return float(val), price_size
    return np.nan, "not_found"


def get_gt_charge(cat_norm: str, inv: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo, hi, gtv = r.get("GT Lower Lim."), r.get("GT Upper Lim."), r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gtv):
            if float(lo) <= inv <= float(hi):
                return float(gtv)
    return 0.0


def get_commission(cat_norm: str, sell: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo, hi, ch = r.get("Lower Lim."), r.get("Upper Lim."), r.get("Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell <= float(hi):
                return float(ch) * sell
    return 0.0


def get_collection_fee(cat_norm: str, sell: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo_raw = r.get("Coll.Lower Lim.")
        hi     = r.get("Coll. Upper Lim.")
        cf     = r.get("Coll.Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).startswith(">")) else float(lo_raw)
        if lo_val < sell <= float(hi):
            return float(cf) * sell
    return 0.0


def parse_charges_df(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df["Category"].notna()].copy()
    df["Category"] = df["Category"].ffill()
    for col in ["Lower Lim.", "Upper Lim.", "Charge",
                "Coll.Lower Lim.", "Coll. Upper Lim.", "Coll.Charge",
                "GT Lower Lim.", "GT Upper Lim.", "GT Charge"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def parse_sku_cat(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(df.iloc[:, 0].astype(str).str.strip().str.upper(),
                    df.iloc[:, 1].astype(str).str.strip()))


def parse_pwn_dict(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"]    = pd.to_numeric(df["PWN+10%+50"], errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(), df["PWN+10%+50"]))


# ═══════════════════════════════════════════════════════════════════════════════
# CORE RECONCILIATION
# ═══════════════════════════════════════════════════════════════════════════════

def run_reconciliation(order_df, charges_df, sku_cat_dict, pwn_dict,
                       fixed_fee, gst_rate, tds_rate, tcs_rate,
                       pwn_overrides: dict = None) -> pd.DataFrame:

    pwn_overrides = pwn_overrides or {}
    rows_out = []

    for _, row in order_df.iterrows():
        sku        = str(row.get("SKU", "")).strip()
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)

        # Category
        sub_cat_raw  = sku_cat_dict.get(sku.upper(), "")
        sub_cat_norm = normalize_cat(sub_cat_raw)

        # ── Step 1 : GT Charge → Selling Price ────────────────────────
        gt_charge  = get_gt_charge(sub_cat_norm, inv_amount, charges_df)
        sell_price = round(inv_amount - gt_charge, 2)

        # ── Step 2 : Commission & Collection (on Selling Price) ────────
        commission = get_commission(sub_cat_norm, sell_price, charges_df)
        coll_fee   = get_collection_fee(sub_cat_norm, sell_price, charges_df)

        # ── Step 3 : Total Charges & GST ──────────────────────────────
        total_charges  = commission + coll_fee + float(fixed_fee)
        gst_on_charges = round(total_charges * gst_rate, 5)

        # ── Step 4 : TDS & TCS (on Selling Price) ─────────────────────
        tds_amount = round(sell_price * tds_rate, 4)
        tcs_amount = round(sell_price * tcs_rate, 4)

        # ── Step 5 : Received Amount ───────────────────────────────────
        # Selling Price − Total Charges − GST − TDS − TCS
        received_amount = round(
            sell_price - total_charges - gst_on_charges - tds_amount - tcs_amount, 2
        )

        # ── Step 6 : PWN lookup ────────────────────────────────────────
        pwn_val, match_method = lookup_pwn(sku, pwn_dict)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = pwn_overrides[sku.upper()], "manual"

        # ── Step 7 : Difference ────────────────────────────────────────
        difference = round(received_amount - pwn_val, 2) if pd.notna(pwn_val) else np.nan

        rows_out.append({
            "Order Id":             order_id,
            "SKU":                  sku,
            "Ordered On":           ordered_on,
            "Category":             sub_cat_raw,
            "Quantity":             quantity,
            "Invoice Amount (₹)":   inv_amount,
            "GT Charge (₹)":        gt_charge,
            "Selling Price (₹)":    sell_price,
            "Commission (₹)":       round(commission, 4),
            "Collection Fee (₹)":   round(coll_fee, 4),
            "Fixed Fee (₹)":        float(fixed_fee),
            "Total Charges (₹)":    round(total_charges, 4),
            "GST on Charges (₹)":   gst_on_charges,
            "TDS (₹)":              tds_amount,
            "TCS (₹)":              tcs_amount,
            "Total Deductions (₹)": round(total_charges + gst_on_charges + tds_amount + tcs_amount, 4),
            "Received Amount (₹)":  received_amount,
            "PWN (₹)":              pwn_val,
            "PWN Match":            match_method,
            "Difference (₹)":       difference,
        })

    return pd.DataFrame(rows_out)


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def to_excel_multi(recon_df: pd.DataFrame, summary_df: pd.DataFrame,
                   charges_breakdown: pd.DataFrame) -> bytes:
    """Export 3-sheet Excel: Reconciliation, Charges Summary, Category Breakdown."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for df, sheet in [
            (recon_df,          "Reconciliation"),
            (charges_breakdown, "Charges Breakdown"),
            (summary_df,        "Charges Summary"),
        ]:
            df.to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            for col_cells in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 38)
    return buf.getvalue()


def build_charges_summary(result_df: pd.DataFrame) -> tuple:
    """
    Returns (summary_totals_df, category_breakdown_df).
    """
    # ── Grand total row ───────────────────────────────────────────────
    num_cols = [
        "Invoice Amount (₹)", "GT Charge (₹)", "Selling Price (₹)",
        "Commission (₹)", "Collection Fee (₹)", "Fixed Fee (₹)",
        "Total Charges (₹)", "GST on Charges (₹)",
        "TDS (₹)", "TCS (₹)", "Total Deductions (₹)",
        "Received Amount (₹)", "Difference (₹)",
    ]
    totals = {c: result_df[c].sum() for c in num_cols if c in result_df.columns}
    totals["Orders"] = len(result_df)
    totals["Avg Received per Order"] = round(result_df["Received Amount (₹)"].mean(), 2)
    totals["Avg Difference per Order"] = round(result_df["Difference (₹)"].mean(), 2)
    totals["Orders with -ve Diff"] = int((result_df["Difference (₹)"] < 0).sum())
    totals["Orders with +ve Diff"] = int((result_df["Difference (₹)"] > 0).sum())
    totals["Orders with no PWN"]   = int(result_df["Difference (₹)"].isna().sum())

    summary_df = pd.DataFrame([
        {"Metric": k, "Value": round(v, 2) if isinstance(v, float) else v}
        for k, v in totals.items()
    ])

    # ── Category breakdown ────────────────────────────────────────────
    cat_df = (
        result_df.groupby("Category")
        .agg(
            Orders           = ("Order Id",             "count"),
            Invoice_Total    = ("Invoice Amount (₹)",    "sum"),
            GT_Total         = ("GT Charge (₹)",         "sum"),
            Selling_Total    = ("Selling Price (₹)",      "sum"),
            Commission_Total = ("Commission (₹)",         "sum"),
            Collection_Total = ("Collection Fee (₹)",     "sum"),
            FixedFee_Total   = ("Fixed Fee (₹)",          "sum"),
            TotalCharges     = ("Total Charges (₹)",      "sum"),
            GST_Total        = ("GST on Charges (₹)",     "sum"),
            TDS_Total        = ("TDS (₹)",                "sum"),
            TCS_Total        = ("TCS (₹)",                "sum"),
            Deductions_Total = ("Total Deductions (₹)",   "sum"),
            Received_Total   = ("Received Amount (₹)",    "sum"),
            NetDiff          = ("Difference (₹)",         "sum"),
            AvgDiff          = ("Difference (₹)",         "mean"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
        .round(2)
    )
    cat_df.columns = [
        "Category", "Orders", "Invoice Total", "GT Total", "Selling Total",
        "Commission Total", "Collection Total", "Fixed Fee Total",
        "Total Charges", "GST Total", "TDS Total", "TCS Total",
        "Total Deductions", "Received Total", "Net Difference", "Avg Difference",
    ]
    return summary_df, cat_df


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k, v in [("pwn_overrides", {}), ("result_df", None),
             ("charges_df", None), ("sku_cat_dict", None),
             ("pwn_dict", None), ("order_df", None)]:
    if k not in st.session_state:
        st.session_state[k] = v


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
if order_file and charges_file:

    with st.spinner("🔄 Reading files…"):
        order_df = pd.read_csv(order_file)
        xl       = pd.read_excel(charges_file, sheet_name=None, header=None)
        sheets   = list(xl.values())
        if len(sheets) < 3:
            st.error("❌ Excel workbook must have at least 3 sheets.")
            st.stop()
        charges_df   = parse_charges_df(sheets[0])
        sku_cat_dict = parse_sku_cat(sheets[1])
        pwn_dict     = parse_pwn_dict(sheets[2])
        st.session_state.update({
            "charges_df": charges_df, "sku_cat_dict": sku_cat_dict,
            "pwn_dict": pwn_dict, "order_df": order_df,
        })

    result_df = run_reconciliation(
        order_df, charges_df, sku_cat_dict, pwn_dict,
        fixed_fee, gst_rate, tds_rate, tcs_rate,
        pwn_overrides=st.session_state["pwn_overrides"],
    )
    st.session_state["result_df"] = result_df

    summary_df, cat_breakdown_df = build_charges_summary(result_df)

    st.success(f"✅ Processed **{len(result_df):,}** orders")

    # ── Tabs ─────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "📋  Reconciliation",
        "💰  Charges Summary",
        "📊  Category Breakdown",
    ])

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 1 – RECONCILIATION                                         ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab1:

        # ── PWN Not-Found editor ─────────────────────────────────────
        missing_df = result_df[result_df["PWN Match"] == "not_found"]
        if len(missing_df):
            with st.expander(
                f"⚠️  **{len(missing_df)} order(s) — PWN price not found** — click to enter manually",
                expanded=True,
            ):
                st.markdown("Enter the correct PWN for each SKU, then click **Save & Recalculate**.")
                missing_skus    = missing_df["SKU"].unique().tolist()
                override_inputs = {}
                n_cols = min(len(missing_skus), 4)
                cols   = st.columns(n_cols)
                for i, sku in enumerate(missing_skus):
                    existing = st.session_state["pwn_overrides"].get(sku.upper(), 0.0)
                    override_inputs[sku] = cols[i % n_cols].number_input(
                        f"PWN – **{sku}**",
                        value=float(existing), min_value=0.0, step=0.5,
                        key=f"pwn_input_{sku}",
                    )
                if st.button("💾 Save PWN Overrides & Recalculate", type="primary"):
                    for sku, val in override_inputs.items():
                        if val > 0:
                            st.session_state["pwn_overrides"][sku.upper()] = val
                    result_df = run_reconciliation(
                        order_df, charges_df, sku_cat_dict, pwn_dict,
                        fixed_fee, gst_rate, tds_rate, tcs_rate,
                        pwn_overrides=st.session_state["pwn_overrides"],
                    )
                    st.session_state["result_df"] = result_df
                    st.success("✅ Overrides saved — reconciliation updated!")
                    st.rerun()

        # ── KPI Cards ────────────────────────────────────────────────
        st.markdown("### 📊 Summary")
        k1,k2,k3,k4,k5,k6,k7,k8 = st.columns(8)
        k1.metric("Orders",            f"{len(result_df):,}")
        k2.metric("Invoice Total",     f"₹{result_df['Invoice Amount (₹)'].sum():,.0f}")
        k3.metric("Selling Pr. Total", f"₹{result_df['Selling Price (₹)'].sum():,.0f}")
        k4.metric("Total Charges",     f"₹{result_df['Total Charges (₹)'].sum():,.0f}")
        k5.metric("GST Total",         f"₹{result_df['GST on Charges (₹)'].sum():,.0f}")
        k6.metric("TDS + TCS Total",   f"₹{result_df['TDS (₹)'].sum() + result_df['TCS (₹)'].sum():,.1f}")
        k7.metric("Received Total",    f"₹{result_df['Received Amount (₹)'].sum():,.0f}")
        net = result_df['Difference (₹)'].sum()
        k8.metric(
            "Net Difference",
            f"₹{net:,.2f}",
            delta=f"{'▲' if net>=0 else '▼'} {abs(net):,.2f}",
            delta_color="normal" if net >= 0 else "inverse",
        )

        st.markdown("---")

        # ── Filters ──────────────────────────────────────────────────
        f1, f2, f3 = st.columns([2, 2, 3])
        all_cats = ["All"] + sorted(result_df["Category"].dropna().unique().tolist())
        sel_cat  = f1.selectbox("Category", all_cats)
        diff_opt = f2.selectbox("Difference type",
                                ["All", "Positive (+)", "Negative (−)",
                                 "Zero / Matched", "No PWN data"])
        search   = f3.text_input("🔎 Search by SKU or Order ID")

        view = result_df.copy()
        if sel_cat != "All":
            view = view[view["Category"] == sel_cat]
        if diff_opt == "Positive (+)":
            view = view[view["Difference (₹)"] > 0]
        elif diff_opt == "Negative (−)":
            view = view[view["Difference (₹)"] < 0]
        elif diff_opt == "Zero / Matched":
            view = view[view["Difference (₹)"] == 0]
        elif diff_opt == "No PWN data":
            view = view[view["Difference (₹)"].isna()]
        if search.strip():
            mask = (
                view["SKU"].str.contains(search.strip(), case=False, na=False) |
                view["Order Id"].str.contains(search.strip(), case=False, na=False)
            )
            view = view[mask]

        st.caption(f"Showing **{len(view):,}** of **{len(result_df):,}** orders")

        display_cols = [
            "Order Id", "SKU", "Ordered On", "Category",
            "Invoice Amount (₹)", "GT Charge (₹)", "Selling Price (₹)",
            "Commission (₹)", "Collection Fee (₹)", "Fixed Fee (₹)",
            "Total Charges (₹)", "GST on Charges (₹)",
            "TDS (₹)", "TCS (₹)", "Total Deductions (₹)",
            "Received Amount (₹)", "PWN (₹)", "Difference (₹)", "PWN Match",
        ]

        money_cols = [c for c in display_cols if "₹" in c]

        def colour_diff(val):
            try:
                v = float(val)
                if v < 0: return "color:red;font-weight:bold"
                if v > 0: return "color:green;font-weight:bold"
            except Exception:
                pass
            return ""

        fmt_inr = lambda x: f"₹{x:.2f}" if pd.notna(x) else "—"

        styled = (
            view[display_cols]
            .style
            .applymap(colour_diff, subset=["Difference (₹)"])
            .format({c: fmt_inr for c in money_cols})
        )
        st.dataframe(styled, use_container_width=True, height=500)

        # ── Downloads ────────────────────────────────────────────────
        st.markdown("### 📥 Download")
        d1, d2, d3 = st.columns(3)
        d1.download_button(
            "⬇  Full Reconciliation (Excel)",
            data=to_excel_multi(result_df, summary_df, cat_breakdown_df),
            file_name="flipkart_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        d2.download_button(
            "⬇  Filtered View (Excel)",
            data=to_excel_multi(view[display_cols].reset_index(drop=True),
                                summary_df, cat_breakdown_df),
            file_name="flipkart_reconciliation_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 2 – CHARGES SUMMARY                                        ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab2:
        st.markdown("### 💰 Total Charges Summary")
        st.caption("Grand total of every charge component across all orders")

        # ── Big metrics ──────────────────────────────────────────────
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("#### 📤 What Flipkart Deducts")
            r1, r2 = st.columns(2)
            r1.metric("GT Charges Total",       f"₹{result_df['GT Charge (₹)'].sum():,.2f}")
            r2.metric("Commission Total",        f"₹{result_df['Commission (₹)'].sum():,.2f}")
            r1.metric("Collection Fee Total",    f"₹{result_df['Collection Fee (₹)'].sum():,.2f}")
            r2.metric("Fixed Fee Total",         f"₹{result_df['Fixed Fee (₹)'].sum():,.2f}")
            r1.metric("GST on Charges Total",    f"₹{result_df['GST on Charges (₹)'].sum():,.2f}")
            r2.metric("TDS Total",               f"₹{result_df['TDS (₹)'].sum():,.2f}")
            r1.metric("TCS Total",               f"₹{result_df['TCS (₹)'].sum():,.2f}")
            r2.metric("Total Deductions",
                      f"₹{result_df['Total Deductions (₹)'].sum():,.2f}",
                      delta="All deductions combined")

        with col_b:
            st.markdown("#### 📥 What You Receive")
            r3, r4 = st.columns(2)
            r3.metric("Total Invoice Amount",   f"₹{result_df['Invoice Amount (₹)'].sum():,.2f}")
            r4.metric("Total Selling Price",    f"₹{result_df['Selling Price (₹)'].sum():,.2f}")
            r3.metric("Total Received Amount",  f"₹{result_df['Received Amount (₹)'].sum():,.2f}")
            net = result_df['Difference (₹)'].sum()
            r4.metric(
                "Net Difference vs PWN",
                f"₹{net:,.2f}",
                delta=f"{'▲' if net>=0 else '▼'} {abs(net):,.2f}",
                delta_color="normal" if net >= 0 else "inverse",
            )
            r3.metric("Orders with -ve Diff",
                      int((result_df["Difference (₹)"] < 0).sum()))
            r4.metric("Orders with +ve Diff",
                      int((result_df["Difference (₹)"] > 0).sum()))

        st.markdown("---")

        # ── Detailed charges table ────────────────────────────────────
        st.markdown("#### 📋 Per-Order Charges Detail")
        charges_view_cols = [
            "Order Id", "SKU", "Category",
            "Invoice Amount (₹)", "GT Charge (₹)", "Selling Price (₹)",
            "Commission (₹)", "Collection Fee (₹)", "Fixed Fee (₹)",
            "Total Charges (₹)", "GST on Charges (₹)",
            "TDS (₹)", "TCS (₹)", "Total Deductions (₹)",
            "Received Amount (₹)",
        ]
        charges_money = [c for c in charges_view_cols if "₹" in c]
        charges_styled = (
            result_df[charges_view_cols]
            .style
            .format({c: fmt_inr for c in charges_money})
            .background_gradient(subset=["Total Deductions (₹)"], cmap="Reds")
        )
        st.dataframe(charges_styled, use_container_width=True, height=480)

        # ── Charges summary card (all as one table) ───────────────────
        st.markdown("---")
        st.markdown("#### 🧾 Charges Summary Table")
        st.dataframe(summary_df, use_container_width=True)

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 3 – CATEGORY BREAKDOWN                                     ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab3:
        st.markdown("### 📊 Category-wise Breakdown")
        st.caption("Every charge component summed per category")

        # colour net diff
        def colour_net(val):
            try:
                v = float(val)
                if v < 0: return "color:red;font-weight:bold"
                if v > 0: return "color:green;font-weight:bold"
            except Exception:
                pass
            return ""

        cat_money_cols = [c for c in cat_breakdown_df.columns if c not in ("Category", "Orders")]
        cat_styled = (
            cat_breakdown_df
            .style
            .applymap(colour_net, subset=["Net Difference", "Avg Difference"])
            .format({c: "₹{:.2f}" for c in cat_money_cols})
        )
        st.dataframe(cat_styled, use_container_width=True)

        st.markdown("---")
        st.markdown("#### 🔢 Charge Components by Category (stacked view)")
        charge_comp_df = cat_breakdown_df[[
            "Category", "Orders",
            "GT Total", "Commission Total", "Collection Total",
            "Fixed Fee Total", "GST Total", "TDS Total", "TCS Total",
            "Total Deductions",
        ]].copy()
        st.dataframe(
            charge_comp_df.style.format(
                {c: "₹{:.2f}" for c in charge_comp_df.columns if c not in ("Category","Orders")}
            ).background_gradient(subset=["Total Deductions"], cmap="YlOrRd"),
            use_container_width=True,
        )

# ═══════════════════════════════════════════════════════════════════════════════
# LANDING SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
else:
    st.info("👈 Upload **both files** in the sidebar to begin.")
    st.markdown("""
---
### How it works

| Step | File | Used columns |
|------|------|-------------|
| 1 | **Order CSV** | `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Quantity` |
| 2 | **Data Excel** | 3-sheet workbook: Charges / Category / Price We Need |

---
### Calculation per order

```
GT Charge       = fixed slab  (based on Invoice Amount)
Selling Price   = Invoice Amount − GT Charge

Commission      = Selling Price × Commission %   (slab by Selling Price)
Collection Fee  = Selling Price × Collection %   (slab by Selling Price)
Total Charges   = Commission + Collection Fee + Fixed Fee

GST Amount      = Total Charges × 18 %
TDS             = Selling Price × 0.095 %
TCS             = Selling Price × 0.477 %

Received Amount = Selling Price − Total Charges − GST − TDS − TCS

Difference      = Received Amount − PWN
```

**Three tabs:**
- **Reconciliation** — order-level detail with filters & download
- **Charges Summary** — grand totals for every charge component
- **Category Breakdown** — all charges grouped by product category
""")
