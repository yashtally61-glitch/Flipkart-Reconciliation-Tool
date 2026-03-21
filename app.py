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
table { width: 100%; }
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
    st.caption("TDS & TCS deducted on **Selling Price** (Invoice − GT Charge)")
    st.markdown("---")
    st.markdown("""
**Excel sheet layout:**
- **Sheet 1** – Charges Description
- **Sheet 2** – Category Description
- **Sheet 3** – Price We Need (PWN)
""")

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════
CAT_MAP = {
    "kurta":      "Kurta",
    "top":        "Top",
    "blouse":     "Blouse",
    "trouser":    "Pant",
    "dress":      "Dresses",
    "night_suit": "Nightsuit Sets",
    "nightsuit":  "Nightsuit Sets",
    "shirt":      "Men's shirt",
    "kaftan":     "Kaftan",
    "ethnic_set": "Co-ords Set",
}

SIZE_EXPAND = {
    "L-XL":      ["L",   "XL"],
    "S-M":       ["S",   "M"],
    "XXL-3XL":   ["XXL", "3XL"],
    "F-S/XXL":   ["F"],
    "F-3xl/5xl": ["F"],
    "XS-S":      ["XS",  "S"],
    "M-L":       ["M",   "L"],
    "XL-XXL":    ["XL",  "XXL"],
    "3XL-4XL":   ["3XL", "4XL"],
    "5XL-6XL":   ["5XL", "6XL"],
    "7XL-8XL":   ["7XL", "8XL"],
    "4XL-5XL":   ["4XL", "5XL"],
    "2XL-3XL":   ["2XL", "3XL"],
    "XS-S-M":    ["XS",  "S",  "M"],
    "L-XL-XXL":  ["L",   "XL", "XXL"],
}

ORDER_TO_PRICE_SIZE: dict = defaultdict(list)
for _ps, _os_list in SIZE_EXPAND.items():
    for _os in _os_list:
        ORDER_TO_PRICE_SIZE[_os.upper()].append(_ps)

# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_cat(raw) -> str:
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

        sub_cat_raw  = sku_cat_dict.get(sku.upper(), "")
        sub_cat_norm = normalize_cat(sub_cat_raw)

        # Step 1 – GT Charge → Selling Price
        gt_charge  = get_gt_charge(sub_cat_norm, inv_amount, charges_df)
        sell_price = round(inv_amount - gt_charge, 2)

        # Step 2 – Commission & Collection (on Selling Price)
        commission = get_commission(sub_cat_norm, sell_price, charges_df)
        coll_fee   = get_collection_fee(sub_cat_norm, sell_price, charges_df)

        # Step 3 – Total Charges & GST
        total_charges  = commission + coll_fee + float(fixed_fee)
        gst_on_charges = round(total_charges * gst_rate, 4)

        # Step 4 – TDS & TCS (on Selling Price)
        tds_amount = round(sell_price * tds_rate, 4)
        tcs_amount = round(sell_price * tcs_rate, 4)

        # Step 5 – Total Deductions = Charges + GST + TDS + TCS  ← FIX
        total_deductions = round(
            total_charges + gst_on_charges + tds_amount + tcs_amount, 4
        )

        # Step 6 – Received Amount = Selling Price − All Deductions
        received_amount = round(sell_price - total_deductions, 2)

        # Step 7 – PWN lookup (direct → size-range → manual override)
        pwn_val, match_method = lookup_pwn(sku, pwn_dict)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku.upper()]), "manual"

        # Step 8 – Difference
        difference = round(received_amount - pwn_val, 2) if pd.notna(pwn_val) else np.nan

        rows_out.append({
            "Order Id":              order_id,
            "SKU":                   sku,
            "Ordered On":            ordered_on,
            "Category":              sub_cat_raw,
            "Qty":                   quantity,
            "Invoice Amount":        inv_amount,
            "GT Charge":             gt_charge,
            "Selling Price":         sell_price,
            "Commission":            round(commission, 4),
            "Collection Fee":        round(coll_fee, 4),
            "Fixed Fee":             float(fixed_fee),
            "Total Charges":         round(total_charges, 4),
            "GST on Charges":        gst_on_charges,
            "TDS":                   tds_amount,
            "TCS":                   tcs_amount,
            "Total Deductions":      total_deductions,
            "Received Amount":       received_amount,
            "PWN":                   pwn_val,
            "PWN Match":             match_method,
            "Difference":            difference,
        })

    return pd.DataFrame(rows_out)


# ═══════════════════════════════════════════════════════════════════════════════
# FORMATTING HELPERS  (NO matplotlib)
# ═══════════════════════════════════════════════════════════════════════════════

MONEY_COLS = [
    "Invoice Amount", "GT Charge", "Selling Price",
    "Commission", "Collection Fee", "Fixed Fee",
    "Total Charges", "GST on Charges", "TDS", "TCS",
    "Total Deductions", "Received Amount", "PWN", "Difference",
]

def fmt_inr(x):
    try:
        if pd.isna(x):
            return "—"
        return f"₹{float(x):,.2f}"
    except Exception:
        return str(x)

def style_table(df: pd.DataFrame, diff_col: str = "Difference") -> object:
    """Apply ₹ formatting + red/green on diff col. No matplotlib needed."""
    fmt_dict = {c: fmt_inr for c in df.columns if c in MONEY_COLS}

    def colour_diff(val):
        try:
            v = float(val)
            if v < 0:  return "color: red; font-weight: bold"
            if v > 0:  return "color: green; font-weight: bold"
        except Exception:
            pass
        return ""

    styler = df.style.format(fmt_dict)
    if diff_col in df.columns:
        styler = styler.applymap(colour_diff, subset=[diff_col])
    return styler


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT  (3 sheets)
# ═══════════════════════════════════════════════════════════════════════════════

def to_excel(recon_df: pd.DataFrame, summary_df: pd.DataFrame,
             cat_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for df, sheet in [(recon_df, "Reconciliation"),
                          (cat_df,   "Category Breakdown"),
                          (summary_df, "Charges Summary")]:
            df.to_excel(w, index=False, sheet_name=sheet)
            ws = w.sheets[sheet]
            for col_cells in ws.columns:
                width = max(len(str(c.value or "")) for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(width + 3, 38)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def build_summary(df: pd.DataFrame) -> tuple:
    """Returns (grand_summary_df, category_breakdown_df)."""
    # Grand summary
    totals = {
        "Metric": [], "Value": [],
    }
    fields = [
        ("Total Orders",             len(df)),
        ("Total Invoice Amount",     df["Invoice Amount"].sum()),
        ("Total GT Charges",         df["GT Charge"].sum()),
        ("Total Selling Price",      df["Selling Price"].sum()),
        ("Total Commission",         df["Commission"].sum()),
        ("Total Collection Fee",     df["Collection Fee"].sum()),
        ("Total Fixed Fee",          df["Fixed Fee"].sum()),
        ("Total Charges (C+F+Fixed)",df["Total Charges"].sum()),
        ("Total GST on Charges",     df["GST on Charges"].sum()),
        ("Total TDS",                df["TDS"].sum()),
        ("Total TCS",                df["TCS"].sum()),
        ("Total Deductions",         df["Total Deductions"].sum()),
        ("Total Received Amount",    df["Received Amount"].sum()),
        ("Net Difference vs PWN",    df["Difference"].sum()),
        ("Orders with -ve Diff",     int((df["Difference"] < 0).sum())),
        ("Orders with +ve Diff",     int((df["Difference"] > 0).sum())),
        ("Orders – No PWN found",    int(df["Difference"].isna().sum())),
        ("Avg Received per Order",   df["Received Amount"].mean()),
        ("Avg Difference per Order", df["Difference"].mean()),
    ]
    for label, val in fields:
        totals["Metric"].append(label)
        totals["Value"].append(round(val, 2) if isinstance(val, float) else val)

    summary_df = pd.DataFrame(totals)

    # Category breakdown
    cat_df = (
        df.groupby("Category")
        .agg(
            Orders          = ("Order Id",          "count"),
            Invoice_Total   = ("Invoice Amount",     "sum"),
            GT_Total        = ("GT Charge",          "sum"),
            Selling_Total   = ("Selling Price",      "sum"),
            Commission      = ("Commission",         "sum"),
            Collection      = ("Collection Fee",     "sum"),
            Fixed           = ("Fixed Fee",          "sum"),
            Total_Charges   = ("Total Charges",      "sum"),
            GST_Total       = ("GST on Charges",     "sum"),
            TDS_Total       = ("TDS",                "sum"),
            TCS_Total       = ("TCS",                "sum"),
            Deductions      = ("Total Deductions",   "sum"),
            Received_Total  = ("Received Amount",    "sum"),
            Net_Diff        = ("Difference",         "sum"),
            Avg_Diff        = ("Difference",         "mean"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
        .round(2)
    )
    cat_df.columns = [
        "Category", "Orders", "Invoice Total", "GT Total", "Selling Total",
        "Commission", "Collection", "Fixed Fee",
        "Total Charges", "GST Total", "TDS Total", "TCS Total",
        "Total Deductions", "Received Total", "Net Difference", "Avg Difference",
    ]
    return summary_df, cat_df


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k, v in [
    ("pwn_overrides", {}),
    ("result_df", None),
    ("charges_df", None),
    ("sku_cat_dict", None),
    ("pwn_dict", None),
    ("order_df", None),
]:
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
            st.error("❌ Excel must have at least 3 sheets.")
            st.stop()
        charges_df   = parse_charges_df(sheets[0])
        sku_cat_dict = parse_sku_cat(sheets[1])
        pwn_dict     = parse_pwn_dict(sheets[2])
        st.session_state.update({
            "charges_df": charges_df, "sku_cat_dict": sku_cat_dict,
            "pwn_dict": pwn_dict, "order_df": order_df,
        })

    # Run
    result_df = run_reconciliation(
        order_df, charges_df, sku_cat_dict, pwn_dict,
        fixed_fee, gst_rate, tds_rate, tcs_rate,
        pwn_overrides=st.session_state["pwn_overrides"],
    )
    st.session_state["result_df"] = result_df
    summary_df, cat_df = build_summary(result_df)

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

        # ── PWN Not-Found editor (FIXED UI) ──────────────────────────
        missing_df = result_df[result_df["PWN Match"] == "not_found"]
        if len(missing_df):
            with st.expander(
                f"⚠️  **{len(missing_df)} SKU(s) have no PWN price — click to enter manually**",
                expanded=True,
            ):
                st.info(
                    "Each row below shows the full SKU name. "
                    "Enter the correct PWN value and click **Save & Recalculate**."
                )
                missing_skus = missing_df["SKU"].unique().tolist()
                override_inputs = {}

                # Table-like layout: SKU label on left, number input on right
                for sku in missing_skus:
                    c_label, c_input = st.columns([3, 2])
                    c_label.markdown(
                        f"<div style='padding-top:8px;font-size:0.95rem;"
                        f"word-break:break-all'><b>{sku}</b></div>",
                        unsafe_allow_html=True,
                    )
                    existing = float(st.session_state["pwn_overrides"].get(sku.upper(), 0.0))
                    override_inputs[sku] = c_input.number_input(
                        "PWN (₹)",
                        value=existing,
                        min_value=0.0,
                        step=0.5,
                        label_visibility="collapsed",
                        key=f"pwn_input_{sku}",
                    )

                if st.button("💾  Save PWN Overrides & Recalculate", type="primary"):
                    for sku, val in override_inputs.items():
                        if val > 0:
                            st.session_state["pwn_overrides"][sku.upper()] = val
                    result_df = run_reconciliation(
                        order_df, charges_df, sku_cat_dict, pwn_dict,
                        fixed_fee, gst_rate, tds_rate, tcs_rate,
                        pwn_overrides=st.session_state["pwn_overrides"],
                    )
                    st.session_state["result_df"] = result_df
                    summary_df, cat_df = build_summary(result_df)
                    st.success("✅ Saved and recalculated!")
                    st.rerun()

        # ── KPI strip ────────────────────────────────────────────────
        st.markdown("### 📊 Summary")
        k1,k2,k3,k4,k5,k6,k7,k8 = st.columns(8)
        k1.metric("Orders",            f"{len(result_df):,}")
        k2.metric("Invoice Total",     f"₹{result_df['Invoice Amount'].sum():,.0f}")
        k3.metric("Selling Pr. Total", f"₹{result_df['Selling Price'].sum():,.0f}")
        k4.metric("Total Charges",     f"₹{result_df['Total Charges'].sum():,.0f}")
        k5.metric("GST Total",         f"₹{result_df['GST on Charges'].sum():,.0f}")
        k6.metric("TDS + TCS",         f"₹{(result_df['TDS']+result_df['TCS']).sum():,.1f}")
        k7.metric("Received Total",    f"₹{result_df['Received Amount'].sum():,.0f}")
        net = result_df["Difference"].sum()
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
            view = view[view["Difference"] > 0]
        elif diff_opt == "Negative (−)":
            view = view[view["Difference"] < 0]
        elif diff_opt == "Zero / Matched":
            view = view[view["Difference"] == 0]
        elif diff_opt == "No PWN data":
            view = view[view["Difference"].isna()]
        if search.strip():
            mask = (
                view["SKU"].str.contains(search.strip(), case=False, na=False) |
                view["Order Id"].str.contains(search.strip(), case=False, na=False)
            )
            view = view[mask]

        st.caption(f"Showing **{len(view):,}** of **{len(result_df):,}** orders")

        display_cols = [
            "Order Id", "SKU", "Ordered On", "Category",
            "Invoice Amount", "GT Charge", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "TDS", "TCS", "Total Deductions",
            "Received Amount", "PWN", "Difference", "PWN Match",
        ]

        st.dataframe(
            style_table(view[display_cols], diff_col="Difference"),
            use_container_width=True,
            height=500,
        )

        # ── Downloads ────────────────────────────────────────────────
        st.markdown("### 📥 Download")
        d1, d2 = st.columns(2)
        d1.download_button(
            "⬇  Full Reconciliation (Excel – 3 sheets)",
            data=to_excel(result_df, summary_df, cat_df),
            file_name="flipkart_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        d2.download_button(
            "⬇  Filtered View (Excel)",
            data=to_excel(view[display_cols].reset_index(drop=True), summary_df, cat_df),
            file_name="flipkart_reconciliation_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 2 – CHARGES SUMMARY                                        ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab2:
        st.markdown("### 💰 Total Charges Summary")
        st.caption("Grand total of every deduction across all orders")

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("#### 📤 Flipkart Deductions")
            a1, a2 = st.columns(2)
            a1.metric("GT Charges",        f"₹{result_df['GT Charge'].sum():,.2f}")
            a2.metric("Commission",         f"₹{result_df['Commission'].sum():,.2f}")
            a1.metric("Collection Fee",     f"₹{result_df['Collection Fee'].sum():,.2f}")
            a2.metric("Fixed Fee",          f"₹{result_df['Fixed Fee'].sum():,.2f}")
            a1.metric("GST on Charges",     f"₹{result_df['GST on Charges'].sum():,.2f}")
            a2.metric("TDS",                f"₹{result_df['TDS'].sum():,.2f}")
            a1.metric("TCS",                f"₹{result_df['TCS'].sum():,.2f}")

            total_ded = result_df["Total Deductions"].sum()
            st.metric(
                "🔴 Total Deductions (all above)",
                f"₹{total_ded:,.2f}",
            )

        with col_b:
            st.markdown("#### 📥 What You Receive")
            b1, b2 = st.columns(2)
            b1.metric("Total Invoice",      f"₹{result_df['Invoice Amount'].sum():,.2f}")
            b2.metric("Total Selling Pr.",  f"₹{result_df['Selling Price'].sum():,.2f}")
            b1.metric("Total Received",     f"₹{result_df['Received Amount'].sum():,.2f}")
            net = result_df["Difference"].sum()
            b2.metric(
                "Net Diff vs PWN",
                f"₹{net:,.2f}",
                delta=f"{'▲' if net>=0 else '▼'} {abs(net):,.2f}",
                delta_color="normal" if net >= 0 else "inverse",
            )
            b1.metric("Orders –ve Diff",    int((result_df["Difference"] < 0).sum()))
            b2.metric("Orders +ve Diff",    int((result_df["Difference"] > 0).sum()))

        st.markdown("---")

        # Per-order charges table (NO background_gradient = no matplotlib crash)
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols = [
            "Order Id", "SKU", "Category",
            "Invoice Amount", "GT Charge", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "TDS", "TCS", "Total Deductions",
            "Received Amount",
        ]
        st.dataframe(
            style_table(result_df[charge_cols]),
            use_container_width=True,
            height=480,
        )

        st.markdown("---")
        st.markdown("#### 🧾 Grand Summary Table")
        st.dataframe(summary_df, use_container_width=True)

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 3 – CATEGORY BREAKDOWN                                     ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab3:
        st.markdown("### 📊 Category-wise Breakdown")
        st.caption("Every charge component summed per product category")

        cat_money = [c for c in cat_df.columns if c not in ("Category", "Orders")]
        st.dataframe(
            style_table(cat_df, diff_col="Net Difference")
            .format({c: "₹{:.2f}" for c in cat_money}),
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("#### 🔢 Charge Components Only (per Category)")
        comp_cols = [
            "Category", "Orders",
            "GT Total", "Commission", "Collection", "Fixed Fee",
            "GST Total", "TDS Total", "TCS Total", "Total Deductions",
        ]
        comp_money = [c for c in comp_cols if c not in ("Category", "Orders")]
        st.dataframe(
            cat_df[comp_cols].style.format(
                {c: "₹{:.2f}" for c in comp_money}
            ),
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

| File | Description |
|------|-------------|
| **Order CSV** | Flipkart Seller Hub export — uses `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Quantity` |
| **Data Excel** | 3-sheet workbook: Charges Description / Category Description / Price We Need |

---
### Calculation per order

```
GT Charge        = fixed slab  (looked up by Invoice Amount)
Selling Price    = Invoice Amount − GT Charge

Commission       = Selling Price × Commission %   (slab by Selling Price)
Collection Fee   = Selling Price × Collection %   (slab by Selling Price)
Total Charges    = Commission + Collection Fee + Fixed Fee

GST Amount       = Total Charges × 18%
TDS              = Selling Price × 0.095%
TCS              = Selling Price × 0.477%

Total Deductions = Total Charges + GST + TDS + TCS
Received Amount  = Selling Price − Total Deductions

Difference       = Received Amount − PWN
```

**PWN lookup:** tries exact SKU first, then auto-maps combined sizes  
(e.g. order SKU `7053YKBLS-L` → matches price sheet `7053YKBLS-L-XL`)

If a PWN is still missing, an editor appears to enter it manually.
""")
