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
    charges_file = st.file_uploader("2️⃣  Data Excel (5-sheet workbook)", type=["xlsx"])
    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5,  min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)",      value=18, min_value=0, step=1) / 100
    st.markdown("---")
    st.markdown("""
**Excel sheet layout (5 sheets):**
- **Sheet 1** – (ignored / any)
- **Sheet 2** – Charges Description
- **Sheet 3** – Category Description  (Seller SKU → Sub-category)
- **Sheet 4** – Price We Need (PWN)
- **Sheet 5** – (ignored / any)

**Charges sheet columns:**
`Category` | `Lower Lim.` | `Upper Lim.` | `Charge` |
`Coll.Lower Lim.` | `Coll. Upper Lim.` | `Coll.Charge` |
`GT Lower Lim.` | `GT Upper Lim.` | `GT Charge`
""")

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Sub-category names that differ from Sheet2 Category names
CAT_MAP = {
    "kurta":      "Kurta",
    "top":        "Top",
    "kaftan":     "Kaftan",
    "trouser":    "Pant",
    "shirt":      "Men's shirt",
    "mens_kurta": "Men's Kurta",
    # All others (ethnic_set, blouse, dress, night_suit, etc.)
    # match Sheet2 directly via case-insensitive lookup
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

VENDOR_PREFIXES = ["GWN-", "GWN_"]


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def strip_vendor_prefix(sku: str) -> str:
    upper = sku.upper()
    for prefix in VENDOR_PREFIXES:
        if upper.startswith(prefix.upper()):
            return sku[len(prefix):]
    return sku


def get_cat_for_lookup(sub_cat_raw: str, charges_df: pd.DataFrame) -> str:
    """
    Map raw sub-category name (from Sheet3) to the Category value used in
    Sheet2 charges table.

    Rule:
      1. Try the explicit CAT_MAP for names that differ (kurta→Kurta, trouser→Pant, etc.)
      2. Fall back to a case-insensitive direct match against Sheet2 Category column.
         e.g. 'ethnic_set', 'blouse', 'dress', 'night_suit' all match directly.
      3. Return empty string if nothing matches (order will be left as NaN).
    """
    if not sub_cat_raw or sub_cat_raw == "nan":
        return ""
    key = sub_cat_raw.strip().lower()
    if key in CAT_MAP:
        return CAT_MAP[key]
    matches = charges_df[charges_df["Category"].str.lower() == key]["Category"].unique()
    return matches[0] if len(matches) else ""


def lookup_pwn(sku: str, pwn_dict: dict) -> tuple:
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


def get_gt_amount(cat: str, inv_amount: float, cdf: pd.DataFrame) -> float:
    """
    Look up fixed GT charge from Invoice Amount slab.
    GT Amount = fixed ₹ value (not a percentage).
    """
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("GT Lower Lim.")
        hi = r.get("GT Upper Lim.")
        gt = r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gt):
            if float(lo) <= inv_amount <= float(hi) + 0.99:
                return float(gt)
    return np.nan


def get_commission(cat: str, sell: float, cdf: pd.DataFrame) -> float:
    """Commission = Selling Price × Commission rate (slab by Selling Price)."""
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("Lower Lim.")
        hi = r.get("Upper Lim.")
        ch = r.get("Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell <= float(hi) + 0.99:
                return round(float(ch) * sell, 5)
    return 0.0


def get_collection_fee(cat: str, sell: float, cdf: pd.DataFrame) -> float:
    """Collection Fee = Selling Price × Collection rate (slab by Selling Price)."""
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo_raw = r.get("Coll.Lower Lim.")
        hi     = r.get("Coll. Upper Lim.")
        cf     = r.get("Coll.Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).strip().startswith(">")) else float(lo_raw)
        if lo_val < sell <= float(hi) + 0.99:
            return round(float(cf) * sell, 5)
    return 0.0


def parse_charges_df(raw: pd.DataFrame) -> pd.DataFrame:
    """Parse Sheet2 (Charges Description)."""
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df["Category"].notna()].copy()
    df["Category"] = df["Category"].ffill()
    numeric_cols = [
        "Lower Lim.", "Upper Lim.", "Charge",
        "Coll.Lower Lim.", "Coll. Upper Lim.", "Coll.Charge",
        "GT Lower Lim.", "GT Upper Lim.", "GT Charge",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def parse_sku_cat(raw: pd.DataFrame) -> dict:
    """Parse Sheet3 (Category Description) → {SKU_UPPER: sub_category_raw}."""
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(
        df.iloc[:, 0].astype(str).str.strip().str.upper(),
        df.iloc[:, 1].astype(str).str.strip(),
    ))


def parse_pwn_dict(raw: pd.DataFrame) -> dict:
    """Parse Sheet4 (Price We Need) → {SKU_UPPER: pwn_value}."""
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
                       fixed_fee, gst_rate,
                       pwn_overrides: dict = None) -> pd.DataFrame:
    """
    Per-order calculation:

        GT Amount        = Fixed ₹ charge from GT slab (Invoice Amount → slab)
        Selling Price    = Invoice Amount − GT Amount       ← derived, NOT from CSV SP column

        Commission       = Selling Price × Commission %    (slab by Selling Price)
        Collection Fee   = Selling Price × Collection %    (slab by Selling Price)
        Total Charges    = Commission + Collection Fee + Fixed Fee

        GST              = Total Charges × 18%

        Total Deductions = Total Charges + GST             (NO TDS, NO TCS)
        Received Amount  = Selling Price − Total Deductions

        Difference       = Received Amount − PWN

        If category or GT slab not found → Selling Price and all charges = NaN
    """
    pwn_overrides = pwn_overrides or {}
    rows_out = []

    for _, row in order_df.iterrows():
        raw_sku    = str(row.get("SKU", "")).strip()
        sku        = strip_vendor_prefix(raw_sku)
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)

        # ── Category ──────────────────────────────────────────────────
        sub_cat_raw = sku_cat_dict.get(sku.upper(), "")
        cat         = get_cat_for_lookup(sub_cat_raw, charges_df)

        # ── GT lookup from Invoice Amount slab ────────────────────────
        gt_val = get_gt_amount(cat, inv_amount, charges_df) if cat else np.nan

        # ── Selling Price = Invoice - GT ──────────────────────────────
        # If no category or GT slab → leave everything as NaN
        if pd.isna(gt_val) or not cat:
            sell_price = np.nan
            gt_val     = np.nan
            commission = coll_fee = total_charges = np.nan
            gst_on_charges = total_deductions = received_amount = np.nan
        else:
            sell_price       = round((inv_amount - gt_val) * quantity, 5)
            commission       = get_commission(cat, sell_price, charges_df)
            coll_fee         = get_collection_fee(cat, sell_price, charges_df)
            total_charges    = round(commission + coll_fee + float(fixed_fee), 5)
            gst_on_charges   = round(total_charges * gst_rate, 5)
            total_deductions = round(total_charges + gst_on_charges, 5)
            received_amount  = round(sell_price - total_deductions, 5)

        # ── PWN ───────────────────────────────────────────────────────
        pwn_val, match_method = lookup_pwn(sku, pwn_dict)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku.upper()]), "manual"

        difference = (
            round(received_amount - pwn_val, 5)
            if (pd.notna(received_amount) and pd.notna(pwn_val))
            else np.nan
        )

        rows_out.append({
            "Order Id":         order_id,
            "SKU":              raw_sku,
            "Lookup SKU":       sku,
            "Ordered On":       ordered_on,
            "Category":         sub_cat_raw,
            "Qty":              quantity,
            "Invoice Amount":   inv_amount,
            "GT (As Per Calc)": gt_val,
            "Selling Price":    sell_price,
            "Commission":       commission,
            "Collection Fee":   coll_fee,
            "Fixed Fee":        float(fixed_fee),
            "Total Charges":    total_charges,
            "GST on Charges":   gst_on_charges,
            "Total Deductions": total_deductions,
            "Received Amount":  received_amount,
            "PWN":              pwn_val,
            "PWN Match":        match_method,
            "Difference":       difference,
        })

    return pd.DataFrame(rows_out)


# ═══════════════════════════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

MONEY_COLS = [
    "Invoice Amount", "GT (As Per Calc)", "Selling Price",
    "Commission", "Collection Fee", "Fixed Fee",
    "Total Charges", "GST on Charges",
    "Total Deductions", "Received Amount", "PWN", "Difference",
]

def fmt_inr(x):
    try:
        if pd.isna(x): return "—"
        return f"₹{float(x):,.2f}"
    except Exception:
        return str(x)

def style_table(df: pd.DataFrame, diff_col: str = "Difference") -> object:
    fmt_dict = {c: fmt_inr for c in df.columns if c in MONEY_COLS}

    def colour_diff(val):
        try:
            v = float(val)
            if v < 0: return "color: red; font-weight: bold"
            if v > 0: return "color: green; font-weight: bold"
        except Exception:
            pass
        return ""

    styler = df.style.format(fmt_dict)
    if diff_col in df.columns:
        styler = styler.applymap(colour_diff, subset=[diff_col])
    return styler


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def to_excel(recon_df: pd.DataFrame, summary_df: pd.DataFrame,
             cat_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for df, sheet in [
            (recon_df,   "Reconciliation"),
            (cat_df,     "Category Breakdown"),
            (summary_df, "Charges Summary"),
        ]:
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
    valid = df[df["Received Amount"].notna()]

    totals = {"Metric": [], "Value": []}
    fields = [
        ("Total Orders",               len(df)),
        ("Orders Calculated",          int(df["Received Amount"].notna().sum())),
        ("Orders NaN (no category)",   int(df["Received Amount"].isna().sum())),
        ("Total Invoice Amount",       df["Invoice Amount"].sum()),
        ("Total GT (As Per Calc)",     valid["GT (As Per Calc)"].sum()),
        ("Total Selling Price",        valid["Selling Price"].sum()),
        ("Total Commission",           valid["Commission"].sum()),
        ("Total Collection Fee",       valid["Collection Fee"].sum()),
        ("Total Fixed Fee",            valid["Fixed Fee"].sum()),
        ("Total Charges (C+F+Fixed)",  valid["Total Charges"].sum()),
        ("Total GST on Charges",       valid["GST on Charges"].sum()),
        ("Total Deductions",           valid["Total Deductions"].sum()),
        ("Total Received Amount",      valid["Received Amount"].sum()),
        ("Net Difference vs PWN",      valid["Difference"].sum()),
        ("Orders with -ve Diff",       int((valid["Difference"] < 0).sum())),
        ("Orders with +ve Diff",       int((valid["Difference"] > 0).sum())),
        ("Orders – No PWN found",      int(df["Difference"].isna().sum())),
        ("Avg Received per Order",     valid["Received Amount"].mean()),
        ("Avg Difference per Order",   valid["Difference"].mean()),
    ]
    for label, val in fields:
        totals["Metric"].append(label)
        totals["Value"].append(round(val, 2) if isinstance(val, float) else val)
    summary_df = pd.DataFrame(totals)

    cat_df = (
        valid.groupby("Category")
        .agg(
            Orders         = ("Order Id",         "count"),
            Invoice_Total  = ("Invoice Amount",    "sum"),
            GT_Total       = ("GT (As Per Calc)",  "sum"),
            Selling_Total  = ("Selling Price",     "sum"),
            Commission     = ("Commission",        "sum"),
            Collection     = ("Collection Fee",    "sum"),
            Fixed          = ("Fixed Fee",         "sum"),
            Total_Charges  = ("Total Charges",     "sum"),
            GST_Total      = ("GST on Charges",    "sum"),
            Deductions     = ("Total Deductions",  "sum"),
            Received_Total = ("Received Amount",   "sum"),
            Net_Diff       = ("Difference",        "sum"),
            Avg_Diff       = ("Difference",        "mean"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
        .round(2)
    )
    cat_df.columns = [
        "Category", "Orders", "Invoice Total", "GT Total", "Selling Total",
        "Commission", "Collection Fee", "Fixed Fee",
        "Total Charges", "GST Total", "Total Deductions",
        "Received Total", "Net Difference", "Avg Difference",
    ]
    return summary_df, cat_df


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k, v in [
    ("pwn_overrides", {}),
    ("result_df",     None),
    ("charges_df",    None),
    ("sku_cat_dict",  None),
    ("pwn_dict",      None),
    ("order_df",      None),
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
        if len(sheets) < 4:
            st.error("❌ Excel must have at least 4 sheets (Sheet2=Charges, Sheet3=Category, Sheet4=PWN).")
            st.stop()
        # Sheet index: 0=Sheet1(ignored), 1=Sheet2(Charges), 2=Sheet3(Category), 3=Sheet4(PWN)
        charges_df   = parse_charges_df(sheets[1])
        sku_cat_dict = parse_sku_cat(sheets[2])
        pwn_dict     = parse_pwn_dict(sheets[3])
        st.session_state.update({
            "charges_df":   charges_df,
            "sku_cat_dict": sku_cat_dict,
            "pwn_dict":     pwn_dict,
            "order_df":     order_df,
        })

    result_df = run_reconciliation(
        order_df, charges_df, sku_cat_dict, pwn_dict,
        fixed_fee, gst_rate,
        pwn_overrides=st.session_state["pwn_overrides"],
    )
    st.session_state["result_df"] = result_df
    summary_df, cat_df = build_summary(result_df)

    st.success(f"✅ Processed **{len(result_df):,}** orders")

    tab1, tab2, tab3 = st.tabs([
        "📋  Reconciliation",
        "💰  Charges Summary",
        "📊  Category Breakdown",
    ])

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 1 – RECONCILIATION                                         ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab1:

        # ── PWN Not-Found editor ──────────────────────────────────────
        missing_df = result_df[result_df["PWN Match"] == "not_found"]
        if len(missing_df):
            with st.expander(
                f"⚠️  **{len(missing_df)} SKU(s) have no PWN price — click to enter manually**",
                expanded=True,
            ):
                st.info(
                    "Each row below shows the full SKU. "
                    "Enter the correct PWN value and click **Save & Recalculate**."
                )
                missing_skus   = missing_df["SKU"].unique().tolist()
                override_inputs = {}

                for sku in missing_skus:
                    c_label, c_input = st.columns([3, 2])
                    c_label.markdown(
                        f"<div style='padding-top:8px;font-size:0.95rem;"
                        f"word-break:break-all'><b>{sku}</b></div>",
                        unsafe_allow_html=True,
                    )
                    stripped = strip_vendor_prefix(sku)
                    existing = float(st.session_state["pwn_overrides"].get(stripped.upper(), 0.0))
                    override_inputs[stripped] = c_input.number_input(
                        "PWN (₹)", value=existing, min_value=0.0, step=0.5,
                        label_visibility="collapsed", key=f"pwn_input_{sku}",
                    )

                if st.button("💾  Save PWN Overrides & Recalculate", type="primary"):
                    for sku, val in override_inputs.items():
                        if val > 0:
                            st.session_state["pwn_overrides"][sku.upper()] = val
                    result_df = run_reconciliation(
                        order_df, charges_df, sku_cat_dict, pwn_dict,
                        fixed_fee, gst_rate,
                        pwn_overrides=st.session_state["pwn_overrides"],
                    )
                    st.session_state["result_df"] = result_df
                    summary_df, cat_df = build_summary(result_df)
                    st.success("✅ Saved and recalculated!")
                    st.rerun()

        # ── KPI strip ────────────────────────────────────────────────
        st.markdown("### 📊 Summary")
        valid = result_df[result_df["Received Amount"].notna()]
        k1,k2,k3,k4,k5,k6,k7,k8 = st.columns(8)
        k1.metric("Orders",            f"{len(result_df):,}")
        k2.metric("Invoice Total",     f"₹{result_df['Invoice Amount'].sum():,.0f}")
        k3.metric("GT Total (ref)",    f"₹{valid['GT (As Per Calc)'].sum():,.0f}")
        k4.metric("Selling Pr. Total", f"₹{valid['Selling Price'].sum():,.0f}")
        k5.metric("Total Charges",     f"₹{valid['Total Charges'].sum():,.0f}")
        k6.metric("GST on Charges",    f"₹{valid['GST on Charges'].sum():,.0f}")
        k7.metric("Received Total",    f"₹{valid['Received Amount'].sum():,.0f}")
        net = valid["Difference"].sum()
        k8.metric(
            "Net Difference",
            f"₹{net:,.2f}",
            delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
            delta_color="normal" if net >= 0 else "inverse",
        )

        st.markdown("---")

        # ── Filters ──────────────────────────────────────────────────
        f1, f2, f3 = st.columns([2, 2, 3])
        all_cats = ["All"] + sorted(result_df["Category"].dropna().unique().tolist())
        sel_cat  = f1.selectbox("Category", all_cats)
        diff_opt = f2.selectbox("Difference type",
                                ["All", "Positive (+)", "Negative (−)",
                                 "Zero / Matched", "No PWN data", "No Category (NaN)"])
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
            view = view[view["PWN Match"] == "not_found"]
        elif diff_opt == "No Category (NaN)":
            view = view[view["Received Amount"].isna()]
        if search.strip():
            mask = (
                view["SKU"].str.contains(search.strip(), case=False, na=False) |
                view["Order Id"].str.contains(search.strip(), case=False, na=False)
            )
            view = view[mask]

        st.caption(f"Showing **{len(view):,}** of **{len(result_df):,}** orders")

        display_cols = [
            "Order Id", "SKU", "Lookup SKU", "Ordered On", "Category",
            "Qty", "Invoice Amount",
            "GT (As Per Calc)",   # Fixed ₹ from GT slab
            "Selling Price",      # Invoice Amount − GT
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "Total Deductions",
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
            data=to_excel(result_df[display_cols], summary_df, cat_df),
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
        st.caption("Grand total of every deduction across all calculated orders")

        valid = result_df[result_df["Received Amount"].notna()]
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("#### 📤 Flipkart Deductions")
            a1, a2 = st.columns(2)
            a1.metric("Commission",     f"₹{valid['Commission'].sum():,.2f}")
            a2.metric("Collection Fee", f"₹{valid['Collection Fee'].sum():,.2f}")
            a1.metric("Fixed Fee",      f"₹{valid['Fixed Fee'].sum():,.2f}")
            a2.metric("GST on Charges", f"₹{valid['GST on Charges'].sum():,.2f}")
            st.metric(
                "🔴 Total Deductions",
                f"₹{valid['Total Deductions'].sum():,.2f}",
            )

        with col_b:
            st.markdown("#### 📥 What You Receive")
            b1, b2 = st.columns(2)
            b1.metric("Total Invoice",   f"₹{result_df['Invoice Amount'].sum():,.2f}")
            b2.metric("GT Total (ref)",  f"₹{valid['GT (As Per Calc)'].sum():,.2f}")
            b1.metric("Selling Total",   f"₹{valid['Selling Price'].sum():,.2f}")
            b2.metric("Total Received",  f"₹{valid['Received Amount'].sum():,.2f}")
            net = valid["Difference"].sum()
            b1.metric(
                "Net Diff vs PWN",
                f"₹{net:,.2f}",
                delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
                delta_color="normal" if net >= 0 else "inverse",
            )
            b2.metric("Orders –ve Diff", int((valid["Difference"] < 0).sum()))

        st.info(
            "ℹ️  **Selling Price** = Invoice Amount − GT Amount (GT looked up from Invoice slab).  "
            "**GT Amount** is the fixed charge deducted to arrive at Selling Price — "
            "it is displayed for reference but is already accounted for in Selling Price."
        )

        st.markdown("---")
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols = [
            "Order Id", "SKU", "Category",
            "Invoice Amount", "GT (As Per Calc)", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "Total Deductions", "Received Amount",
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
        st.caption("Every charge component summed per product category (NaN rows excluded)")

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
            "GT Total", "Commission", "Collection Fee", "Fixed Fee",
            "GST Total", "Total Deductions",
        ]
        comp_money = [c for c in comp_cols if c not in ("Category", "Orders")]
        st.dataframe(
            cat_df[comp_cols].style.format({c: "₹{:.2f}" for c in comp_money}),
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
| **Data Excel** | 5-sheet workbook: Sheet2=Charges / Sheet3=Category / Sheet4=Price We Need |

---
### Calculation per order

```
GT Amount        = Fixed ₹ charge from GT slab  (Invoice Amount → slab lookup)

Selling Price    = Invoice Amount − GT Amount    ← NOT the CSV SP column

Commission       = Selling Price × Commission %  (slab lookup by Selling Price)
Collection Fee   = Selling Price × Collection %  (slab lookup by Selling Price)
Total Charges    = Commission + Collection Fee + Fixed Fee (₹5)

GST              = Total Charges × 18%

Total Deductions = Total Charges + GST           ← NO TDS, NO TCS
Received Amount  = Selling Price − Total Deductions

Difference       = Received Amount − PWN
```

**If category or GT slab not found** → all charges shown as — (not calculated).

**PWN lookup:** tries exact SKU first, then auto-maps combined sizes  
(e.g. order SKU `7053YKBLS-L` → matches price sheet `7053YKBLS-L-XL`)

**SKU prefix stripping:** vendor prefix `GWN-` / `GWN_` stripped before lookups.

**Category matching:** sub-category names from Sheet3 are matched to Sheet2  
using a case-insensitive lookup (e.g. `kurta` → `Kurta`, `trouser` → `Pant`,  
`ethnic_set`, `blouse`, `dress`, `night_suit` etc. match directly).
""")
