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
[data-testid="stMetricValue"] { font-size: 1.35rem; font-weight: 700; }
.block-container { padding-top: 1.4rem; }
div[data-testid="stDataFrame"] table thead tr th {
    background-color: #1f4e79 !important;
    color: white !important;
}
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
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5, min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)", value=18, min_value=0, step=1) / 100
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

# Price-sheet size suffix  →  list of individual order sizes it covers
SIZE_EXPAND = {
    "L-XL":      ["L", "XL"],
    "S-M":       ["S", "M"],
    "XXL-3XL":   ["XXL", "3XL"],
    "F-S/XXL":   ["F"],
    "F-3xl/5xl": ["F"],
    "XS-S":      ["XS", "S"],
    "M-L":       ["M", "L"],
    "XL-XXL":    ["XL", "XXL"],
    "3XL-4XL":   ["3XL", "4XL"],
    "5XL-6XL":   ["5XL", "6XL"],
    "7XL-8XL":   ["7XL", "8XL"],
    "4XL-5XL":   ["4XL", "5XL"],
    "2XL-3XL":   ["2XL", "3XL"],
    "XS-S-M":    ["XS", "S", "M"],
    "L-XL-XXL":  ["L", "XL", "XXL"],
}

# Reverse: order size → candidate price-sheet suffixes to try
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
    """
    Returns (pwn_value, match_method).
    match_method: 'direct' | '<size-range-key>' | 'not_found'
    """
    key = sku.strip().upper()
    # 1. Direct
    val = pwn_dict.get(key)
    if pd.notna(val):
        return float(val), "direct"
    # 2. Size-range fallback
    parts = key.rsplit("-", 1)
    if len(parts) == 2:
        base, size = parts[0], parts[1]
        for price_size in ORDER_TO_PRICE_SIZE.get(size, []):
            candidate = f"{base}-{price_size.upper()}"
            val = pwn_dict.get(candidate)
            if pd.notna(val):
                return float(val), price_size
    return np.nan, "not_found"


def get_gt_charge(cat_norm: str, inv_amount: float, charges_df: pd.DataFrame) -> float:
    """GT charge is a fixed slab value looked up by Invoice Amount."""
    rows = charges_df[charges_df["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo, hi, gtv = r.get("GT Lower Lim."), r.get("GT Upper Lim."), r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gtv):
            if float(lo) <= inv_amount <= float(hi):
                return float(gtv)
    return 0.0


def get_commission(cat_norm: str, sell_price: float, charges_df: pd.DataFrame) -> float:
    """Commission % applied to Selling Price."""
    rows = charges_df[charges_df["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo, hi = r.get("Lower Lim."), r.get("Upper Lim.")
        ch = r.get("Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell_price <= float(hi):
                return float(ch) * sell_price
    return 0.0


def get_collection_fee(cat_norm: str, sell_price: float, charges_df: pd.DataFrame) -> float:
    """Collection fee % applied to Selling Price."""
    rows = charges_df[charges_df["Category"].str.lower() == cat_norm.lower()]
    for _, r in rows.iterrows():
        lo_raw = r.get("Coll.Lower Lim.")
        hi     = r.get("Coll. Upper Lim.")
        cf     = r.get("Coll.Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).startswith(">")) else float(lo_raw)
        if lo_val < sell_price <= float(hi):
            return float(cf) * sell_price
    return 0.0


def parse_charges_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()
    df.columns = raw_df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df["Category"].notna()].copy()
    df["Category"] = df["Category"].ffill()
    num_cols = [
        "Lower Lim.", "Upper Lim.", "Charge",
        "Coll.Lower Lim.", "Coll. Upper Lim.", "Coll.Charge",
        "GT Lower Lim.", "GT Upper Lim.", "GT Charge",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def parse_sku_cat(raw_df: pd.DataFrame) -> dict:
    df = raw_df.copy()
    df.columns = raw_df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(
        zip(df.iloc[:, 0].astype(str).str.strip().str.upper(),
            df.iloc[:, 1].astype(str).str.strip())
    )


def parse_price_dict(raw_df: pd.DataFrame) -> dict:
    df = raw_df.copy()
    df.columns = raw_df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"] = pd.to_numeric(df["PWN+10%+50"], errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(), df["PWN+10%+50"]))


def run_reconciliation(order_df, charges_df, sku_cat_dict, pwn_dict,
                       fixed_fee, gst_rate, pwn_overrides: dict = None):
    pwn_overrides = pwn_overrides or {}
    rows_out = []

    for _, row in order_df.iterrows():
        sku        = str(row.get("SKU", "")).strip()
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)

        # Sub-category
        sub_cat_raw  = sku_cat_dict.get(sku.upper(), "")
        sub_cat_norm = normalize_cat(sub_cat_raw)

        # ── GT Charge → Selling Price ───────────────────────────────────
        gt_charge  = get_gt_charge(sub_cat_norm, inv_amount, charges_df)
        sell_price = round(inv_amount - gt_charge, 2)

        # ── Commission & Collection (based on Selling Price) ────────────
        commission = get_commission(sub_cat_norm, sell_price, charges_df)
        coll_fee   = get_collection_fee(sub_cat_norm, sell_price, charges_df)

        # ── Total Charges & GST ─────────────────────────────────────────
        total_charges  = commission + coll_fee + float(fixed_fee)
        gst_on_charges = round(total_charges * gst_rate, 5)

        # ── Received Amount ─────────────────────────────────────────────
        received_amount = round(sell_price - total_charges - gst_on_charges, 2)

        # ── PWN lookup (direct → size-range → override) ─────────────────
        pwn_val, match_method = lookup_pwn(sku, pwn_dict)
        # Check manual override
        if sku.upper() in pwn_overrides:
            pwn_val      = pwn_overrides[sku.upper()]
            match_method = "manual"

        pwn_status = match_method  # 'direct' | size-range key | 'manual' | 'not_found'

        # ── Difference ──────────────────────────────────────────────────
        difference = round(received_amount - pwn_val, 2) if pd.notna(pwn_val) else np.nan

        rows_out.append({
            "Order Id":            order_id,
            "SKU":                 sku,
            "Ordered On":          ordered_on,
            "Category":            sub_cat_raw,
            "Invoice Amount (₹)":  inv_amount,
            "GT Charge (₹)":       gt_charge,
            "Selling Price (₹)":   sell_price,
            "Commission (₹)":      round(commission, 4),
            "Collection Fee (₹)":  round(coll_fee, 4),
            "Fixed Fee (₹)":       float(fixed_fee),
            "Total Charges (₹)":   round(total_charges, 4),
            "GST Amount (₹)":      gst_on_charges,
            "Received Amount (₹)": received_amount,
            "PWN (₹)":             pwn_val,
            "PWN Match":           pwn_status,
            "Difference (₹)":      difference,
            "Quantity":            quantity,
        })

    return pd.DataFrame(rows_out)


def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reconciliation")
        ws = writer.sheets["Reconciliation"]
        for col_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 36)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE  (PWN overrides persist across reruns)
# ═══════════════════════════════════════════════════════════════════════════════
if "pwn_overrides" not in st.session_state:
    st.session_state["pwn_overrides"] = {}   # {SKU_UPPER: float}
if "result_df" not in st.session_state:
    st.session_state["result_df"] = None
if "charges_df" not in st.session_state:
    st.session_state["charges_df"] = None
if "sku_cat_dict" not in st.session_state:
    st.session_state["sku_cat_dict"] = None
if "pwn_dict" not in st.session_state:
    st.session_state["pwn_dict"] = None
if "order_df" not in st.session_state:
    st.session_state["order_df"] = None


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
        pwn_dict     = parse_price_dict(sheets[2])

        st.session_state["charges_df"]   = charges_df
        st.session_state["sku_cat_dict"] = sku_cat_dict
        st.session_state["pwn_dict"]     = pwn_dict
        st.session_state["order_df"]     = order_df

    # ── Run reconciliation ───────────────────────────────────────────────
    result_df = run_reconciliation(
        order_df, charges_df, sku_cat_dict, pwn_dict,
        fixed_fee, gst_rate,
        pwn_overrides=st.session_state["pwn_overrides"],
    )
    st.session_state["result_df"] = result_df

    st.success(f"✅ Processed **{len(result_df):,}** orders")

    # ── PWN Not-Found banner ─────────────────────────────────────────────
    missing_df = result_df[result_df["PWN Match"] == "not_found"].copy()
    if len(missing_df):
        with st.expander(
            f"⚠️  **{len(missing_df)} order(s) have no PWN price** — click to enter manually",
            expanded=True,
        ):
            st.markdown("Enter the correct PWN for each SKU, then click **Save PWN Overrides**.")
            missing_skus = missing_df["SKU"].unique().tolist()
            override_inputs = {}
            cols = st.columns(min(len(missing_skus), 4))
            for i, sku in enumerate(missing_skus):
                existing = st.session_state["pwn_overrides"].get(sku.upper(), 0.0)
                override_inputs[sku] = cols[i % len(cols)].number_input(
                    f"PWN for **{sku}**",
                    value=float(existing),
                    min_value=0.0,
                    step=0.5,
                    key=f"pwn_input_{sku}",
                )
            if st.button("💾 Save PWN Overrides & Recalculate", type="primary"):
                for sku, val in override_inputs.items():
                    if val > 0:
                        st.session_state["pwn_overrides"][sku.upper()] = val
                # Rerun to recalculate with overrides
                result_df = run_reconciliation(
                    order_df, charges_df, sku_cat_dict, pwn_dict,
                    fixed_fee, gst_rate,
                    pwn_overrides=st.session_state["pwn_overrides"],
                )
                st.session_state["result_df"] = result_df
                st.success("✅ PWN overrides saved and reconciliation updated!")
                st.rerun()

    # ── KPI Cards ────────────────────────────────────────────────────────
    st.markdown("### 📊 Summary")
    total_inv      = result_df["Invoice Amount (₹)"].sum()
    total_sell     = result_df["Selling Price (₹)"].sum()
    total_charges  = result_df["Total Charges (₹)"].sum()
    total_gst      = result_df["GST Amount (₹)"].sum()
    total_received = result_df["Received Amount (₹)"].sum()
    net_diff       = result_df["Difference (₹)"].sum()
    pos_diff       = int((result_df["Difference (₹)"] > 0).sum())
    neg_diff       = int((result_df["Difference (₹)"] < 0).sum())

    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric("Orders",            f"{len(result_df):,}")
    k2.metric("Total Invoice",     f"₹{total_inv:,.0f}")
    k3.metric("Total Selling Pr.", f"₹{total_sell:,.0f}")
    k4.metric("Total Charges",     f"₹{total_charges:,.0f}")
    k5.metric("Total GST",         f"₹{total_gst:,.0f}")
    k6.metric("Total Received",    f"₹{total_received:,.0f}")
    k7.metric(
        "Net Difference",
        f"₹{net_diff:,.2f}",
        delta=f"{'▲' if net_diff >= 0 else '▼'} {abs(net_diff):,.2f}",
        delta_color="normal" if net_diff >= 0 else "inverse",
    )

    st.markdown("---")

    # ── Filters ───────────────────────────────────────────────────────────
    st.markdown("### 🔍 Filter & View")
    f1, f2, f3 = st.columns([2, 2, 3])
    all_cats = ["All"] + sorted(result_df["Category"].dropna().unique().tolist())
    sel_cat  = f1.selectbox("Category", all_cats)
    diff_opt = f2.selectbox(
        "Difference type",
        ["All", "Positive (+)", "Negative (−)", "Zero / Matched", "No PWN data"],
    )
    search = f3.text_input("🔎 Search by SKU or Order ID")

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

    # Display columns
    display_cols = [
        "Order Id", "SKU", "Ordered On", "Category",
        "Invoice Amount (₹)", "GT Charge (₹)", "Selling Price (₹)",
        "Commission (₹)", "Collection Fee (₹)", "Fixed Fee (₹)",
        "Total Charges (₹)", "GST Amount (₹)",
        "Received Amount (₹)", "PWN (₹)", "Difference (₹)", "PWN Match",
    ]

    def colour_diff(val):
        try:
            v = float(val)
            if v < 0: return "color:red;font-weight:bold"
            if v > 0: return "color:green;font-weight:bold"
        except Exception:
            pass
        return ""

    def fmt_inr(x):
        return f"₹{x:.2f}" if pd.notna(x) else "—"

    money_cols = [
        "Invoice Amount (₹)", "GT Charge (₹)", "Selling Price (₹)",
        "Commission (₹)", "Collection Fee (₹)", "Fixed Fee (₹)",
        "Total Charges (₹)", "GST Amount (₹)", "Received Amount (₹)",
        "PWN (₹)", "Difference (₹)",
    ]

    styled = (
        view[display_cols]
        .style
        .applymap(colour_diff, subset=["Difference (₹)"])
        .format({c: fmt_inr for c in money_cols})
    )
    st.dataframe(styled, use_container_width=True, height=500)

    # ── Downloads ─────────────────────────────────────────────────────────
    st.markdown("### 📥 Download")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇  Full Reconciliation (Excel)",
        data=to_excel(result_df),
        file_name="flipkart_reconciliation_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button(
        "⬇  Filtered View (Excel)",
        data=to_excel(view[display_cols]),
        file_name="flipkart_reconciliation_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Category breakdown ────────────────────────────────────────────────
    st.markdown("### 📈 Category-wise Breakdown")
    grp = (
        result_df.groupby("Category")
        .agg(
            Orders         = ("Order Id",            "count"),
            Invoice_Total  = ("Invoice Amount (₹)",   "sum"),
            Selling_Total  = ("Selling Price (₹)",     "sum"),
            Charges_Total  = ("Total Charges (₹)",     "sum"),
            Received_Total = ("Received Amount (₹)",   "sum"),
            Avg_Diff       = ("Difference (₹)",        "mean"),
            Net_Diff       = ("Difference (₹)",        "sum"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
        .round(2)
    )
    grp.columns = [
        "Category", "Orders", "Invoice Total (₹)", "Selling Total (₹)",
        "Charges Total (₹)", "Received Total (₹)", "Avg Diff (₹)", "Net Diff (₹)",
    ]
    st.dataframe(grp, use_container_width=True)

# ── Landing screen ─────────────────────────────────────────────────────────────
else:
    st.info("👈 Upload **both files** in the sidebar to begin.")
    st.markdown("""
---
### How it works

| Step | File | Used columns |
|------|------|-------------|
| 1 | **Order CSV** | `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Selling Price Per Item`, `Quantity` |
| 2 | **Data Excel** | 3-sheet workbook (Charges / Category / Price We Need) |

---
### Calculation per order

```
GT Charge       = fixed slab from Charges sheet  (looked up by Invoice Amount)
Selling Price   = Invoice Amount − GT Charge

Commission      = Selling Price × Commission %   (slab by Selling Price)
Collection Fee  = Selling Price × Collection %   (slab by Selling Price)
Total Charges   = Commission + Collection Fee + Fixed Fee

GST Amount      = Total Charges × 18 %
Received Amount = Selling Price − Total Charges − GST Amount

Difference      = Received Amount − PWN (Price We Need)
```

**PWN lookup** tries the exact SKU first, then auto-maps combined sizes:
`L-XL → L or XL`, `S-M → S or M`, `XXL-3XL → XXL or 3XL`, etc.

If a PWN is still not found, you can enter it manually and recalculate instantly.
""")
