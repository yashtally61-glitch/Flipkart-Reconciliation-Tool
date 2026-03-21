import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Flipkart Reconciliation – Ashirwad Garments",
    layout="wide",
    page_icon="🧾",
)

# ──────────────────────────────────────────────────────────────────────────────
# Styling
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 700; }
.block-container { padding-top: 1.5rem; }
thead tr th { background-color: #1f4e79 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

st.title("🧾 Flipkart Reconciliation Tool")
st.caption("Ashirwad Garments — auto-calculate Flipkart charges & reconcile orders")

# ──────────────────────────────────────────────────────────────────────────────
# Sidebar – uploads + settings
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Upload Files")
    order_file   = st.file_uploader("1️⃣  Order CSV (Flipkart export)", type=["csv"])
    charges_file = st.file_uploader("2️⃣  Data Excel (3-sheet workbook)", type=["xlsx"])

    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5, min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)", value=18, min_value=0, step=1) / 100

    st.markdown("---")
    st.markdown("""
**Excel sheet layout expected:**
- **Sheet 1** → `Charges Description` (commission / collection / GT slabs)
- **Sheet 2** → `Category Description` (SKU → sub-category)
- **Sheet 3** → `Price We Need` (SKU → PWN+10%+50)
""")


# ──────────────────────────────────────────────────────────────────────────────
# Helper – sub-category → Flipkart tier name
# ──────────────────────────────────────────────────────────────────────────────
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

def normalize_cat(raw):
    if pd.isna(raw):
        return ""
    return CAT_MAP.get(str(raw).strip().lower(), str(raw).strip().title())


# ──────────────────────────────────────────────────────────────────────────────
# Helper – lookup charges from the charges dataframe
# ──────────────────────────────────────────────────────────────────────────────
def lookup_charges(cat_norm: str, inv_amount: float, charges_df: pd.DataFrame):
    """
    Returns (commission_amt, collection_fee, gt_charge)
    charges_df already has proper column names and numeric values.
    """
    rows = charges_df[charges_df["Category"].str.lower() == cat_norm.lower()]
    if rows.empty:
        return 0.0, 0.0, 0.0

    # --- Commission (% of invoice) ----------------------------------------
    comm_pct = np.nan
    for _, r in rows.iterrows():
        lo, hi = r.get("Lower Lim."), r.get("Upper Lim.")
        if pd.notna(lo) and pd.notna(hi) and lo <= inv_amount <= hi:
            comm_pct = r.get("Charge", np.nan)
            break
    commission = float(comm_pct) * inv_amount if pd.notna(comm_pct) else 0.0

    # --- Collection fee (% of invoice) ------------------------------------
    coll_fee = 0.0
    for _, r in rows.iterrows():
        lo_raw = r.get("Coll.Lower Lim.")
        hi     = r.get("Coll. Upper Lim.")
        cf     = r.get("Coll.Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        # "Coll.Lower Lim." may be stored as 0 for "> ₹0" rows
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).startswith(">")) else float(lo_raw)
        if lo_val < inv_amount <= float(hi):
            coll_fee = float(cf) * inv_amount
            break

    # --- GT Charge (fixed slab value) -------------------------------------
    gt_charge = 0.0
    for _, r in rows.iterrows():
        lo, hi = r.get("GT Lower Lim."), r.get("GT Upper Lim.")
        gtv    = r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gtv):
            if float(lo) <= inv_amount <= float(hi):
                gt_charge = float(gtv)
                break

    return commission, coll_fee, gt_charge


# ──────────────────────────────────────────────────────────────────────────────
# Core reconciliation function
# ──────────────────────────────────────────────────────────────────────────────
def run_reconciliation(order_df, charges_raw, sku_cat_raw, price_raw, fixed_fee, gst_rate):

    # ── Parse Charges sheet ─────────────────────────────────────────────
    charges_df = charges_raw.copy()
    charges_df.columns = charges_raw.iloc[0].tolist()
    charges_df = charges_df.iloc[1:].reset_index(drop=True)
    charges_df = charges_df[charges_df["Category"].notna()].copy()
    charges_df["Category"] = charges_df["Category"].ffill()
    for col in ["Lower Lim.", "Upper Lim.", "Charge",
                "Coll.Lower Lim.", "Coll. Upper Lim.", "Coll.Charge",
                "GT Lower Lim.", "GT Upper Lim.", "GT Charge"]:
        if col in charges_df.columns:
            charges_df[col] = pd.to_numeric(charges_df[col], errors="coerce")

    # ── Parse SKU-category sheet ─────────────────────────────────────────
    sku_cat_df = sku_cat_raw.copy()
    sku_cat_df.columns = sku_cat_raw.iloc[0].tolist()
    sku_cat_df = sku_cat_df.iloc[1:].reset_index(drop=True)
    col0, col1 = sku_cat_df.columns[0], sku_cat_df.columns[1]
    sku_cat_dict = dict(
        zip(sku_cat_df[col0].astype(str).str.strip().str.upper(),
            sku_cat_df[col1].astype(str).str.strip())
    )

    # ── Parse Price sheet ────────────────────────────────────────────────
    price_df = price_raw.copy()
    price_df.columns = price_raw.iloc[0].tolist()
    price_df = price_df.iloc[1:].reset_index(drop=True)
    price_df["OMS Child SKU"] = price_df["OMS Child SKU"].astype(str).str.strip()
    pwn_col = "PWN+10%+50"
    price_df[pwn_col] = pd.to_numeric(price_df[pwn_col], errors="coerce")
    pwn_dict = dict(zip(price_df["OMS Child SKU"].str.upper(), price_df[pwn_col]))

    # ── Process each order row ───────────────────────────────────────────
    rows_out = []
    for _, row in order_df.iterrows():
        sku        = str(row.get("SKU", "")).strip()
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)
        sell_price = float(row.get("Selling Price Per Item", 0) or 0)

        # PWN lookup
        pwn = pwn_dict.get(sku.upper(), np.nan)

        # Category lookup
        sub_cat_raw = sku_cat_dict.get(sku.upper(), "")
        sub_cat_norm = normalize_cat(sub_cat_raw)

        # Charges
        commission, coll_fee, gt_charge = lookup_charges(sub_cat_norm, inv_amount, charges_df)

        total_charges   = commission + coll_fee + float(fixed_fee) + gt_charge
        gst_on_charges  = round(total_charges * gst_rate, 5)
        received_amount = round(inv_amount - total_charges - gst_on_charges, 2)

        # Difference vs PWN
        difference = round(received_amount - pwn, 2) if pd.notna(pwn) else np.nan

        rows_out.append({
            "Order Id":          order_id,
            "SKU":               sku,
            "Ordered On":        ordered_on,
            "PWN (₹)":           pwn,
            "Invoice Amount":    inv_amount,
            "Selling Price":     sell_price,
            "Quantity":          quantity,
            "Category":          sub_cat_raw,
            "Commission (₹)":    round(commission, 4),
            "Commission %":      round(commission / inv_amount * 100, 2) if inv_amount else 0,
            "Collection Fee (₹)": round(coll_fee, 4),
            "GT Charge (₹)":     gt_charge,
            "Fixed Fee (₹)":     fixed_fee,
            "Total Charges (₹)": round(total_charges, 4),
            "GST on Charges (₹)": gst_on_charges,
            "Received Amount (₹)": received_amount,
            "Difference (₹)":    difference,
        })

    return pd.DataFrame(rows_out)


# ──────────────────────────────────────────────────────────────────────────────
# Excel export helper
# ──────────────────────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reconciliation")
        ws = writer.sheets["Reconciliation"]
        # Auto-width columns
        for col_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 35)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# Main – only runs when both files are uploaded
# ──────────────────────────────────────────────────────────────────────────────
if order_file and charges_file:

    with st.spinner("🔄 Reading files and running reconciliation…"):
        order_df = pd.read_csv(order_file)

        xl = pd.read_excel(charges_file, sheet_name=None, header=None)
        sheets = list(xl.values())

        if len(sheets) < 3:
            st.error("❌ Excel workbook must have at least 3 sheets.")
            st.stop()

        charges_raw = sheets[0]   # Charges Description
        sku_cat_raw = sheets[1]   # Category Description
        price_raw   = sheets[2]   # Price We Need

        try:
            result_df = run_reconciliation(
                order_df, charges_raw, sku_cat_raw, price_raw, fixed_fee, gst_rate
            )
        except Exception as e:
            st.error(f"❌ Reconciliation error: {e}")
            st.stop()

    st.success(f"✅ Processed **{len(result_df):,}** orders successfully!")

    # ── KPI Summary ───────────────────────────────────────────────────────
    total_invoice  = result_df["Invoice Amount"].sum()
    total_charges  = result_df["Total Charges (₹)"].sum()
    total_received = result_df["Received Amount (₹)"].sum()
    matched        = result_df["Difference (₹)"].notna().sum()
    pos_diff       = int((result_df["Difference (₹)"] > 0).sum())
    neg_diff       = int((result_df["Difference (₹)"] < 0).sum())
    total_diff     = result_df["Difference (₹)"].sum()

    st.markdown("### 📊 Summary")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Orders",         f"{len(result_df):,}")
    c2.metric("Total Invoice",        f"₹{total_invoice:,.0f}")
    c3.metric("Total Charges",        f"₹{total_charges:,.0f}")
    c4.metric("Total Received (Est.)",f"₹{total_received:,.0f}")
    c5.metric("Orders (+) Diff",      f"{pos_diff}", delta=f"+{pos_diff}")
    c6.metric("Orders (−) Diff",      f"{neg_diff}", delta=f"-{neg_diff}", delta_color="inverse")

    net_col, _ = st.columns([1, 3])
    net_col.metric(
        "Net Difference (Received − PWN)",
        f"₹{total_diff:,.2f}",
        delta=f"{'▲' if total_diff >= 0 else '▼'} {abs(total_diff):,.2f}",
        delta_color="normal" if total_diff >= 0 else "inverse",
    )

    st.markdown("---")

    # ── Filters ───────────────────────────────────────────────────────────
    st.markdown("### 🔍 Filter Results")
    f1, f2, f3 = st.columns([2, 2, 3])
    cats     = ["All"] + sorted(result_df["Category"].dropna().unique().tolist())
    sel_cat  = f1.selectbox("Category", cats)
    diff_opt = f2.selectbox("Difference", ["All", "Positive (+)", "Negative (−)", "Zero / Matched", "No PWN data"])
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

    # ── Coloured table ────────────────────────────────────────────────────
    def colour_diff(val):
        try:
            v = float(val)
            if v < 0:  return "color:red; font-weight:bold"
            if v > 0:  return "color:green; font-weight:bold"
        except:
            pass
        return ""

    display_cols = [
        "Order Id", "SKU", "Ordered On", "Category", "Invoice Amount",
        "PWN (₹)", "Commission (₹)", "Collection Fee (₹)", "GT Charge (₹)",
        "Fixed Fee (₹)", "Total Charges (₹)", "GST on Charges (₹)",
        "Received Amount (₹)", "Difference (₹)",
    ]

    styled = (
        view[display_cols]
        .style
        .applymap(colour_diff, subset=["Difference (₹)"])
        .format({
            "Invoice Amount":    "₹{:.2f}",
            "PWN (₹)":          lambda x: f"₹{x:.2f}" if pd.notna(x) else "—",
            "Commission (₹)":   "₹{:.2f}",
            "Collection Fee (₹)":"₹{:.2f}",
            "GT Charge (₹)":    "₹{:.2f}",
            "Fixed Fee (₹)":    "₹{}",
            "Total Charges (₹)":"₹{:.2f}",
            "GST on Charges (₹)":"₹{:.2f}",
            "Received Amount (₹)":"₹{:.2f}",
            "Difference (₹)":   lambda x: f"₹{x:.2f}" if pd.notna(x) else "—",
        })
    )
    st.dataframe(styled, use_container_width=True, height=480)

    # ── Downloads ─────────────────────────────────────────────────────────
    st.markdown("### 📥 Download")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇ Full Reconciliation (Excel)",
        data=to_excel(result_df),
        file_name="flipkart_reconciliation_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button(
        "⬇ Filtered View (Excel)",
        data=to_excel(view[display_cols]),
        file_name="flipkart_reconciliation_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Category breakdown ────────────────────────────────────────────────
    st.markdown("### 📈 Category-wise Breakdown")
    grp = (
        result_df.groupby("Category")
        .agg(
            Orders        = ("Order Id",           "count"),
            Invoice_Total = ("Invoice Amount",      "sum"),
            Charges_Total = ("Total Charges (₹)",   "sum"),
            Received_Total= ("Received Amount (₹)", "sum"),
            Avg_Diff      = ("Difference (₹)",      "mean"),
            Net_Diff      = ("Difference (₹)",      "sum"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
    )
    for c in ["Invoice_Total", "Charges_Total", "Received_Total", "Avg_Diff", "Net_Diff"]:
        grp[c] = grp[c].round(2)
    grp.columns = ["Category", "Orders", "Invoice Total (₹)", "Charges Total (₹)",
                   "Received Total (₹)", "Avg Diff (₹)", "Net Diff (₹)"]
    st.dataframe(grp, use_container_width=True)

    # ── SKUs missing from Price sheet ─────────────────────────────────────
    missing_pwn = result_df[result_df["PWN (₹)"].isna()]["SKU"].unique()
    if len(missing_pwn):
        with st.expander(f"⚠️ {len(missing_pwn)} SKUs not found in Price We Need sheet"):
            st.write(list(missing_pwn))

# ── Landing screen ────────────────────────────────────────────────────────────
else:
    st.info("👈 Upload **both files** in the sidebar to begin reconciliation.")
    st.markdown("""
---
### How to use

| Step | File | Description |
|------|------|-------------|
| 1 | **Order CSV** | Exported from Flipkart Seller Hub — columns used: `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Selling Price Per Item`, `Quantity` |
| 2 | **Data Excel** | 3-sheet workbook with charge slabs, SKU-category map, and PWN price list |

### Calculation logic (per order)

```
Commission   = Invoice Amount × Commission % (from Charges sheet, slab-based)
Collection   = Invoice Amount × Collection % (from Charges sheet, slab-based)
GT Charge    = Fixed slab value (from Charges sheet)
Total Charges = Commission + Collection + GT Charge + Fixed Fee
GST           = Total Charges × 18%
Received Amt  = Invoice Amount − Total Charges − GST
Difference    = Received Amount − PWN (Price We Need)
```

Positive difference → you receive **more** than expected.  
Negative difference → you receive **less** than expected.
    """)
