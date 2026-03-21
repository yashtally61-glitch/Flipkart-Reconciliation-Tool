import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Flipkart Reconciliation Tool", layout="wide", page_icon="🛒")

st.markdown("""
<style>
.metric-card {background:#f0f2f6;padding:15px;border-radius:10px;text-align:center;margin:5px}
.diff-positive {color:green;font-weight:bold}
.diff-negative {color:red;font-weight:bold}
</style>
""", unsafe_allow_html=True)

st.title("🛒 Flipkart Reconciliation Tool")
st.markdown("Upload your **Order CSV** and **Price/Charges Excel** to auto-calculate charges and reconcile.")

# ─────────────────────────────────────────────
# Sidebar – file uploads
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Upload Files")
    order_file   = st.file_uploader("Order CSV (Flipkart export)", type=["csv"])
    charges_file = st.file_uploader("Charges Excel (4-sheet workbook)", type=["xlsx"])
    st.markdown("---")
    st.markdown("**Sheet layout expected:**")
    st.markdown("- Sheet1 → Reconciliation output")
    st.markdown("- Sheet2 → Category-wise charges")
    st.markdown("- Sheet3 → SKU → Sub-category mapping")
    st.markdown("- Sheet4 → SKU price list (PWN)")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5, min_value=0)


# ─────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────

def normalize_cat(cat_raw: str) -> str:
    """Normalise sub-category string to match Sheet2 Category column."""
    if pd.isna(cat_raw):
        return ""
    m = str(cat_raw).strip().lower()
    mapping = {
        "kurta": "Kurta",
        "ethnic_set": "Co-ords Set",   # fallback – many ethnic sets map here
        "top": "Top",
        "blouse": "Blouse",
        "trouser": "Pant",
        "dress": "Dresses",
        "night_suit": "Nightsuit Sets",
        "nightsuit": "Nightsuit Sets",
        "shirt": "Men's shirt",
        "kaftan": "Kaftan",
        "men_kurta": "Men's Kurta",
    }
    return mapping.get(m, str(cat_raw).strip().title())


def lookup_charges(category_norm: str, inv_amount: float, charges_df: pd.DataFrame):
    """
    Returns (commission, collection_fee, gt_charge) given category and invoice amount.
    charges_df columns: Category, Lower Lim., Upper Lim., Charge,
                        Coll.Lower Lim., Coll. Upper Lim., Coll.Charge,
                        GT Lower Lim., GT Upper Lim., GT Charge
    """
    rows = charges_df[charges_df['Category'].str.lower() == category_norm.lower()]
    if rows.empty:
        return np.nan, np.nan, np.nan

    # Commission – percentage based on invoice amount vs Lower/Upper Lim.
    commission_pct = np.nan
    for _, r in rows.iterrows():
        lo = r.get('Lower Lim.')
        hi = r.get('Upper Lim.')
        if pd.notna(lo) and pd.notna(hi):
            if lo <= inv_amount <= hi:
                commission_pct = r.get('Charge', np.nan)
                break
    commission = (commission_pct * inv_amount) if pd.notna(commission_pct) else 0.0

    # Collection fee
    coll_fee = np.nan
    for _, r in rows.iterrows():
        lo = r.get('Coll.Lower Lim.')
        hi = r.get('Coll. Upper Lim.')
        if pd.notna(lo) and pd.notna(hi):
            # Handle "> ₹0" style stored as string or numeric 0
            lo_val = 0.0 if str(lo).startswith(">") else float(lo)
            hi_val = float(hi)
            if lo_val < inv_amount <= hi_val or (lo_val == 0 and inv_amount <= hi_val):
                cf = r.get('Coll.Charge', np.nan)
                if pd.notna(cf):
                    coll_fee = cf * inv_amount if cf < 1 else cf  # percentage or flat?
                    # All values in sheet are fractions (0.006 etc.)
                    coll_fee = cf * inv_amount
                break
    if pd.isna(coll_fee):
        coll_fee = 0.0

    # GT charge (fixed slab)
    gt_charge = np.nan
    for _, r in rows.iterrows():
        lo = r.get('GT Lower Lim.')
        hi = r.get('GT Upper Lim.')
        if pd.notna(lo) and pd.notna(hi):
            if lo <= inv_amount <= hi:
                gt_charge = r.get('GT Charge', np.nan)
                break
    if pd.isna(gt_charge):
        gt_charge = 0.0

    return commission, coll_fee, gt_charge


def lookup_pwn(sku: str, price_df: pd.DataFrame):
    """Lookup PWN from Sheet4."""
    row = price_df[price_df['OMS Child SKU'].str.upper() == str(sku).upper()]
    if row.empty:
        return np.nan, np.nan
    return row.iloc[0].get('PWN+10%', np.nan), row.iloc[0].get('PWN+10%+50', np.nan)


def run_reconciliation(order_df, charges_df, sku_cat_df, price_df, fixed_fee):
    # ── Normalise charges_df column names ──────────────────────────────
    charges_df = charges_df.copy()
    charges_df.columns = [str(c).strip() for c in charges_df.iloc[0]]
    charges_df = charges_df.iloc[1:].reset_index(drop=True)
    charges_df = charges_df[charges_df['Category'].notna()]
    # Forward-fill category
    charges_df['Category'] = charges_df['Category'].ffill()
    for col in ['Lower Lim.', 'Upper Lim.', 'Charge',
                'Coll.Lower Lim.', 'Coll. Upper Lim.', 'Coll.Charge',
                'GT Lower Lim.', 'GT Upper Lim.', 'GT Charge']:
        if col in charges_df.columns:
            charges_df[col] = pd.to_numeric(charges_df[col], errors='coerce')

    # ── Normalise sku_cat_df ────────────────────────────────────────────
    sku_cat_df = sku_cat_df.copy()
    sku_cat_df.columns = [str(c).strip() for c in sku_cat_df.iloc[0]]
    sku_cat_df = sku_cat_df.iloc[1:].reset_index(drop=True)
    sku_cat_df = sku_cat_df.rename(columns={sku_cat_df.columns[0]: 'SKU', sku_cat_df.columns[1]: 'sub_cat'})
    sku_cat_dict = dict(zip(sku_cat_df['SKU'].str.upper(), sku_cat_df['sub_cat']))

    # ── Normalise price_df ──────────────────────────────────────────────
    price_df = price_df.copy()
    price_df.columns = [str(c).strip() for c in price_df.iloc[0]]
    price_df = price_df.iloc[1:].reset_index(drop=True)
    price_df['OMS Child SKU'] = price_df['OMS Child SKU'].astype(str).str.strip()
    for col in ['PWN+10%', 'PWN+10%+50']:
        if col in price_df.columns:
            price_df[col] = pd.to_numeric(price_df[col], errors='coerce')

    # ── Process orders ──────────────────────────────────────────────────
    results = []
    for _, row in order_df.iterrows():
        sku          = str(row.get('SKU', '')).strip()
        order_id     = str(row.get('Order Id', '')).strip()
        ordered_on   = row.get('Ordered On', '')
        inv_amount   = float(row.get('Invoice Amount', 0) or 0)
        quantity     = int(row.get('Quantity', 1) or 1)

        # PWN from Sheet4
        pwn10, pwn10_50 = lookup_pwn(sku, price_df)
        pwn = pwn10 if pd.notna(pwn10) else np.nan

        # Sub-category from Sheet3
        sub_cat_raw = sku_cat_dict.get(sku.upper(), '')
        sub_cat_norm = normalize_cat(sub_cat_raw)

        # Charges from Sheet2
        commission, coll_fee, gt_charge = lookup_charges(sub_cat_norm, inv_amount, charges_df)

        # Total charges = commission + collection_fee + fixed_fee + gt_charge
        if pd.notna(commission) and pd.notna(coll_fee) and pd.notna(gt_charge):
            total_charges = commission + coll_fee + fixed_fee + gt_charge
            gst_on_charges = round(total_charges * 0.18, 5)
            received_amount = round(inv_amount - total_charges - gst_on_charges, 5)
        else:
            total_charges = fixed_fee
            gst_on_charges = round(fixed_fee * 0.18, 5)
            received_amount = round(inv_amount - total_charges - gst_on_charges, 5)

        as_per_calc = round(pwn - fixed_fee, 2) if pd.notna(pwn) else np.nan
        difference  = round(received_amount - pwn, 4) if pd.notna(pwn) else np.nan

        results.append({
            'Ordered On':        ordered_on,
            'Order Id':          order_id,
            'SKU':               sku,
            'Category':          sub_cat_raw,
            'PWN':               pwn,
            'Invoice Amount':    inv_amount,
            'Quantity':          quantity,
            'Commission (₹)':    round(commission, 4) if pd.notna(commission) else '',
            'Commission %':      round(commission / inv_amount * 100, 2) if (pd.notna(commission) and inv_amount) else '',
            'Collection Fee':    round(coll_fee, 4) if pd.notna(coll_fee) else '',
            'Fixed Fee':         fixed_fee,
            'GT Charge':         gt_charge if pd.notna(gt_charge) else '',
            'Total Charges':     round(total_charges, 4),
            'GST Amount':        gst_on_charges,
            'Received Amount':   received_amount,
            'As Per Calculation': as_per_calc,
            'Difference':        difference,
        })

    return pd.DataFrame(results)


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reconciliation')
    return output.getvalue()


# ─────────────────────────────────────────────
# Main app
# ─────────────────────────────────────────────
if order_file and charges_file:
    with st.spinner("Processing files…"):
        # Read order CSV
        order_df = pd.read_csv(order_file)

        # Read charges workbook (4 sheets)
        xl = pd.read_excel(charges_file, sheet_name=None, header=None)
        sheet_names = list(xl.keys())

        # Map sheets by index (Sheet1=recon, Sheet2=charges, Sheet3=sku-cat, Sheet4=prices)
        if len(sheet_names) < 4:
            st.error("⚠️ The charges Excel must have at least 4 sheets.")
            st.stop()

        charges_raw  = xl[sheet_names[1]]   # Sheet2 – charges
        sku_cat_raw  = xl[sheet_names[2]]   # Sheet3 – SKU category
        price_raw    = xl[sheet_names[3]]   # Sheet4 – price (PWN)

    # ── Run ──────────────────────────────────────────────────────────────
    try:
        result_df = run_reconciliation(
            order_df, charges_raw, sku_cat_raw, price_raw, fixed_fee
        )
    except Exception as e:
        st.error(f"Error during reconciliation: {e}")
        st.stop()

    # ── Summary KPIs ─────────────────────────────────────────────────────
    total_orders    = len(result_df)
    total_invoice   = result_df['Invoice Amount'].sum()
    total_charges   = result_df['Total Charges'].sum()
    total_received  = result_df['Received Amount'].sum()
    matched         = result_df['Difference'].notna().sum()
    pos_diff        = (result_df['Difference'] > 0).sum()
    neg_diff        = (result_df['Difference'] < 0).sum()

    st.markdown("### 📊 Summary")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Orders",          f"{total_orders:,}")
    c2.metric("Total Invoice Amt",     f"₹{total_invoice:,.0f}")
    c3.metric("Total Charges",         f"₹{total_charges:,.0f}")
    c4.metric("Total Received (Est.)", f"₹{total_received:,.0f}")
    c5.metric("Orders with +Diff / -Diff", f"{pos_diff} / {neg_diff}")

    # ── Filters ───────────────────────────────────────────────────────────
    st.markdown("### 🔍 Filter & View")
    col1, col2, col3 = st.columns(3)
    cats = ['All'] + sorted(result_df['Category'].dropna().unique().tolist())
    sel_cat  = col1.selectbox("Category", cats)
    diff_opt = col2.selectbox("Difference", ['All', 'Positive (+)', 'Negative (−)', 'Zero', 'N/A'])
    search   = col3.text_input("Search SKU / Order ID")

    view_df = result_df.copy()
    if sel_cat != 'All':
        view_df = view_df[view_df['Category'] == sel_cat]
    if diff_opt == 'Positive (+)':
        view_df = view_df[view_df['Difference'] > 0]
    elif diff_opt == 'Negative (−)':
        view_df = view_df[view_df['Difference'] < 0]
    elif diff_opt == 'Zero':
        view_df = view_df[view_df['Difference'] == 0]
    elif diff_opt == 'N/A':
        view_df = view_df[view_df['Difference'].isna()]
    if search:
        mask = (view_df['SKU'].str.contains(search, case=False, na=False) |
                view_df['Order Id'].str.contains(search, case=False, na=False))
        view_df = view_df[mask]

    st.markdown(f"**Showing {len(view_df):,} of {total_orders:,} orders**")

    # Highlight negative difference
    def highlight_diff(val):
        if pd.isna(val) or val == '':
            return ''
        try:
            v = float(val)
            if v < 0:   return 'color: red'
            if v > 0:   return 'color: green'
        except:
            pass
        return ''

    st.dataframe(
        view_df.style.applymap(highlight_diff, subset=['Difference']),
        use_container_width=True,
        height=500
    )

    # ── Download ──────────────────────────────────────────────────────────
    st.markdown("### 📥 Download")
    col_dl1, col_dl2 = st.columns(2)
    col_dl1.download_button(
        "⬇ Download Full Reconciliation (Excel)",
        data=to_excel(result_df),
        file_name="flipkart_reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    col_dl2.download_button(
        "⬇ Download Filtered View (Excel)",
        data=to_excel(view_df),
        file_name="flipkart_reconciliation_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ── Category breakdown ────────────────────────────────────────────────
    st.markdown("### 📈 Category Breakdown")
    grp = result_df.groupby('Category').agg(
        Orders=('Order Id', 'count'),
        Invoice=('Invoice Amount', 'sum'),
        Charges=('Total Charges', 'sum'),
        Received=('Received Amount', 'sum'),
        Avg_Diff=('Difference', 'mean'),
    ).reset_index().sort_values('Invoice', ascending=False)
    grp['Invoice']  = grp['Invoice'].round(0)
    grp['Charges']  = grp['Charges'].round(0)
    grp['Received'] = grp['Received'].round(0)
    grp['Avg_Diff'] = grp['Avg_Diff'].round(2)
    st.dataframe(grp, use_container_width=True)

else:
    st.info("👈 Please upload **both** files in the sidebar to start reconciliation.")
    st.markdown("""
    #### How it works
    1. **Order CSV** – exported from Flipkart Seller Hub (columns used: `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Quantity`)
    2. **Charges Excel (4 sheets)**:
       - **Sheet 2** – Category-wise commission % and collection fee slabs
       - **Sheet 3** – Maps each SKU to its sub-category
       - **Sheet 4** – PWN (Price We Need) reference prices per SKU
    3. The tool looks up the category for each SKU, applies the correct commission, collection fee, and GT charge, then calculates `Received Amount` and compares it to the expected `PWN`.
    """)
