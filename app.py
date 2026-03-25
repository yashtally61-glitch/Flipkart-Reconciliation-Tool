import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import defaultdict
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
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
    order_file      = st.file_uploader("1️⃣  Order CSV  (Flipkart export)", type=["csv"])
    charges_file    = st.file_uploader("2️⃣  Data Excel (Ashirwad workbook)", type=["xlsx"])
    replace_sku_file = st.file_uploader("3️⃣  Replace SKU Excel (optional override)", type=["xlsx"])
    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5,  min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)",      value=18, min_value=0, step=1) / 100

    # TDS / TCS rates (fixed as per requirement)
    TDS_RATE = 0.001   # 0.1%
    TCS_RATE = 0.005   # 0.5%

    st.markdown("---")
    st.markdown("""
**Excel sheet layout (by position):**
- **Sheet 1** (index 0) – Charges Decription
- **Sheet 2** (index 1) – Category Discription
- **Sheet 3** (index 2) – Price We Need

> Categories & rates are read **live from Excel** —
> no hardcoded values.
""")
    st.markdown("""
**TDS / TCS rates (fixed):**
- TDS = 0.1%
- TCS = 0.5%
- Taxable Value = Selling Price − (Selling Price / 105 × 5)
""")

# ═══════════════════════════════════════════════════════════════════════════════
# SIZE EXPAND MAP
# ═══════════════════════════════════════════════════════════════════════════════
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
# DYNAMIC CATEGORY MAP — built at runtime from Excel
# ═══════════════════════════════════════════════════════════════════════════════

def build_cat_map(sub_cats: list, charge_cats: list) -> dict:
    charge_lower = {c.lower(): c for c in charge_cats}
    mapping = {}
    for sc in sub_cats:
        key = sc.strip().lower()
        if key in charge_lower:
            mapping[key] = charge_lower[key]
    return mapping


def get_cat_for_lookup(sub_cat_raw: str, cat_map: dict, manual_map: dict) -> str:
    if not sub_cat_raw or sub_cat_raw == "nan":
        return ""
    key = sub_cat_raw.strip().lower()
    if key in manual_map:
        return manual_map[key]
    return cat_map.get(key, "")


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def strip_vendor_prefix(sku: str) -> str:
    upper = sku.upper()
    for prefix in VENDOR_PREFIXES:
        if upper.startswith(prefix.upper()):
            return sku[len(prefix):]
    return sku


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


def lookup_pwn_with_replace(sku: str, pwn_dict: dict, replace_map: dict) -> tuple:
    """
    Try direct/size-expand lookup first.
    If not found, try replacing the Seller SKU → OMS SKU via replace_map,
    then retry the lookup on the mapped OMS SKU.
    """
    pwn_val, method = lookup_pwn(sku, pwn_dict)
    if method != "not_found":
        return pwn_val, method

    # Try replace map: Seller SKU Id → OMS SKU
    upper_sku = sku.strip().upper()
    oms_sku = replace_map.get(upper_sku)
    if oms_sku:
        pwn_val2, method2 = lookup_pwn(oms_sku, pwn_dict)
        if method2 != "not_found":
            return pwn_val2, f"replace→{method2}"

    return np.nan, "not_found"


def get_gt_amount(cat: str, inv_amount: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("GT Lower Limit")
        hi = r.get("GT Upper Limit")
        gt = r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gt):
            if float(lo) <= inv_amount <= float(hi) + 0.99:
                return float(gt)
    return np.nan


def get_commission(cat: str, sell: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("Lower Limit Commision")
        hi = r.get("Upper Limit Commision")
        ch = r.get("Commision Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell <= float(hi) + 0.99:
                return round(float(ch) * sell, 5)
    return 0.0


def get_collection_fee(cat: str, sell: float, cdf: pd.DataFrame) -> float:
    rows = cdf[cdf["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo_raw = r.get("Collection Lower Limit")
        hi     = r.get("Collection Upper Limit")
        cf     = r.get("Collection Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).strip().startswith(">")) else float(lo_raw)
        if lo_val < sell <= float(hi) + 0.99:
            return round(float(cf) * sell, 5)
    return 0.0


def calc_taxable_value(selling_price: float) -> float:
    """Taxable Value = Selling Price - (Selling Price / 105 * 5)"""
    return selling_price - (selling_price / 105 * 5)


def calc_tds(taxable_value: float, rate: float = TDS_RATE) -> float:
    return round(taxable_value * rate, 5)


def calc_tcs(taxable_value: float, rate: float = TCS_RATE) -> float:
    return round(taxable_value * rate, 5)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_charges_df(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df["Category"].notna()].copy()
    df["Category"] = df["Category"].ffill()
    numeric_cols = [
        "Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
        "Collection Lower Limit", "Collection Upper Limit", "Collection Charge",
        "GT Lower Limit", "GT Upper Limit", "GT Charge",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def parse_sku_cat(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(
        df["Seller SKU Id"].astype(str).str.strip().str.upper(),
        df["Sub-category"].astype(str).str.strip(),
    ))


def parse_pwn_dict(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"]    = pd.to_numeric(df["PWN+10%+50"], errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(), df["PWN+10%+50"]))


def parse_replace_map(file) -> dict:
    """Parse Replace_SKU.xlsx: Seller SKU Id → OMS SKU (uppercase keys)."""
    xl = pd.read_excel(file, header=None)
    df = xl.copy()
    df.columns = xl.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(
        df["Seller SKU Id"].astype(str).str.strip().str.upper(),
        df["OMS SKU"].astype(str).str.strip().str.upper(),
    ))


# ═══════════════════════════════════════════════════════════════════════════════
# CORE RECONCILIATION
# ═══════════════════════════════════════════════════════════════════════════════

def run_reconciliation(order_df, charges_df, sku_cat_dict, pwn_dict,
                       cat_map, manual_cat_map,
                       fixed_fee, gst_rate,
                       replace_map: dict = None,
                       pwn_overrides: dict = None) -> pd.DataFrame:
    pwn_overrides = pwn_overrides or {}
    replace_map   = replace_map   or {}
    rows_out = []

    for _, row in order_df.iterrows():
        raw_sku    = str(row.get("SKU", "")).strip()
        sku        = strip_vendor_prefix(raw_sku)
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)

        sub_cat_raw = sku_cat_dict.get(sku.upper(), "")
        cat         = get_cat_for_lookup(sub_cat_raw, cat_map, manual_cat_map)
        gt_val      = get_gt_amount(cat, inv_amount, charges_df) if cat else np.nan

        if pd.isna(gt_val) or not cat:
            sell_price     = np.nan
            gt_val         = np.nan
            commission     = coll_fee = total_charges = np.nan
            gst_on_charges = np.nan
            taxable_value  = np.nan
            tds            = np.nan
            tcs            = np.nan
            total_deductions = received_amount = np.nan
        else:
            sell_price       = round((inv_amount - gt_val) * quantity, 5)
            commission       = get_commission(cat, sell_price, charges_df)
            coll_fee         = get_collection_fee(cat, sell_price, charges_df)
            total_charges    = round(commission + coll_fee + float(fixed_fee), 5)
            gst_on_charges   = round(total_charges * gst_rate, 5)

            # TDS / TCS on taxable value
            taxable_value    = round(calc_taxable_value(sell_price), 5)
            tds              = calc_tds(taxable_value)
            tcs              = calc_tcs(taxable_value)

            # Received Amount = Selling Price − Total Charges − GST on Charges − TDS − TCS
            total_deductions = round(total_charges + gst_on_charges + tds + tcs, 5)
            received_amount  = round(sell_price - total_charges - gst_on_charges - tds - tcs, 5)

        # PWN lookup — with Replace map fallback
        pwn_val, match_method = lookup_pwn_with_replace(sku, pwn_dict, replace_map)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku.upper()]), "manual"

        # Difference — if Qty > 1, use Qty × PWN as the benchmark
        if pd.notna(received_amount) and pd.notna(pwn_val):
            pwn_benchmark = pwn_val * quantity  # Qty × PWN
            difference    = round(received_amount - pwn_benchmark, 5)
        else:
            pwn_benchmark = np.nan
            difference    = np.nan

        rows_out.append({
            "Order Id":         order_id,
            "SKU":              raw_sku,
            "Lookup SKU":       sku,
            "Ordered On":       ordered_on,
            "Sub-Category":     sub_cat_raw,
            "Charges Category": cat,
            "Qty":              quantity,
            "Invoice Amount":   inv_amount,
            "GT (As Per Calc)": gt_val,
            "Selling Price":    sell_price,
            "Commission":       commission,
            "Collection Fee":   coll_fee,
            "Fixed Fee":        float(fixed_fee),
            "Total Charges":    total_charges,
            "GST on Charges":   gst_on_charges,
            "Taxable Value":    taxable_value,
            "TDS":              tds,
            "TCS":              tcs,
            "Total Deductions": total_deductions,
            "Received Amount":  received_amount,
            "PWN":              pwn_val,
            "PWN Benchmark":    pwn_benchmark,
            "PWN Match":        match_method,
            "Difference":       difference,
        })

    return pd.DataFrame(rows_out)


# ═══════════════════════════════════════════════════════════════════════════════
# FORMATTING
# ═══════════════════════════════════════════════════════════════════════════════

MONEY_COLS = [
    "Invoice Amount", "GT (As Per Calc)", "Selling Price",
    "Commission", "Collection Fee", "Fixed Fee",
    "Total Charges", "GST on Charges", "Taxable Value",
    "TDS", "TCS",
    "Total Deductions", "Received Amount", "PWN", "PWN Benchmark", "Difference",
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
# STYLED EXCEL EXPORT  (beautiful ROC sheet)
# ═══════════════════════════════════════════════════════════════════════════════

def apply_roc_sheet_style(ws, df: pd.DataFrame):
    """Apply rich formatting to the Reconciliation sheet."""

    # ── Colour palette ──────────────────────────────────────────────────────
    C_HEADER_BG   = "1A3C5E"   # deep navy
    C_HEADER_FG   = "FFFFFF"   # white text
    C_ALT1        = "EAF2FB"   # light blue row
    C_ALT2        = "FFFFFF"   # white row
    C_GREEN_BG    = "D6EFDD"   # positive diff bg
    C_RED_BG      = "FDDEDE"   # negative diff bg
    C_ZERO_BG     = "FFF9E6"   # zero diff bg
    C_NAN_BG      = "F5F5F5"   # n/a row bg
    C_SECTION_BG  = "2980B9"   # section sub-header
    C_SECTION_FG  = "FFFFFF"
    C_TOTAL_BG    = "1A3C5E"
    C_TOTAL_FG    = "FFD700"   # gold
    C_BORDER      = "B0C4D8"

    thin  = Side(style="thin",   color=C_BORDER)
    thick = Side(style="medium", color="1A3C5E")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
    bdr_header = Border(left=thick, right=thick, top=thick, bottom=thick)

    # Money columns by name (for number format)
    money_names = set(MONEY_COLS)

    # ── Header row (row 1) ──────────────────────────────────────────────────
    for cell in ws[1]:
        cell.fill      = PatternFill("solid", fgColor=C_HEADER_BG)
        cell.font      = Font(bold=True, color=C_HEADER_FG, size=10, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = bdr_header
    ws.row_dimensions[1].height = 30

    # ── Column widths ───────────────────────────────────────────────────────
    col_widths = {
        "Order Id": 20, "SKU": 28, "Lookup SKU": 22,
        "Ordered On": 14, "Sub-Category": 20, "Charges Category": 18,
        "Qty": 6, "Invoice Amount": 15, "GT (As Per Calc)": 15,
        "Selling Price": 15, "Commission": 14, "Collection Fee": 15,
        "Fixed Fee": 10, "Total Charges": 15, "GST on Charges": 15,
        "Taxable Value": 14, "TDS": 10, "TCS": 10,
        "Total Deductions": 16, "Received Amount": 16,
        "PWN": 12, "PWN Benchmark": 15, "PWN Match": 14, "Difference": 14,
    }
    for i, col_name in enumerate(df.columns, start=1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = col_widths.get(col_name, 14)

    # Identify Difference column index
    diff_col_idx = None
    if "Difference" in df.columns:
        diff_col_idx = df.columns.tolist().index("Difference") + 1

    # ── Data rows ───────────────────────────────────────────────────────────
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        alt_fill = PatternFill("solid", fgColor=C_ALT1 if r_idx % 2 == 0 else C_ALT2)
        diff_val = getattr(row_data, "Difference", None)

        for c_idx, (col_name, val) in enumerate(zip(df.columns, row_data), start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value  = None if (isinstance(val, float) and np.isnan(val)) else val
            cell.border = bdr
            cell.font   = Font(size=9, name="Calibri")

            # Alternating row background
            cell.fill = alt_fill

            # Number format for money columns
            if col_name in money_names and isinstance(val, (int, float)) and not (isinstance(val, float) and np.isnan(val)):
                cell.number_format = '₹#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name == "Qty":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

            # Colour the Difference cell
            if c_idx == diff_col_idx and isinstance(val, float) and not np.isnan(val):
                if val < 0:
                    cell.fill = PatternFill("solid", fgColor=C_RED_BG)
                    cell.font = Font(color="C0392B", bold=True, size=9, name="Calibri")
                elif val > 0:
                    cell.fill = PatternFill("solid", fgColor=C_GREEN_BG)
                    cell.font = Font(color="1E8449", bold=True, size=9, name="Calibri")
                else:
                    cell.fill = PatternFill("solid", fgColor=C_ZERO_BG)
                    cell.font = Font(color="7D6608", bold=True, size=9, name="Calibri")

        ws.row_dimensions[r_idx].height = 16

    # ── Freeze top row ───────────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Auto-filter ──────────────────────────────────────────────────────────
    ws.auto_filter.ref = ws.dimensions

    # ── Totals row at the bottom ─────────────────────────────────────────────
    last_data_row = len(df) + 1
    total_row     = last_data_row + 2   # blank gap

    label_col = 1
    ws.cell(row=total_row, column=label_col).value = "TOTALS"
    ws.cell(row=total_row, column=label_col).font  = Font(bold=True, color=C_TOTAL_FG, size=10, name="Calibri")
    ws.cell(row=total_row, column=label_col).fill  = PatternFill("solid", fgColor=C_TOTAL_BG)

    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=total_row, column=c_idx)
        cell.fill   = PatternFill("solid", fgColor=C_TOTAL_BG)
        cell.font   = Font(bold=True, color=C_TOTAL_FG, size=10, name="Calibri")
        cell.border = bdr_header
        if col_name in money_names:
            col_letter = get_column_letter(c_idx)
            cell.value         = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
            cell.number_format = '₹#,##0.00'
            cell.alignment     = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[total_row].height = 22


def apply_summary_style(ws):
    """Style the summary/charges sheet."""
    C_H = "2C3E50"; C_FG = "FFFFFF"
    C_ODD = "EBF5FB"; C_EVEN = "FFFFFF"
    thin = Side(style="thin", color="AED6F1")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill  = PatternFill("solid", fgColor=C_H)
        cell.font  = Font(bold=True, color=C_FG, size=10, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = bdr
    ws.row_dimensions[1].height = 24

    for r_idx in range(2, ws.max_row + 1):
        fill = PatternFill("solid", fgColor=C_ODD if r_idx % 2 == 0 else C_EVEN)
        for cell in ws[r_idx]:
            cell.fill   = fill
            cell.font   = Font(size=9, name="Calibri")
            cell.border = bdr
            cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[r_idx].height = 15

    for col_cells in ws.columns:
        width = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(width + 4, 40)

    ws.freeze_panes = "A2"


def to_excel(recon_df, summary_df, cat_df) -> bytes:
    buf = BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Write DataFrames first (plain)
        recon_df.to_excel(writer,   index=False, sheet_name="Reconciliation")
        cat_df.to_excel(writer,     index=False, sheet_name="Category Breakdown")
        summary_df.to_excel(writer, index=False, sheet_name="Charges Summary")

        # Now apply rich styling
        apply_roc_sheet_style(writer.sheets["Reconciliation"], recon_df)
        apply_summary_style(writer.sheets["Category Breakdown"])
        apply_summary_style(writer.sheets["Charges Summary"])

    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def build_summary(df: pd.DataFrame) -> tuple:
    valid = df[df["Received Amount"].notna()]
    totals = {"Metric": [], "Value": []}
    fields = [
        ("Total Orders",              len(df)),
        ("Orders Calculated",         int(df["Received Amount"].notna().sum())),
        ("Orders NaN (no category)",  int(df["Received Amount"].isna().sum())),
        ("Total Invoice Amount",      df["Invoice Amount"].sum()),
        ("Total GT (As Per Calc)",    valid["GT (As Per Calc)"].sum()),
        ("Total Selling Price",       valid["Selling Price"].sum()),
        ("Total Commission",          valid["Commission"].sum()),
        ("Total Collection Fee",      valid["Collection Fee"].sum()),
        ("Total Fixed Fee",           valid["Fixed Fee"].sum()),
        ("Total Charges (C+F+Fixed)", valid["Total Charges"].sum()),
        ("Total GST on Charges",      valid["GST on Charges"].sum()),
        ("Total TDS",                 valid["TDS"].sum()),
        ("Total TCS",                 valid["TCS"].sum()),
        ("Total Deductions",          valid["Total Deductions"].sum()),
        ("Total Received Amount",     valid["Received Amount"].sum()),
        ("Net Difference vs PWN",     valid["Difference"].sum()),
        ("Orders with -ve Diff",      int((valid["Difference"] < 0).sum())),
        ("Orders with +ve Diff",      int((valid["Difference"] > 0).sum())),
        ("Orders – No PWN found",     int(df["Difference"].isna().sum())),
        ("Avg Received per Order",    valid["Received Amount"].mean()),
        ("Avg Difference per Order",  valid["Difference"].mean()),
    ]
    for label, val in fields:
        totals["Metric"].append(label)
        totals["Value"].append(round(val, 2) if isinstance(val, float) else val)
    summary_df = pd.DataFrame(totals)

    cat_df = (
        valid.groupby("Sub-Category")
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
            TDS_Total      = ("TDS",               "sum"),
            TCS_Total      = ("TCS",               "sum"),
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
        "Sub-Category", "Orders", "Invoice Total", "GT Total", "Selling Total",
        "Commission", "Collection Fee", "Fixed Fee",
        "Total Charges", "GST Total", "TDS Total", "TCS Total",
        "Total Deductions",
        "Received Total", "Net Difference", "Avg Difference",
    ]
    return summary_df, cat_df


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k, v in [
    ("pwn_overrides",  {}),
    ("manual_cat_map", {}),
    ("result_df",      None),
    ("charges_df",     None),
    ("sku_cat_dict",   None),
    ("pwn_dict",       None),
    ("order_df",       None),
    ("cat_map",        {}),
    ("charge_cats",    []),
    ("unmapped_cats",  []),
    ("replace_map",    {}),
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
            st.error(f"❌ Excel must have at least 3 sheets. Found: {list(xl.keys())}")
            st.stop()

        charges_df   = parse_charges_df(sheets[0])
        sku_cat_dict = parse_sku_cat(sheets[1])
        pwn_dict     = parse_pwn_dict(sheets[2])

        # Replace SKU map
        replace_map = {}
        if replace_sku_file:
            replace_map = parse_replace_map(replace_sku_file)
            st.sidebar.success(f"✅ Replace SKU loaded: {len(replace_map):,} entries")

        charge_cats  = charges_df["Category"].unique().tolist()
        all_sub_cats = sorted(set(v for v in sku_cat_dict.values() if v and v != "nan"))
        cat_map      = build_cat_map(all_sub_cats, charge_cats)
        unmapped     = [sc for sc in all_sub_cats if sc.lower() not in cat_map]

        st.session_state.update({
            "charges_df":    charges_df,
            "sku_cat_dict":  sku_cat_dict,
            "pwn_dict":      pwn_dict,
            "order_df":      order_df,
            "cat_map":       cat_map,
            "charge_cats":   charge_cats,
            "unmapped_cats": unmapped,
            "replace_map":   replace_map,
        })

    # ── Unmapped category resolver ──────────────────────────────────────────
    unmapped       = st.session_state["unmapped_cats"]
    manual_cat_map = st.session_state["manual_cat_map"]

    if unmapped:
        with st.expander(
            f"⚠️  **{len(unmapped)} sub-category(s) couldn't be auto-matched — assign manually**",
            expanded=True,
        ):
            st.info(
                "These sub-categories from **Category Discription** have no matching "
                "entry in **Charges Decription**. Pick the correct charges category for each."
            )
            charge_options = ["— skip / leave as NaN —"] + sorted(st.session_state["charge_cats"])
            new_manual = {}
            cols = st.columns(2)
            for i, sc in enumerate(unmapped):
                col = cols[i % 2]
                existing = manual_cat_map.get(sc.lower(), "— skip / leave as NaN —")
                chosen = col.selectbox(
                    f"Sub-cat: **{sc}**",
                    charge_options,
                    index=charge_options.index(existing) if existing in charge_options else 0,
                    key=f"manual_cat_{sc}",
                )
                if chosen != "— skip / leave as NaN —":
                    new_manual[sc.lower()] = chosen

            if st.button("💾  Save Category Mapping & Recalculate", type="primary"):
                st.session_state["manual_cat_map"] = new_manual
                st.rerun()

    result_df = run_reconciliation(
        st.session_state["order_df"],
        st.session_state["charges_df"],
        st.session_state["sku_cat_dict"],
        st.session_state["pwn_dict"],
        st.session_state["cat_map"],
        st.session_state["manual_cat_map"],
        fixed_fee, gst_rate,
        replace_map=st.session_state["replace_map"],
        pwn_overrides=st.session_state["pwn_overrides"],
    )
    st.session_state["result_df"] = result_df
    summary_df, cat_df = build_summary(result_df)

    # Count how many were resolved by replace map
    replace_resolved = result_df[result_df["PWN Match"].str.startswith("replace", na=False)]

    st.success(
        f"✅ Processed **{len(result_df):,}** orders  |  "
        f"**{int(result_df['Received Amount'].notna().sum()):,}** calculated  |  "
        f"**{int(result_df['Received Amount'].isna().sum()):,}** skipped (no category/GT match)"
        + (f"  |  **{len(replace_resolved):,}** PWN found via Replace SKU map" if len(replace_resolved) else "")
    )

    # ── Live category map viewer ────────────────────────────────────────────
    with st.expander("🗺️  View live category mapping", expanded=False):
        map_rows = []
        for sc, cc in sorted(st.session_state["cat_map"].items()):
            map_rows.append({"Sub-Category": sc, "→ Charges Category": cc, "Source": "✅ auto-matched"})
        for sc, cc in st.session_state["manual_cat_map"].items():
            map_rows.append({"Sub-Category": sc, "→ Charges Category": cc, "Source": "🖊️ manual"})
        for sc in st.session_state["unmapped_cats"]:
            if sc.lower() not in st.session_state["manual_cat_map"]:
                map_rows.append({"Sub-Category": sc, "→ Charges Category": "⚠️ NOT MAPPED", "Source": "—"})
        st.dataframe(pd.DataFrame(map_rows), use_container_width=True, hide_index=True)

    tab1, tab2, tab3 = st.tabs([
        "📋  Reconciliation",
        "💰  Charges Summary",
        "📊  Category Breakdown",
    ])

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 1 – RECONCILIATION                                         ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab1:

        missing_df = result_df[result_df["PWN Match"] == "not_found"]
        if len(missing_df):
            with st.expander(
                f"⚠️  **{len(missing_df)} SKU(s) have no PWN price — click to enter manually**",
                expanded=False,
            ):
                st.info("Enter PWN value for each SKU and click Save & Recalculate.")
                missing_skus    = missing_df["SKU"].unique().tolist()
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
                    st.rerun()

        st.markdown("### 📊 Summary")
        valid = result_df[result_df["Received Amount"].notna()]
        k1,k2,k3,k4,k5,k6,k7,k8,k9,k10 = st.columns(10)
        k1.metric("Orders",            f"{len(result_df):,}")
        k2.metric("Invoice Total",     f"₹{result_df['Invoice Amount'].sum():,.0f}")
        k3.metric("GT Total (ref)",    f"₹{valid['GT (As Per Calc)'].sum():,.0f}")
        k4.metric("Selling Pr. Total", f"₹{valid['Selling Price'].sum():,.0f}")
        k5.metric("Total Charges",     f"₹{valid['Total Charges'].sum():,.0f}")
        k6.metric("GST on Charges",    f"₹{valid['GST on Charges'].sum():,.0f}")
        k7.metric("Total TDS",         f"₹{valid['TDS'].sum():,.2f}")
        k8.metric("Total TCS",         f"₹{valid['TCS'].sum():,.2f}")
        k9.metric("Received Total",    f"₹{valid['Received Amount'].sum():,.0f}")
        net = valid["Difference"].sum()
        k10.metric(
            "Net Difference",
            f"₹{net:,.2f}",
            delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
            delta_color="normal" if net >= 0 else "inverse",
        )

        st.markdown("---")

        f1, f2, f3 = st.columns([2, 2, 3])
        all_cats_opt = ["All"] + sorted(result_df["Sub-Category"].dropna().unique().tolist())
        sel_cat  = f1.selectbox("Sub-Category", all_cats_opt)
        diff_opt = f2.selectbox("Difference type",
                                ["All", "Positive (+)", "Negative (−)",
                                 "Zero / Matched", "No PWN data", "No Category (NaN)"])
        search   = f3.text_input("🔎 Search by SKU or Order ID")

        view = result_df.copy()
        if sel_cat != "All":
            view = view[view["Sub-Category"] == sel_cat]
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
            "Order Id", "SKU", "Lookup SKU", "Ordered On",
            "Sub-Category", "Charges Category",
            "Qty", "Invoice Amount",
            "GT (As Per Calc)", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "Taxable Value", "TDS", "TCS",
            "Total Deductions", "Received Amount",
            "PWN", "PWN Benchmark", "Difference", "PWN Match",
        ]

        st.dataframe(
            style_table(view[display_cols], diff_col="Difference"),
            use_container_width=True,
            height=500,
        )

        st.markdown("### 📥 Download")
        d1, d2 = st.columns(2)
        d1.download_button(
            "⬇  Full Reconciliation (Excel – 3 sheets, styled)",
            data=to_excel(result_df[display_cols], summary_df, cat_df),
            file_name="flipkart_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        d2.download_button(
            "⬇  Filtered View (Excel, styled)",
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
            a1.metric("TDS (0.1%)",     f"₹{valid['TDS'].sum():,.2f}")
            a2.metric("TCS (0.5%)",     f"₹{valid['TCS'].sum():,.2f}")
            st.metric("🔴 Total Deductions", f"₹{valid['Total Deductions'].sum():,.2f}")

        with col_b:
            st.markdown("#### 📥 What You Receive")
            b1, b2 = st.columns(2)
            b1.metric("Total Invoice",  f"₹{result_df['Invoice Amount'].sum():,.2f}")
            b2.metric("GT Total (ref)", f"₹{valid['GT (As Per Calc)'].sum():,.2f}")
            b1.metric("Selling Total",  f"₹{valid['Selling Price'].sum():,.2f}")
            b2.metric("Total Received", f"₹{valid['Received Amount'].sum():,.2f}")
            net = valid["Difference"].sum()
            b1.metric(
                "Net Diff vs PWN",
                f"₹{net:,.2f}",
                delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
                delta_color="normal" if net >= 0 else "inverse",
            )
            b2.metric("Orders –ve Diff", int((valid["Difference"] < 0).sum()))

        st.info(
            "ℹ️  **Received Amount** = Selling Price − Total Charges − GST on Charges − TDS − TCS  \n"
            "**Taxable Value** = Selling Price − (Selling Price / 105 × 5)  \n"
            "**TDS** = Taxable Value × 0.1%  |  **TCS** = Taxable Value × 0.5%  \n"
            "**Difference** = Received Amount − (Qty × PWN)"
        )

        st.markdown("---")
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols = [
            "Order Id", "SKU", "Sub-Category", "Charges Category",
            "Invoice Amount", "GT (As Per Calc)", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges",
            "Taxable Value", "TDS", "TCS",
            "Total Deductions", "Received Amount",
        ]
        st.dataframe(style_table(result_df[charge_cols]), use_container_width=True, height=480)

        st.markdown("---")
        st.markdown("#### 🧾 Grand Summary Table")
        st.dataframe(summary_df, use_container_width=True)

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 3 – CATEGORY BREAKDOWN                                     ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab3:
        st.markdown("### 📊 Sub-Category-wise Breakdown")
        st.caption("Every charge component summed per sub-category (NaN rows excluded)")

        cat_money = [c for c in cat_df.columns if c not in ("Sub-Category", "Orders")]
        st.dataframe(
            style_table(cat_df, diff_col="Net Difference")
            .format({c: "₹{:.2f}" for c in cat_money}),
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("#### 🔢 Charge Components Only (per Sub-Category)")
        comp_cols = [
            "Sub-Category", "Orders",
            "GT Total", "Commission", "Collection Fee", "Fixed Fee",
            "GST Total", "TDS Total", "TCS Total", "Total Deductions",
        ]
        comp_money = [c for c in comp_cols if c not in ("Sub-Category", "Orders")]
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
| **Order CSV** | Flipkart Seller Hub export — needs: `Order Id`, `SKU`, `Ordered On`, `Invoice Amount`, `Quantity` |
| **Data Excel** | Ashirwad workbook — 3 sheets by position (see below) |
| **Replace SKU Excel** *(optional)* | Maps Seller SKU Id → OMS SKU for PWN fallback lookup |

**Excel sheet positions:**

| Position | Sheet Name | Used For |
|----------|-----------|----------|
| Index 0 | Charges Decription | Commission / Collection / GT slabs |
| Index 1 | Category Discription | Seller SKU → Sub-category |
| Index 2 | Price We Need | OMS Child SKU → PWN price |

---
### ✨ What's new in this version

| # | Feature |
|---|---------|
| 1 | **Replace SKU fallback** — if PWN not found by SKU, look up via Replace SKU sheet (Seller SKU → OMS SKU) and retry |
| 2 | **TDS & TCS deducted** — Taxable Value = SP − SP/105×5; TDS = 0.1%, TCS = 0.5% |
| 3 | **Qty-aware Difference** — Difference = Received Amount − (Qty × PWN) |
| 4 | **Beautiful styled Excel** — navy headers, alternating rows, red/green diff colouring, total row |

---
### Calculation per order

```
GT Amount        = Fixed ₹ from GT slab     (Invoice Amount → slab lookup)
Selling Price    = (Invoice Amount − GT Amount) × Qty

Commission       = Selling Price × Commission %   (slab by Selling Price)
Collection Fee   = Selling Price × Collection %   (slab by Selling Price)
Total Charges    = Commission + Collection Fee + Fixed Fee

GST              = Total Charges × 18%

Taxable Value    = Selling Price − (Selling Price / 105 × 5)
TDS              = Taxable Value × 0.1%
TCS              = Taxable Value × 0.5%

Received Amount  = Selling Price − Total Charges − GST − TDS − TCS

PWN Benchmark    = PWN × Qty   (if Qty > 1)
Difference       = Received Amount − PWN Benchmark
```
""")
