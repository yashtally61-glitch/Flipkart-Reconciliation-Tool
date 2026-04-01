import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import defaultdict
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
)
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Flipkart Reconciliation – Yash Gallery Private Limited",
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
st.caption("Yash Gallery Private Limited — Tool made by Ashu Bhatt | Finance Team")

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📂 Upload Files")
    order_files      = st.file_uploader(
        "1️⃣  Order File(s)  (CSV / XLSX / XLS — multiple allowed)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )
    charges_file     = st.file_uploader("2️⃣  Data Excel", type=["xlsx"])
    replace_sku_file = st.file_uploader("3️⃣  Replace SKU Excel (optional override)", type=["xlsx"])
    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5,  min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)",      value=18, min_value=0, step=1) / 100

    TDS_RATE = 0.001
    TCS_RATE = 0.005

    st.markdown("---")
    st.markdown("""
**Excel sheet layout (by position):**
- **Sheet 0** — Charges Rates (Brand + Category → Commission / Collection / GT slabs)
- **Sheet 1** — Category Description (SKU + Brand + Sub-category)
- **Sheet 2** — Price We Need (PWN prices)

> **Brand** is read from **Sheet 1** (stored per SKU).  
> Charges use **Brand + Sub-category** from Sheet 0.  
> GT, Commission, Collection are looked up **independently**.
""")
    st.markdown("""
**TDS / TCS rates (fixed):**
- TDS = 0.1%  |  TCS = 0.5%
- Taxable Value = Selling Price − (Selling Price / 105 × 5)
""")

# ═══════════════════════════════════════════════════════════════════════════════
# SIZE EXPAND MAP
# ═══════════════════════════════════════════════════════════════════════════════
SIZE_EXPAND = {
    "L-XL": ["L", "XL"], "S-M": ["S", "M"], "XXL-3XL": ["XXL", "3XL"],
    "F-S/XXL": ["F"], "F-3xl/5xl": ["F"], "XS-S": ["XS", "S"],
    "M-L": ["M", "L"], "XL-XXL": ["XL", "XXL"], "3XL-4XL": ["3XL", "4XL"],
    "5XL-6XL": ["5XL", "6XL"], "7XL-8XL": ["7XL", "8XL"],
    "4XL-5XL": ["4XL", "5XL"], "2XL-3XL": ["2XL", "3XL"],
    "XS-S-M": ["XS", "S", "M"], "L-XL-XXL": ["L", "XL", "XXL"],
}
ORDER_TO_PRICE_SIZE: dict = defaultdict(list)
for _ps, _os_list in SIZE_EXPAND.items():
    for _os in _os_list:
        ORDER_TO_PRICE_SIZE[_os.upper()].append(_ps)

VENDOR_PREFIXES = ["GWN-", "GWN_", "GWN", "SPF-", "SPF_", "SPF", "KL_", "KL-", "KL"]


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def strip_vendor_prefix(sku: str) -> str:
    import re
    upper = sku.upper()
    for prefix in VENDOR_PREFIXES:
        if upper.startswith(prefix.upper()):
            sku = sku[len(prefix):]
            break
    sku = re.sub(r"(?i)YKN", "YK", sku)
    sku = re.sub(r"(?i)YKC", "YK", sku)
    return sku


def get_sku_base(sku: str) -> str:
    import re
    key = re.sub(r"\s*-\s*", "-", sku.strip().upper())
    parts = key.rsplit("-", 1)
    return parts[0] if len(parts) == 2 else key


def lookup_cat_by_base(sku: str, sku_info_dict: dict) -> tuple:
    """Find sub-category + brand by base SKU code if exact match fails."""
    base = get_sku_base(sku)
    if not base:
        return "", "", ""
    for candidate_sku, info in sku_info_dict.items():
        if get_sku_base(candidate_sku) == base and info.get("sub_cat") and info["sub_cat"] != "nan":
            return info["sub_cat"], info.get("brand", ""), candidate_sku
    return "", "", ""


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
                return float(val), f"size-expand({price_size})"
    base = get_sku_base(key)
    if base:
        for candidate, cval in pwn_dict.items():
            if get_sku_base(candidate) == base and pd.notna(cval):
                return float(cval), f"base-match({candidate})"
    return np.nan, "not_found"


def lookup_pwn_with_replace(sku: str, pwn_dict: dict, replace_map: dict) -> tuple:
    pwn_val, method = lookup_pwn(sku, pwn_dict)
    if method != "not_found":
        return pwn_val, method
    oms_sku = replace_map.get(sku.strip().upper())
    if oms_sku:
        pwn_val2, method2 = lookup_pwn(oms_sku, pwn_dict)
        if method2 != "not_found":
            return pwn_val2, f"replace→{method2}"
    return np.nan, "not_found"


# ═══════════════════════════════════════════════════════════════════════════════
# INDEPENDENT SLAB LOOKUPS
# Sheet 0 has Brand Name + Category + slab rows.
# Commission, Collection, and GT each occupy their OWN rows in the same block.
# They MUST be looked up independently — not expected to be on the same row.
#
# GT        → looked up against Invoice Amount
# Commission → looked up against Selling Price (Invoice − GT)
# Collection → looked up against Selling Price
# ═══════════════════════════════════════════════════════════════════════════════

def _filter_brand_cat(charges_df: pd.DataFrame, brand: str, cat: str) -> pd.DataFrame:
    """Filter charges by brand and category with better error handling"""
    if not brand or not cat:
        return pd.DataFrame()
    
    # Ensure columns exist
    if "Brand Name" not in charges_df.columns or "Category" not in charges_df.columns:
        return pd.DataFrame()
    
    # Normalize inputs - handle nan values
    brand_norm = str(brand).strip().lower()
    cat_norm = str(cat).strip().lower()
    
    # Skip if normalized values are 'nan'
    if brand_norm == 'nan' or cat_norm == 'nan':
        return pd.DataFrame()
    
    # Create normalized columns for matching
    brand_mask = charges_df["Brand Name"].fillna("").astype(str).str.strip().str.lower() == brand_norm
    cat_mask = charges_df["Category"].fillna("").astype(str).str.strip().str.lower() == cat_norm
    
    return charges_df[brand_mask & cat_mask].copy()


def debug_charge_lookup(brand: str, cat: str, charges_df: pd.DataFrame) -> dict:
    """Debug helper to see what's being matched"""
    info = {
        "brand_input": brand,
        "cat_input": cat,
        "brands_in_sheet": charges_df["Brand Name"].dropna().unique().tolist() if "Brand Name" in charges_df.columns else [],
        "categories_in_sheet": charges_df["Category"].dropna().unique().tolist() if "Category" in charges_df.columns else [],
        "exact_match_found": False,
        "matched_rows": 0
    }
    
    if brand and cat:
        matched = _filter_brand_cat(charges_df, brand, cat)
        info["exact_match_found"] = len(matched) > 0
        info["matched_rows"] = len(matched)
        
        # Try to find close matches
        if len(matched) == 0:
            close_brands = [b for b in info["brands_in_sheet"] 
                          if b and brand and str(b).lower().strip() == str(brand).lower().strip()]
            close_cats = [c for c in info["categories_in_sheet"] 
                        if c and cat and str(c).lower().strip() == str(cat).lower().strip()]
            info["close_brand_matches"] = close_brands
            info["close_cat_matches"] = close_cats
    
    return info


def lookup_gt(brand: str, cat: str, inv_amount: float, charges_df: pd.DataFrame) -> float:
    """Scan all rows for brand+cat; find GT Lower Limit <= inv_amount <= GT Upper Limit."""
    rows = _filter_brand_cat(charges_df, brand, cat)
    for _, r in rows.iterrows():
        lo = r.get("GT Lower Limit")
        hi = r.get("GT Upper Limit")
        gt = r.get("GT Charge")
        if pd.isna(lo) or pd.isna(hi) or pd.isna(gt):
            continue
        try:
            if float(lo) <= inv_amount <= float(hi) + 0.99:
                return float(gt)
        except (ValueError, TypeError):
            continue
    return np.nan


def lookup_commission(brand: str, cat: str, sell_price: float, charges_df: pd.DataFrame) -> float:
    """Scan all rows for brand+cat; find commission slab for sell_price."""
    rows = _filter_brand_cat(charges_df, brand, cat)
    for _, r in rows.iterrows():
        lo = r.get("Lower Limit Commision")
        hi = r.get("Upper Limit Commision")
        ch = r.get("Commision Charge")
        if pd.isna(lo) or pd.isna(hi) or pd.isna(ch):
            continue
        try:
            if float(lo) <= sell_price <= float(hi) + 0.99:
                return round(float(ch) * sell_price, 5)
        except (ValueError, TypeError):
            continue
    return np.nan


def lookup_collection(brand: str, cat: str, sell_price: float, charges_df: pd.DataFrame) -> float:
    """Scan all rows for brand+cat; find collection slab for sell_price."""
    rows = _filter_brand_cat(charges_df, brand, cat)
    for _, r in rows.iterrows():
        lo_raw = r.get("Collection Lower Limit")
        hi     = r.get("Collection Upper Limit")
        cf     = r.get("Collection Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        try:
            cf_val = float(cf) if pd.notna(cf) else 0.0
            lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).strip().startswith(">")) else float(lo_raw)
            if lo_val < sell_price <= float(hi) + 0.99:
                return round(cf_val * sell_price, 5)
        except (ValueError, TypeError):
            continue
    return np.nan


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_charges_df(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 0: Brand Name | Category | commission slabs | collection slabs | GT slabs.
    Brand Name and Category are forward-filled so every slab row carries them.
    Commission, Collection, GT are on INDEPENDENT rows — each must be scanned separately.
    """
    df = raw.copy()
    df.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)

    if "Brand Name" in df.columns:
        df["Brand Name"] = df["Brand Name"].ffill()
    if "Category" in df.columns:
        df["Category"] = df["Category"].ffill()

    df = df[df["Category"].notna()].copy()

    # Numeric columns — convert; Collection Charge may contain '₹0'
    for col in ["Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
                "Collection Lower Limit", "Collection Upper Limit",
                "GT Lower Limit", "GT Upper Limit", "GT Charge"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "Collection Charge" in df.columns:
        df["Collection Charge"] = (
            df["Collection Charge"].astype(str)
            .str.replace("₹", "", regex=False).str.strip()
            .pipe(pd.to_numeric, errors="coerce")
        )

    return df


def parse_sku_info(raw: pd.DataFrame) -> dict:
    """
    Sheet 1: Brand | Seller SKU Id | Sub-category
    Returns: UPPER_SKU → {"sub_cat": ..., "brand": ...}
    Brand from here is the authoritative source — not guessed from Product name.
    """
    df = raw.copy()
    df.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)

    sku_info = {}
    for _, row in df.iterrows():
        sku     = str(row.get("Seller SKU Id", "")).strip().upper()
        sub_cat = str(row.get("Sub-category", "")).strip()
        brand   = str(row.get("Brand", "")).strip() if "Brand" in df.columns else ""
        if sku:
            sku_info[sku] = {"sub_cat": sub_cat, "brand": brand}
    return sku_info


def parse_pwn_dict(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"]    = pd.to_numeric(df["PWN+10%+50"], errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(), df["PWN+10%+50"]))


def parse_replace_map(file) -> dict:
    xl = pd.read_excel(file, header=None)
    df = xl.copy()
    df.columns = [str(c).strip() for c in xl.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(
        df["Seller SKU Id"].astype(str).str.strip().str.upper(),
        df["OMS SKU"].astype(str).str.strip().str.upper(),
    ))


# ═══════════════════════════════════════════════════════════════════════════════
# MULTI-FILE ORDER READER
# ═══════════════════════════════════════════════════════════════════════════════

REQUIRED_ORDER_COLS = {"Order Id", "SKU", "Invoice Amount", "Quantity", "Product"}


def read_order_file(f) -> tuple:
    name = f.name.lower()
    try:
        if name.endswith(".csv"):
            raw = f.read()
            for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(BytesIO(raw), encoding=enc); break
                except UnicodeDecodeError:
                    continue
            else:
                return pd.DataFrame(), f"Could not decode '{f.name}'."
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
            df = pd.read_excel(f, engine=engine)
        else:
            return pd.DataFrame(), f"Unsupported file type: '{f.name}'"

        df.columns = [str(c).strip() for c in df.columns]
        missing = REQUIRED_ORDER_COLS - set(df.columns)
        if missing:
            return pd.DataFrame(), f"'{f.name}' missing columns: {', '.join(sorted(missing))}"
        df["_source_file"] = f.name
        return df, ""
    except Exception as e:
        return pd.DataFrame(), f"Error reading '{f.name}': {e}"


def load_all_order_files(files) -> tuple:
    frames, file_info, errors = [], [], []
    for f in files:
        df, err = read_order_file(f)
        if err:
            errors.append(err)
            file_info.append({"File": f.name, "Rows": 0, "Status": "❌ Error"})
        else:
            frames.append(df)
            file_info.append({"File": f.name, "Rows": len(df), "Status": "✅ OK"})
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, file_info, errors


# ═══════════════════════════════════════════════════════════════════════════════
# CORE RECONCILIATION
# ═══════════════════════════════════════════════════════════════════════════════

def run_reconciliation(order_df, charges_df, sku_info_dict, pwn_dict,
                       fixed_fee, gst_rate,
                       replace_map:     dict = None,
                       pwn_overrides:   dict = None,
                       sku_corrections: dict = None) -> pd.DataFrame:
    replace_map     = replace_map     or {}
    pwn_overrides   = pwn_overrides   or {}
    sku_corrections = sku_corrections or {}
    rows_out = []

    for _, row in order_df.iterrows():
        raw_sku    = str(row.get("SKU", "")).strip()
        product    = str(row.get("Product", "")).strip()
        order_id   = str(row.get("Order Id", "")).strip()
        ordered_on = row.get("Ordered On", "")
        inv_amount = float(row.get("Invoice Amount", 0) or 0)
        quantity   = int(row.get("Quantity", 1) or 1)

        # STEP 1: Apply manual SKU correction FIRST (from UI corrections)
        corrected_raw = sku_corrections.get(raw_sku.upper(), raw_sku)
        
        # STEP 2: Apply Replace SKU mapping (from Replace SKU Excel)
        lookup_sku = corrected_raw.strip().upper()
        if lookup_sku in replace_map:
            corrected_raw = replace_map[lookup_sku]
            
        # STEP 3: Strip vendor prefix
        sku = strip_vendor_prefix(corrected_raw)

        # ── Brand + Sub-category from Sheet 1 ────────────────────────────
        info = sku_info_dict.get(sku.upper(), {})
        sub_cat_raw    = info.get("sub_cat", "")
        brand_name     = info.get("brand", "")
        cat_match_note = ""

        # Base-code fallback
        if not sub_cat_raw or str(sub_cat_raw).lower() == "nan":
            fb_sub, fb_brand, fb_sku = lookup_cat_by_base(sku, sku_info_dict)
            if fb_sub:
                sub_cat_raw    = fb_sub
                brand_name     = fb_brand or brand_name
                cat_match_note = f"base-cat({fb_sku})"

        # Normalize the category (remove "nan" string)
        cat = sub_cat_raw.strip() if sub_cat_raw and str(sub_cat_raw).lower() != "nan" else ""
        
        # Normalize brand (remove empty/nan)
        brand_name = brand_name.strip() if brand_name and str(brand_name).lower() != "nan" else ""

        # ── Independent slab lookups ──────────────────────────────────────
        # Step 1: GT from Invoice Amount slab
        # Step 2: Selling Price = Invoice − GT
        # Step 3: Commission from Selling Price slab (independent rows)
        # Step 4: Collection from Selling Price slab (independent rows)

        gt_val     = np.nan
        sell_price = np.nan
        commission = np.nan
        coll_fee   = np.nan
        charge_method = "not_found"

        if brand_name and cat:
            gt_val = lookup_gt(brand_name, cat, inv_amount, charges_df)

            if pd.notna(gt_val):
                sell_price = round(inv_amount - gt_val, 5)
                commission = lookup_commission(brand_name, cat, sell_price, charges_df)
                coll_fee   = lookup_collection(brand_name, cat, sell_price, charges_df)

                if pd.notna(commission) and pd.notna(coll_fee):
                    charge_method = f"{brand_name} | {cat}"
                else:
                    # partial — reset
                    gt_val = sell_price = commission = coll_fee = np.nan

        # ── Final amounts ─────────────────────────────────────────────────
        if pd.isna(gt_val):
            sell_price = gt_val = commission = coll_fee = np.nan
            total_charges = gst_on_charges = taxable_value = np.nan
            tds = tcs = total_deductions = received_amount = np.nan
            charge_method = "not_found"
        else:
            commission    = commission if pd.notna(commission) else 0.0
            coll_fee      = coll_fee   if pd.notna(coll_fee)   else 0.0
            total_charges  = round(commission + coll_fee + float(fixed_fee), 5)
            gst_on_charges = round(total_charges * gst_rate, 5)
            taxable_value  = round(sell_price - (sell_price / 105 * 5), 5)
            tds            = round(taxable_value * TDS_RATE, 5)
            tcs            = round(taxable_value * TCS_RATE, 5)
            total_deductions = round(total_charges + gst_on_charges + tds + tcs, 5)
            received_amount  = round(sell_price - total_charges - gst_on_charges - tds - tcs, 5)

        # ── PWN lookup ────────────────────────────────────────────────────
        pwn_val, match_method = lookup_pwn_with_replace(sku, pwn_dict, replace_map)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku.upper()]), "manual"

        full_match_note = match_method
        if cat_match_note:
            full_match_note = f"{match_method} | cat:{cat_match_note}"

        # ── Difference ────────────────────────────────────────────────────
        if pd.notna(received_amount) and pd.notna(pwn_val):
            pwn_benchmark = round(pwn_val * quantity, 5)
            difference    = round(received_amount - pwn_benchmark, 5)
        else:
            pwn_benchmark = np.nan
            difference    = np.nan

        rows_out.append({
            "Order Id":         order_id,
            "SKU":              raw_sku,
            "Product":          product,
            "Brand Name":       brand_name,
            "Ordered On":       ordered_on,
            "Sub-Category":     sub_cat_raw,
            "Charge Method":    charge_method,
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
            "PWN Match":        full_match_note,
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
    "TDS", "TCS", "Total Deductions", "Received Amount",
    "PWN", "PWN Benchmark", "Difference",
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
        # Fixed: Changed from applymap to map for pandas compatibility
        styler = styler.map(colour_diff, subset=[diff_col])
    return styler


# ═══════════════════════════════════════════════════════════════════════════════
# STYLED EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def apply_roc_sheet_style(ws, df: pd.DataFrame):
    C_HEADER_BG = "1A3C5E"; C_HEADER_FG = "FFFFFF"
    C_ALT1 = "EAF2FB";      C_ALT2  = "FFFFFF"
    C_GREEN_BG = "D6EFDD";  C_RED_BG = "FDDEDE"; C_ZERO_BG = "FFF9E6"
    C_TOTAL_BG = "1A3C5E";  C_TOTAL_FG = "FFD700"; C_BORDER = "B0C4D8"
    thin  = Side(style="thin",   color=C_BORDER)
    thick = Side(style="medium", color="1A3C5E")
    bdr        = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
    bdr_header = Border(left=thick, right=thick, top=thick, bottom=thick)

    money_names = set(MONEY_COLS)
    cols = df.columns.tolist()
    C    = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}

    col_widths = {
        "Order Id": 20, "SKU": 28, "Product": 40, "Brand Name": 18,
        "Ordered On": 14, "Sub-Category": 20, "Charge Method": 28,
        "Qty": 6, "Invoice Amount": 15, "GT (As Per Calc)": 15,
        "Selling Price": 15, "Commission": 14, "Collection Fee": 15,
        "Fixed Fee": 10, "Total Charges": 15, "GST on Charges": 15,
        "Taxable Value": 14, "TDS": 10, "TCS": 10,
        "Total Deductions": 16, "Received Amount": 16,
        "PWN": 12, "PWN Benchmark": 15, "PWN Match": 16, "Difference": 14,
    }
    for i, col_name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = col_widths.get(col_name, 14)

    for cell in ws[1]:
        cell.fill      = PatternFill("solid", fgColor=C_HEADER_BG)
        cell.font      = Font(bold=True, color=C_HEADER_FG, size=10, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = bdr_header
    ws.row_dimensions[1].height = 30

    diff_col_idx = cols.index("Difference") + 1 if "Difference" in cols else None

    def row_formulas(r):
        sp   = C.get("Selling Price", "");    inv  = C.get("Invoice Amount", "")
        gt   = C.get("GT (As Per Calc)", ""); qty  = C.get("Qty", "")
        com  = C.get("Commission", "");       colf = C.get("Collection Fee", "")
        ff   = C.get("Fixed Fee", "");        tc   = C.get("Total Charges", "")
        gst  = C.get("GST on Charges", "");   tv   = C.get("Taxable Value", "")
        tds  = C.get("TDS", "");              tcs  = C.get("TCS", "")
        td   = C.get("Total Deductions", ""); ra   = C.get("Received Amount", "")
        pwn  = C.get("PWN", "");              pb   = C.get("PWN Benchmark", "")
        diff = C.get("Difference", "")
        fmls = {}
        if sp and inv and gt:
            fmls["Selling Price"]    = f'=IF(OR({gt}{r}="",{gt}{r}=0),"",ROUND({inv}{r}-{gt}{r},2))'
        if tc and com and colf and ff:
            fmls["Total Charges"]    = f'=IF({sp}{r}="","",ROUND({com}{r}+{colf}{r}+{ff}{r},2))'
        if gst and tc:
            fmls["GST on Charges"]   = f'=IF({tc}{r}="","",ROUND({tc}{r}*0.18,2))'
        if tv and sp:
            fmls["Taxable Value"]    = f'=IF({sp}{r}="","",ROUND({sp}{r}-{sp}{r}/105*5,2))'
        if tds and tv:
            fmls["TDS"]              = f'=IF({tv}{r}="","",ROUND({tv}{r}*0.001,2))'
        if tcs and tv:
            fmls["TCS"]              = f'=IF({tv}{r}="","",ROUND({tv}{r}*0.005,2))'
        if td and tc and gst and tds and tcs:
            fmls["Total Deductions"] = f'=IF({tc}{r}="","",ROUND({tc}{r}+{gst}{r}+{tds}{r}+{tcs}{r},2))'
        if ra and sp and tc and gst and tds and tcs:
            fmls["Received Amount"]  = f'=IF({sp}{r}="","",ROUND({sp}{r}-{tc}{r}-{gst}{r}-{tds}{r}-{tcs}{r},2))'
        if pb and pwn and qty:
            fmls["PWN Benchmark"]    = f'=IF({pwn}{r}="","",ROUND({pwn}{r}*{qty}{r},2))'
        if diff and ra and pb:
            fmls["Difference"]       = f'=IF(OR({ra}{r}="",{pb}{r}=""),"",ROUND({ra}{r}-{pb}{r},2))'
        return fmls

    for r_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        alt_fill = PatternFill("solid", fgColor=C_ALT1 if r_idx % 2 == 0 else C_ALT2)
        fmls     = row_formulas(r_idx)
        for c_idx, (col_name, val) in enumerate(zip(cols, row_data), start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if col_name in fmls:
                cell.value = fmls[col_name]
            else:
                cell.value = None if (isinstance(val, float) and np.isnan(val)) else val
            cell.border = bdr
            cell.font   = Font(size=9, name="Calibri")
            cell.fill   = alt_fill
            if col_name in money_names:
                cell.number_format = '₹#,##0.00'
                cell.alignment     = Alignment(horizontal="right", vertical="center")
            elif col_name == "Qty":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            if c_idx == diff_col_idx:
                try:
                    v = float(val)
                    if not np.isnan(v):
                        if v < 0:
                            cell.fill = PatternFill("solid", fgColor=C_RED_BG)
                            cell.font = Font(color="C0392B", bold=True, size=9, name="Calibri")
                        elif v > 0:
                            cell.fill = PatternFill("solid", fgColor=C_GREEN_BG)
                            cell.font = Font(color="1E8449", bold=True, size=9, name="Calibri")
                        else:
                            cell.fill = PatternFill("solid", fgColor=C_ZERO_BG)
                            cell.font = Font(color="7D6608", bold=True, size=9, name="Calibri")
                except Exception:
                    pass
        ws.row_dimensions[r_idx].height = 16

    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = ws.dimensions
    last_data_row = len(df) + 1
    total_row     = last_data_row + 2
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=total_row, column=c_idx)
        cell.fill   = PatternFill("solid", fgColor=C_TOTAL_BG)
        cell.font   = Font(bold=True, color=C_TOTAL_FG, size=10, name="Calibri")
        cell.border = bdr_header
        if c_idx == 1:
            cell.value     = "TOTALS"
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif col_name in money_names:
            col_l = get_column_letter(c_idx)
            cell.value         = f"=SUM({col_l}2:{col_l}{last_data_row})"
            cell.number_format = '₹#,##0.00'
            cell.alignment     = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[total_row].height = 22


def apply_summary_style(ws):
    C_H = "2C3E50"; C_FG = "FFFFFF"
    C_ODD = "EBF5FB"; C_EVEN = "FFFFFF"
    thin = Side(style="thin", color="AED6F1")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill      = PatternFill("solid", fgColor=C_H)
        cell.font      = Font(bold=True, color=C_FG, size=10, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = bdr
    ws.row_dimensions[1].height = 24
    for r_idx in range(2, ws.max_row + 1):
        fill = PatternFill("solid", fgColor=C_ODD if r_idx % 2 == 0 else C_EVEN)
        for cell in ws[r_idx]:
            cell.fill = fill; cell.font = Font(size=9, name="Calibri")
            cell.border = bdr; cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[r_idx].height = 15
    for col_cells in ws.columns:
        width = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(width + 4, 40)
    last_row  = ws.max_row
    total_row = last_row + 2
    for c_idx in range(1, ws.max_column + 1):
        sample_cell = ws.cell(row=2, column=c_idx)
        tot_cell    = ws.cell(row=total_row, column=c_idx)
        tot_cell.fill   = PatternFill("solid", fgColor="2C3E50")
        tot_cell.font   = Font(bold=True, color="FFD700", size=10, name="Calibri")
        tot_cell.border = Border(
            left=Side(style="medium", color="2C3E50"), right=Side(style="medium", color="2C3E50"),
            top=Side(style="medium",  color="2C3E50"), bottom=Side(style="medium", color="2C3E50"),
        )
        if c_idx == 1:
            tot_cell.value = "TOTALS"
        elif isinstance(sample_cell.value, (int, float)):
            col_l = get_column_letter(c_idx)
            tot_cell.value  = f"=SUM({col_l}2:{col_l}{last_row})"
            tot_cell.number_format = '₹#,##0.00'
            tot_cell.alignment     = Alignment(horizontal="right", vertical="center")
    ws.freeze_panes = "A2"


def to_excel(recon_df, summary_df, cat_df) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        recon_df.to_excel(writer,   index=False, sheet_name="Reconciliation")
        cat_df.to_excel(writer,     index=False, sheet_name="Category Breakdown")
        summary_df.to_excel(writer, index=False, sheet_name="Charges Summary")
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
    for label, val in [
        ("Total Orders",              len(df)),
        ("Orders Calculated",         int(df["Received Amount"].notna().sum())),
        ("Orders NaN (no match)",     int(df["Received Amount"].isna().sum())),
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
    ]:
        totals["Metric"].append(label)
        totals["Value"].append(round(val, 2) if isinstance(val, float) else val)
    summary_df = pd.DataFrame(totals)

    cat_df = (
        valid.groupby("Sub-Category")
        .agg(
            Orders         = ("Order Id",        "count"),
            Invoice_Total  = ("Invoice Amount",   "sum"),
            GT_Total       = ("GT (As Per Calc)", "sum"),
            Selling_Total  = ("Selling Price",    "sum"),
            Commission     = ("Commission",       "sum"),
            Collection     = ("Collection Fee",   "sum"),
            Fixed          = ("Fixed Fee",        "sum"),
            Total_Charges  = ("Total Charges",    "sum"),
            GST_Total      = ("GST on Charges",   "sum"),
            TDS_Total      = ("TDS",              "sum"),
            TCS_Total      = ("TCS",              "sum"),
            Deductions     = ("Total Deductions", "sum"),
            Received_Total = ("Received Amount",  "sum"),
            Net_Diff       = ("Difference",       "sum"),
            Avg_Diff       = ("Difference",       "mean"),
        )
        .reset_index()
        .sort_values("Invoice_Total", ascending=False)
        .round(2)
    )
    cat_df.columns = [
        "Sub-Category", "Orders", "Invoice Total", "GT Total", "Selling Total",
        "Commission", "Collection Fee", "Fixed Fee", "Total Charges",
        "GST Total", "TDS Total", "TCS Total", "Total Deductions",
        "Received Total", "Net Difference", "Avg Difference",
    ]
    return summary_df, cat_df


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k, v in [
    ("pwn_overrides",   {}),
    ("sku_corrections", {}),
    ("result_df",       None),
    ("charges_df",      None),
    ("sku_info_dict",   {}),
    ("pwn_dict",        {}),
    ("order_df",        None),
    ("replace_map",     {}),
]:
    if k not in st.session_state:
        st.session_state[k] = v


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
if order_files and charges_file:

    with st.spinner("🔄 Reading files…"):
        order_df, file_info, file_errors = load_all_order_files(order_files)

        st.sidebar.markdown("---")
        st.sidebar.markdown("**📄 Uploaded Order Files**")
        st.sidebar.dataframe(pd.DataFrame(file_info), hide_index=True, use_container_width=True)
        st.sidebar.caption(f"Total rows loaded: **{len(order_df):,}**")

        for err in file_errors:
            st.warning(f"⚠️ {err}")
        if order_df.empty:
            st.error("❌ No valid order rows loaded. Check column names.")
            st.stop()

        xl     = pd.read_excel(charges_file, sheet_name=None, header=None)
        sheets = list(xl.values())
        if len(sheets) < 3:
            st.error(f"❌ Data Excel must have at least 3 sheets. Found: {list(xl.keys())}")
            st.stop()

        charges_df    = parse_charges_df(sheets[0])
        sku_info_dict = parse_sku_info(sheets[1])
        pwn_dict      = parse_pwn_dict(sheets[2])

        replace_map = {}
        if replace_sku_file:
            replace_map = parse_replace_map(replace_sku_file)
            st.sidebar.success(f"✅ Replace SKU loaded: {len(replace_map):,} entries")

        known_brands = sorted(charges_df["Brand Name"].dropna().unique().tolist()) if "Brand Name" in charges_df.columns else []
        st.sidebar.markdown(f"**Brands in Sheet 0:** {', '.join(known_brands)}")

        st.session_state.update({
            "charges_df":    charges_df,
            "sku_info_dict": sku_info_dict,
            "pwn_dict":      pwn_dict,
            "order_df":      order_df,
            "replace_map":   replace_map,
        })

        # ═══════════════════════════════════════════════════════════════════════════
        # DIAGNOSTIC SECTION - Check Data Mapping
        # ═══════════════════════════════════════════════════════════════════════════
        with st.expander("🔍 Debug: Check Data Mapping", expanded=False):
            st.markdown("#### Sample SKU Lookups")
            
            # Show sample SKUs and their mappings
            sample_skus = order_df["SKU"].head(15).tolist()
            debug_results = []
            
            for raw_sku in sample_skus:
                # Apply same logic as reconciliation
                corrected_raw = st.session_state.get("sku_corrections", {}).get(raw_sku.upper(), raw_sku)
                lookup_sku = corrected_raw.strip().upper()
                if lookup_sku in replace_map:
                    corrected_raw = replace_map[lookup_sku]
                    
                sku = strip_vendor_prefix(corrected_raw)
                info = sku_info_dict.get(sku.upper(), {})
                brand = info.get("brand", "NOT FOUND")
                sub_cat = info.get("sub_cat", "NOT FOUND")
                
                # Normalize
                brand_clean = brand.strip() if brand and str(brand).lower() != "nan" else ""
                cat_clean = sub_cat.strip() if sub_cat and str(sub_cat).lower() != "nan" else ""
                
                charge_debug = debug_charge_lookup(brand_clean, cat_clean, charges_df) if brand_clean and cat_clean else {"exact_match_found": False, "matched_rows": 0}
                
                debug_results.append({
                    "Original SKU": raw_sku,
                    "After Replace": corrected_raw if corrected_raw != raw_sku else "—",
                    "Cleaned SKU": sku,
                    "Brand Found": brand_clean or "❌ NOT FOUND",
                    "Sub-Category": cat_clean or "❌ NOT FOUND",
                    "Charge Match": "✅ Yes" if charge_debug["exact_match_found"] else "❌ No",
                    "Matched Rows": charge_debug["matched_rows"]
                })
            
            st.dataframe(pd.DataFrame(debug_results), use_container_width=True)
            
            st.markdown("#### Available Brands in Sheet 0 (Charges)")
            if "Brand Name" in charges_df.columns:
                brands_list = sorted([str(b).strip() for b in charges_df["Brand Name"].dropna().unique() if str(b).lower() != 'nan'])
                st.write(", ".join(brands_list))
            
            st.markdown("#### Available Categories in Sheet 0 (Charges)")
            if "Category" in charges_df.columns:
                cats_list = sorted([str(c).strip() for c in charges_df["Category"].dropna().unique() if str(c).lower() != 'nan'])
                st.write(", ".join(cats_list))
            
            st.markdown("#### Brand-Category Combinations in Sheet 0")
            if "Brand Name" in charges_df.columns and "Category" in charges_df.columns:
                combos = charges_df[["Brand Name", "Category"]].drop_duplicates().dropna()
                combos = combos[
                    (combos["Brand Name"].astype(str).str.lower() != 'nan') & 
                    (combos["Category"].astype(str).str.lower() != 'nan')
                ].head(30)
                st.dataframe(combos, use_container_width=True)

        # ═══════════════════════════════════════════════════════════════════════════
        # SHEET STRUCTURE VERIFICATION
        # ═══════════════════════════════════════════════════════════════════════════
        with st.expander("📊 Verify Excel Sheet Structure", expanded=False):
            st.markdown("### Sheet 0: Charges Rates")
            st.write(f"**Rows:** {len(charges_df)}")
            st.write(f"**Columns:** {list(charges_df.columns)}")
            st.dataframe(charges_df.head(15), use_container_width=True)
            
            st.markdown("### Sheet 1: Category Description (Sample)")
            sample_sku_info = pd.DataFrame([
                {"SKU": k, "Brand": v.get("brand", ""), "Sub-Category": v.get("sub_cat", "")}
                for k, v in list(sku_info_dict.items())[:15]
            ])
            st.dataframe(sample_sku_info, use_container_width=True)
            
            st.markdown("### Sheet 2: PWN Prices (Sample)")
            sample_pwn = pd.DataFrame([
                {"SKU": k, "PWN Price": v}
                for k, v in list(pwn_dict.items())[:15]
            ])
            st.dataframe(sample_pwn, use_container_width=True)
            
            if replace_map:
                st.markdown("### Sheet 3: Replace SKU Mapping (Sample)")
                sample_replace = pd.DataFrame([
                    {"Seller SKU": k, "→ OMS SKU": v}
                    for k, v in list(replace_map.items())[:15]
                ])
                st.dataframe(sample_replace, use_container_width=True)

    result_df = run_reconciliation(
        st.session_state["order_df"],
        st.session_state["charges_df"],
        st.session_state["sku_info_dict"],
        st.session_state["pwn_dict"],
        fixed_fee, gst_rate,
        replace_map=st.session_state["replace_map"],
        pwn_overrides=st.session_state["pwn_overrides"],
        sku_corrections=st.session_state["sku_corrections"],
    )
    st.session_state["result_df"] = result_df
    summary_df, cat_df = build_summary(result_df)

    replace_resolved = result_df[result_df["PWN Match"].str.startswith("replace", na=False)]
    st.success(
        f"✅ Processed **{len(result_df):,}** orders  |  "
        f"**{int(result_df['Received Amount'].notna().sum()):,}** calculated  |  "
        f"**{int(result_df['Received Amount'].isna().sum()):,}** skipped (no match)"
        + (f"  |  **{len(replace_resolved):,}** PWN via Replace SKU map" if len(replace_resolved) else "")
    )

    tab1, tab2, tab3 = st.tabs(["📋  Reconciliation", "💰  Charges Summary", "📊  Category Breakdown"])

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 1 – RECONCILIATION                                         ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab1:
        broken_df = result_df[
            result_df["Received Amount"].isna() | (result_df["PWN Match"] == "not_found")
        ]
        if len(broken_df):
            no_cat = int(broken_df["Received Amount"].isna().sum())
            no_pwn = int((broken_df["PWN Match"] == "not_found").sum())
            with st.expander(
                f"✏️  **{len(broken_df)} SKU(s) have lookup issues — correct SKU name to retry**",
                expanded=False,
            ):
                st.info("Type the corrected SKU — the tool re-runs all lookups using the corrected name.")
                st.caption(f"🔴 No category/GT: **{no_cat}**  |  🟡 No PWN: **{no_pwn}**")
                broken_skus = broken_df["SKU"].unique().tolist()
                correction_inputs = {}
                h1, h2, h3, h4 = st.columns([3, 2, 2, 3])
                h1.markdown("**Original SKU**"); h2.markdown("**Issue**")
                h3.markdown("**Corrected SKU**"); h4.markdown("**Live preview**")
                st.markdown("---")
                for sku in broken_skus:
                    sku_rows  = broken_df[broken_df["SKU"] == sku]
                    issues    = []
                    if sku_rows["Received Amount"].isna().any(): issues.append("❌ No category/GT")
                    if (sku_rows["PWN Match"] == "not_found").any(): issues.append("⚠️ No PWN")
                    existing  = st.session_state["sku_corrections"].get(sku.upper(), "")
                    c1, c2, c3, c4 = st.columns([3, 2, 2, 3])
                    c1.markdown(f"<div style='padding-top:6px;font-size:0.88rem;word-break:break-all'><code>{sku}</code></div>", unsafe_allow_html=True)
                    c2.markdown(f"<div style='padding-top:6px;font-size:0.82rem'>{'  &  '.join(issues)}</div>", unsafe_allow_html=True)
                    corrected = c3.text_input("Corrected SKU", value=existing, placeholder="e.g. YK1234-L",
                                              label_visibility="collapsed", key=f"sku_corr_{sku}")
                    correction_inputs[sku] = corrected.strip()
                    if existing:
                        lsku = strip_vendor_prefix(existing)
                        info = st.session_state["sku_info_dict"].get(lsku.upper(), {})
                        sub_cat_p = info.get("sub_cat", "")
                        pwn_v, _ = lookup_pwn_with_replace(lsku, st.session_state["pwn_dict"], st.session_state["replace_map"])
                        parts = []
                        if sub_cat_p and sub_cat_p != "nan": parts.append(f"📦 Sub-cat: *{sub_cat_p}*")
                        if pd.notna(pwn_v): parts.append(f"💰 PWN: ₹{pwn_v:,.2f}")
                        html = ("<div style='padding-top:4px;font-size:0.80rem;color:#1a7a3c;line-height:1.6'>" + "<br>".join(parts) + "</div>") if parts else "<div style='padding-top:6px;font-size:0.80rem;color:#c0392b'>⚠️ Still unresolved</div>"
                        c4.markdown(html, unsafe_allow_html=True)
                    else:
                        c4.markdown("<div style='padding-top:6px;font-size:0.80rem;color:#aaa'>— type a correction to preview —</div>", unsafe_allow_html=True)
                st.markdown("---")
                cs, cc = st.columns([2, 1])
                if cs.button("💾  Save SKU Corrections & Recalculate", type="primary"):
                    st.session_state["sku_corrections"] = {o.upper(): c for o, c in correction_inputs.items() if c}
                    st.rerun()
                if cc.button("🗑️  Clear All Corrections"):
                    st.session_state["sku_corrections"] = {}
                    st.rerun()

        if st.session_state["sku_corrections"]:
            with st.expander(f"✅  **{len(st.session_state['sku_corrections'])} SKU correction(s) active**", expanded=False):
                st.dataframe(pd.DataFrame([{"Original SKU": o, "→ Corrected SKU": c}
                    for o, c in st.session_state["sku_corrections"].items()]),
                    use_container_width=True, hide_index=True)
                if st.button("🗑️  Clear All Active Corrections", key="clear_corr_summary"):
                    st.session_state["sku_corrections"] = {}
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
        k10.metric("Net Difference", f"₹{net:,.2f}",
                   delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
                   delta_color="normal" if net >= 0 else "inverse")
        st.markdown("---")

        f1, f2, f3, f4 = st.columns([2, 2, 2, 3])
        sel_cat   = f1.selectbox("Sub-Category", ["All"] + sorted(result_df["Sub-Category"].dropna().unique().tolist()))
        sel_brand = f2.selectbox("Brand",         ["All"] + sorted(result_df["Brand Name"].dropna().unique().tolist()))
        diff_opt  = f3.selectbox("Difference type", ["All", "Positive (+)", "Negative (−)", "Zero / Matched", "No PWN data", "No Category (NaN)"])
        search    = f4.text_input("🔎 Search by SKU or Order ID")

        view = result_df.copy()
        if sel_cat   != "All": view = view[view["Sub-Category"] == sel_cat]
        if sel_brand != "All": view = view[view["Brand Name"]   == sel_brand]
        if diff_opt == "Positive (+)":        view = view[view["Difference"] > 0]
        elif diff_opt == "Negative (−)":      view = view[view["Difference"] < 0]
        elif diff_opt == "Zero / Matched":    view = view[view["Difference"] == 0]
        elif diff_opt == "No PWN data":       view = view[view["PWN Match"] == "not_found"]
        elif diff_opt == "No Category (NaN)": view = view[view["Received Amount"].isna()]
        if search.strip():
            mask = (view["SKU"].str.contains(search.strip(), case=False, na=False) |
                    view["Order Id"].str.contains(search.strip(), case=False, na=False))
            view = view[mask]

        st.caption(f"Showing **{len(view):,}** of **{len(result_df):,}** orders")

        display_cols = [
            "Order Id", "SKU", "Product", "Brand Name", "Ordered On",
            "Sub-Category", "Charge Method",
            "Qty", "Invoice Amount", "GT (As Per Calc)", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges", "Taxable Value", "TDS", "TCS",
            "Total Deductions", "Received Amount",
            "PWN", "PWN Benchmark", "Difference", "PWN Match",
        ]
        st.dataframe(style_table(view[display_cols], diff_col="Difference"),
                     use_container_width=True, height=500)

        st.markdown("### 📥 Download")
        d1, d2 = st.columns(2)
        d1.download_button("⬇  Full Reconciliation (Excel – 3 sheets, styled)",
            data=to_excel(result_df[display_cols], summary_df, cat_df),
            file_name="flipkart_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("⬇  Filtered View (Excel, styled)",
            data=to_excel(view[display_cols].reset_index(drop=True), summary_df, cat_df),
            file_name="flipkart_reconciliation_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ╔══════════════════════════════════════════════════════════════════╗
    # ║  TAB 2 – CHARGES SUMMARY                                        ║
    # ╚══════════════════════════════════════════════════════════════════╝
    with tab2:
        st.markdown("### 💰 Total Charges Summary")
        valid = result_df[result_df["Received Amount"].notna()]
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("#### 📤 Flipkart Deductions")
            a1, a2 = st.columns(2)
            a1.metric("Commission",      f"₹{valid['Commission'].sum():,.2f}")
            a2.metric("Collection Fee",  f"₹{valid['Collection Fee'].sum():,.2f}")
            a1.metric("Fixed Fee",       f"₹{valid['Fixed Fee'].sum():,.2f}")
            a2.metric("GST on Charges",  f"₹{valid['GST on Charges'].sum():,.2f}")
            a1.metric("TDS (0.1%)",      f"₹{valid['TDS'].sum():,.2f}")
            a2.metric("TCS (0.5%)",      f"₹{valid['TCS'].sum():,.2f}")
            st.metric("🔴 Total Deductions", f"₹{valid['Total Deductions'].sum():,.2f}")
        with col_b:
            st.markdown("#### 📥 What You Receive")
            b1, b2 = st.columns(2)
            b1.metric("Total Invoice",   f"₹{result_df['Invoice Amount'].sum():,.2f}")
            b2.metric("GT Total (ref)",  f"₹{valid['GT (As Per Calc)'].sum():,.2f}")
            b1.metric("Selling Total",   f"₹{valid['Selling Price'].sum():,.2f}")
            b2.metric("Total Received",  f"₹{valid['Received Amount'].sum():,.2f}")
            net = valid["Difference"].sum()
            b1.metric("Net Diff vs PWN", f"₹{net:,.2f}",
                      delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
                      delta_color="normal" if net >= 0 else "inverse")
            b2.metric("Orders –ve Diff", int((valid["Difference"] < 0).sum()))

        st.info(
            "ℹ️  **GT** → slab lookup on Invoice Amount (fixed ₹)  \n"
            "**Commission & Collection** → independent slab lookups on Selling Price  \n"
            "**Received Amount** = Selling Price − Total Charges − GST − TDS − TCS  \n"
            "**Taxable Value** = Selling Price − (Selling Price / 105 × 5)  \n"
            "**Difference** = Received Amount − (Qty × PWN)"
        )
        st.markdown("---")
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols = [
            "Order Id", "SKU", "Brand Name", "Sub-Category", "Charge Method",
            "Invoice Amount", "GT (As Per Calc)", "Selling Price",
            "Commission", "Collection Fee", "Fixed Fee",
            "Total Charges", "GST on Charges", "Taxable Value", "TDS", "TCS",
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
        cat_money = [c for c in cat_df.columns if c not in ("Sub-Category", "Orders")]
        st.dataframe(
            style_table(cat_df, diff_col="Net Difference").format({c: "₹{:.2f}" for c in cat_money}),
            use_container_width=True,
        )
        st.markdown("---")
        st.markdown("#### 🔢 Charge Components Only (per Sub-Category)")
        comp_cols = ["Sub-Category", "Orders", "GT Total", "Commission", "Collection Fee",
                     "Fixed Fee", "GST Total", "TDS Total", "TCS Total", "Total Deductions"]
        comp_money = [c for c in comp_cols if c not in ("Sub-Category", "Orders")]
        st.dataframe(cat_df[comp_cols].style.format({c: "₹{:.2f}" for c in comp_money}),
                     use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# LANDING SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
else:
    st.info("👈 Upload **order file(s)** and the **Data Excel** in the sidebar to begin.")
    st.markdown("""
---
### How it works

| File | Description |
|------|-------------|
| **Order File(s)** | Flipkart Seller Hub export — CSV, XLSX or XLS. Needs: `Order Id`, `SKU`, `Product`, `Ordered On`, `Invoice Amount`, `Quantity` |
| **Data Excel** | Yash Gallery workbook — 3 sheets |
| **Replace SKU Excel** *(optional)* | Maps Seller SKU Id → OMS SKU for PWN fallback |

**Excel sheet positions:**

| Position | Sheet | Used For |
|----------|-------|----------|
| Index 0 | Data For Flipkart Roc | Brand + Category → GT / Commission / Collection slabs |
| Index 1 | Category Discription | Seller SKU → Brand + Sub-category |
| Index 2 | Price We Need | OMS Child SKU → PWN price |

---
### ✅ Correct Charge Calculation Logic

**Step 1 — SKU Resolution Flow**
1. Check if SKU has manual correction (from UI)
2. Apply Replace SKU mapping (from Replace SKU Excel) 
3. Strip vendor prefix (GWN-, SPF-, KL-, etc.)
4. Look up in Category Description sheet

**Step 2 — Brand & Sub-category from Sheet 1**
Each SKU in Sheet 1 already has Brand + Sub-category. This is used directly — brand is not guessed from product name.

**Step 3 — Three INDEPENDENT slab lookups from Sheet 0**

| Charge | Input | How |
|--------|-------|-----|
| **GT** | Invoice Amount | Find row where GT Lower ≤ Invoice ≤ GT Upper → fixed ₹ |
| **Commission** | Selling Price (= Invoice − GT) | Find row where Comm Lower ≤ Sell ≤ Comm Upper → Sell × % |
| **Collection** | Selling Price | Find row where Coll Lower < Sell ≤ Coll Upper → Sell × % |

Each charge scans **all rows** for that Brand+Category independently.

**Step 4 — Final calculation**
```
Total Charges    = Commission + Collection Fee + Fixed Fee
GST on Charges   = Total Charges × 18%
Taxable Value    = Selling Price − (Selling Price / 105 × 5)
TDS              = Taxable Value × 0.1%
TCS              = Taxable Value × 0.5%
Received Amount  = Selling Price − Total Charges − GST − TDS − TCS
Difference       = Received Amount − (Qty × PWN)
```
""")
