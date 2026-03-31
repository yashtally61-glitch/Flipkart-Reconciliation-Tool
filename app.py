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
    charges_file    = st.file_uploader("2️⃣  Data Excel", type=["xlsx"])
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
- **Sheet 0** — Charges Description (Brand + Category based rates)
- **Sheet 1** — Category Description (SKU → Sub-category)
- **Sheet 2** — Price We Need (PWN prices)

> **Brand-based charging:** Brand is read from the **Product** column  
> in the Order File. Rates are looked up by **Brand + Category**  
> from Sheet 0. Falls back to category-only if brand not found.
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

VENDOR_PREFIXES = ["GWN-", "GWN_", "GWN", "SPF-", "SPF_", "SPF", "KL_", "KL-", "KL"]


# ═══════════════════════════════════════════════════════════════════════════════
# BRAND EXTRACTION FROM PRODUCT COLUMN (ORDER FILE)
# ═══════════════════════════════════════════════════════════════════════════════

def extract_brand_from_product(product: str, known_brands: list) -> str:
    """
    Extract brand name from Product column in the Order File.
    Tries to match against known brands (from Sheet 0) first,
    then falls back to first 1-2 capitalized words.
    """
    if not product or product == "nan" or pd.isna(product):
        return ""

    product = str(product).strip()

    # Match against known brands (longest match first to avoid partial hits)
    for brand in sorted(known_brands, key=len, reverse=True):
        if product.lower().startswith(brand.lower()):
            return brand

    # Fallback: take first 1-2 words if capitalized
    words = product.split()
    if len(words) >= 2 and words[0][0].isupper() and words[1][0].isupper():
        return f"{words[0]} {words[1]}"
    return words[0] if words else ""


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
    if len(parts) == 2:
        return parts[0]
    return key


def lookup_cat_by_base(sku: str, sku_cat_dict: dict) -> tuple[str, str]:
    base = get_sku_base(sku)
    if not base:
        return "", ""
    for candidate_sku, sub_cat in sku_cat_dict.items():
        if get_sku_base(candidate_sku) == base and sub_cat and sub_cat != "nan":
            return sub_cat, candidate_sku
    return "", ""


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

    upper_sku = sku.strip().upper()
    oms_sku = replace_map.get(upper_sku)
    if oms_sku:
        pwn_val2, method2 = lookup_pwn(oms_sku, pwn_dict)
        if method2 != "not_found":
            return pwn_val2, f"replace→{method2}"

    return np.nan, "not_found"


# ═══════════════════════════════════════════════════════════════════════════════
# BRAND + CATEGORY CHARGE LOOKUP (Sheet 0 has both Brand Name and Category)
# ═══════════════════════════════════════════════════════════════════════════════

def get_brand_cat_gt_amount(brand: str, cat: str, inv_amount: float, charges_df: pd.DataFrame) -> float:
    """Lookup GT charge by Brand Name + Category from Sheet 0."""
    if not brand or not cat or charges_df is None or charges_df.empty:
        return np.nan

    rows = charges_df[
        (charges_df["Brand Name"].str.lower() == brand.strip().lower()) &
        (charges_df["Category"].str.lower() == cat.strip().lower())
    ]
    for _, r in rows.iterrows():
        lo = r.get("GT Lower Limit")
        hi = r.get("GT Upper Limit")
        gt = r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gt):
            if float(lo) <= inv_amount <= float(hi) + 0.99:
                return float(gt)
    return np.nan


def get_brand_cat_commission(brand: str, cat: str, sell: float, charges_df: pd.DataFrame) -> float:
    """Lookup commission by Brand Name + Category from Sheet 0."""
    if not brand or not cat or charges_df is None or charges_df.empty:
        return np.nan

    rows = charges_df[
        (charges_df["Brand Name"].str.lower() == brand.strip().lower()) &
        (charges_df["Category"].str.lower() == cat.strip().lower())
    ]
    for _, r in rows.iterrows():
        lo = r.get("Lower Limit Commision")
        hi = r.get("Upper Limit Commision")
        ch = r.get("Commision Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell <= float(hi) + 0.99:
                return round(float(ch) * sell, 5)
    return np.nan


def get_brand_cat_collection_fee(brand: str, cat: str, sell: float, charges_df: pd.DataFrame) -> float:
    """Lookup collection fee by Brand Name + Category from Sheet 0."""
    if not brand or not cat or charges_df is None or charges_df.empty:
        return np.nan

    rows = charges_df[
        (charges_df["Brand Name"].str.lower() == brand.strip().lower()) &
        (charges_df["Category"].str.lower() == cat.strip().lower())
    ]
    for _, r in rows.iterrows():
        lo_raw = r.get("Collection Lower Limit")
        hi     = r.get("Collection Upper Limit")
        cf     = r.get("Collection Charge")
        if pd.isna(hi) or pd.isna(cf):
            continue
        lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).strip().startswith(">")) else float(lo_raw)
        if lo_val < sell <= float(hi) + 0.99:
            return round(float(cf) * sell, 5)
    return np.nan


# Category-only fallback functions (for when brand not found in Sheet 0)
def get_gt_amount(cat: str, inv_amount: float, charges_df: pd.DataFrame) -> float:
    rows = charges_df[charges_df["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("GT Lower Limit")
        hi = r.get("GT Upper Limit")
        gt = r.get("GT Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(gt):
            if float(lo) <= inv_amount <= float(hi) + 0.99:
                return float(gt)
    return np.nan


def get_commission(cat: str, sell: float, charges_df: pd.DataFrame) -> float:
    rows = charges_df[charges_df["Category"].str.lower() == cat.strip().lower()]
    for _, r in rows.iterrows():
        lo = r.get("Lower Limit Commision")
        hi = r.get("Upper Limit Commision")
        ch = r.get("Commision Charge")
        if pd.notna(lo) and pd.notna(hi) and pd.notna(ch):
            if float(lo) <= sell <= float(hi) + 0.99:
                return round(float(ch) * sell, 5)
    return 0.0


def get_collection_fee(cat: str, sell: float, charges_df: pd.DataFrame) -> float:
    rows = charges_df[charges_df["Category"].str.lower() == cat.strip().lower()]
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
    return selling_price - (selling_price / 105 * 5)


def calc_tds(taxable_value: float, rate: float = 0.001) -> float:
    return round(taxable_value * rate, 5)


def calc_tcs(taxable_value: float, rate: float = 0.005) -> float:
    return round(taxable_value * rate, 5)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_charges_df(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse Sheet 0: has Brand Name + Category + commission/collection/GT slabs.
    Header is in row 0, data starts from row 1.
    """
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    # Keep rows where Category is present
    df = df[df["Category"].notna()].copy()
    # Forward-fill Brand Name and Category if they span multiple slab rows
    if "Brand Name" in df.columns:
        df["Brand Name"] = df["Brand Name"].ffill()
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


def parse_sku_cat(raw: pd.DataFrame) -> tuple[dict, dict]:
    """
    Parse Sheet 1 (Category Discription): Seller SKU Id → Sub-category.
    Also returns sku → brand mapping if Brand column exists.
    Returns (sku_cat_dict, sku_brand_dict).
    """
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    sku_cat_dict = dict(zip(
        df["Seller SKU Id"].astype(str).str.strip().str.upper(),
        df["Sub-category"].astype(str).str.strip(),
    ))

    sku_brand_dict = {}
    if "Brand" in df.columns:
        sku_brand_dict = dict(zip(
            df["Seller SKU Id"].astype(str).str.strip().str.upper(),
            df["Brand"].astype(str).str.strip(),
        ))

    return sku_cat_dict, sku_brand_dict


def parse_pwn_dict(raw: pd.DataFrame) -> dict:
    df = raw.copy()
    df.columns = raw.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"]    = pd.to_numeric(df["PWN+10%+50"], errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(), df["PWN+10%+50"]))


def parse_replace_map(file) -> dict:
    xl = pd.read_excel(file, header=None)
    df = xl.copy()
    df.columns = xl.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(
        df["Seller SKU Id"].astype(str).str.strip().str.upper(),
        df["OMS SKU"].astype(str).str.strip().str.upper(),
    ))


# ═══════════════════════════════════════════════════════════════════════════════
# MULTI-FILE ORDER READER
# ═══════════════════════════════════════════════════════════════════════════════

REQUIRED_ORDER_COLS = {"Order Id", "SKU", "Invoice Amount", "Quantity", "Product"}

def read_order_file(f) -> tuple[pd.DataFrame, str]:
    name = f.name.lower()
    try:
        if name.endswith(".csv"):
            raw = f.read()
            for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(BytesIO(raw), encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                return pd.DataFrame(), f"Could not decode '{f.name}' with any known encoding."
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
            df = pd.read_excel(f, engine=engine)
        else:
            return pd.DataFrame(), f"Unsupported file type: '{f.name}'"

        df.columns = [str(c).strip() for c in df.columns]
        missing = REQUIRED_ORDER_COLS - set(df.columns)
        if missing:
            return pd.DataFrame(), (
                f"'{f.name}' is missing required columns: {', '.join(sorted(missing))}"
            )

        df["_source_file"] = f.name
        return df, ""

    except Exception as e:
        return pd.DataFrame(), f"Error reading '{f.name}': {e}"


def load_all_order_files(files) -> tuple[pd.DataFrame, list[dict], list[str]]:
    frames    = []
    file_info = []
    errors    = []

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

def run_reconciliation(order_df, charges_df, sku_cat_dict, pwn_dict,
                       cat_map, manual_cat_map,
                       fixed_fee, gst_rate,
                       known_brands: list = None,
                       replace_map: dict = None,
                       pwn_overrides: dict = None,
                       sku_corrections: dict = None) -> pd.DataFrame:
    known_brands    = known_brands    or []
    pwn_overrides   = pwn_overrides   or {}
    replace_map     = replace_map     or {}
    sku_corrections = sku_corrections or {}
    rows_out = []

    for _, row in order_df.iterrows():
        raw_sku    = str(row.get("SKU", "")).strip()
        product    = str(row.get("Product", "")).strip()

        # Apply manual SKU correction if user provided one
        corrected_raw = sku_corrections.get(raw_sku.upper(), raw_sku)
        sku           = strip_vendor_prefix(corrected_raw)

        order_id    = str(row.get("Order Id", "")).strip()
        ordered_on  = row.get("Ordered On", "")
        inv_amount  = float(row.get("Invoice Amount", 0) or 0)
        quantity    = int(row.get("Quantity", 1) or 1)

        # ── Extract Brand from Product column (Order File) ────────────────
        brand_name = extract_brand_from_product(product, known_brands)

        # ── Category lookup: exact → base-code fallback ──────────────────
        sub_cat_raw    = sku_cat_dict.get(sku.upper(), "")
        cat_match_note = ""

        if not sub_cat_raw or sub_cat_raw == "nan":
            fallback_sub, fallback_sku = lookup_cat_by_base(sku, sku_cat_dict)
            if fallback_sub:
                sub_cat_raw    = fallback_sub
                cat_match_note = f"base-cat({fallback_sku})"

        cat = get_cat_for_lookup(sub_cat_raw, cat_map, manual_cat_map)

        # ── CHARGE LOOKUP: Brand + Category first, then Category-only ────
        charge_method = ""
        gt_val = np.nan

        # Try Brand + Category lookup (Sheet 0 has both columns)
        if brand_name and cat:
            gt_val     = get_brand_cat_gt_amount(brand_name, cat, inv_amount, charges_df)
            commission = np.nan
            coll_fee   = np.nan

            if pd.notna(gt_val):
                sell_price = round(inv_amount - gt_val, 5)
                commission = get_brand_cat_commission(brand_name, cat, sell_price, charges_df)
                coll_fee   = get_brand_cat_collection_fee(brand_name, cat, sell_price, charges_df)

                if pd.notna(commission) and pd.notna(coll_fee):
                    charge_method = f"brand+cat:{brand_name}|{cat}"
                else:
                    gt_val = np.nan  # partial match — fall through

        # Fallback: Category-only lookup
        if pd.isna(gt_val) and cat:
            gt_val     = get_gt_amount(cat, inv_amount, charges_df)
            commission = np.nan
            coll_fee   = np.nan

            if pd.notna(gt_val):
                sell_price = round(inv_amount - gt_val, 5)
                commission = get_commission(cat, sell_price, charges_df)
                coll_fee   = get_collection_fee(cat, sell_price, charges_df)
                charge_method = f"category:{cat}"

        # ── Calculate final amounts ───────────────────────────────────────
        if pd.isna(gt_val) or (pd.isna(commission) and pd.isna(coll_fee)):
            sell_price       = np.nan
            gt_val           = np.nan
            commission       = coll_fee = total_charges = np.nan
            gst_on_charges   = np.nan
            taxable_value    = np.nan
            tds              = np.nan
            tcs              = np.nan
            total_deductions = received_amount = np.nan
            charge_method    = "not_found"
        else:
            sell_price       = round(inv_amount - gt_val, 5)
            commission       = commission if pd.notna(commission) else 0.0
            coll_fee         = coll_fee   if pd.notna(coll_fee)   else 0.0
            total_charges    = round(commission + coll_fee + float(fixed_fee), 5)
            gst_on_charges   = round(total_charges * gst_rate, 5)

            taxable_value    = round(calc_taxable_value(sell_price), 5)
            tds              = calc_tds(taxable_value)
            tcs              = calc_tcs(taxable_value)

            total_deductions = round(total_charges + gst_on_charges + tds + tcs, 5)
            received_amount  = round(sell_price - total_charges - gst_on_charges - tds - tcs, 5)

        # ── PWN lookup ───────────────────────────────────────────────────
        pwn_val, match_method = lookup_pwn_with_replace(sku, pwn_dict, replace_map)
        if sku.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku.upper()]), "manual"

        full_match_note = match_method
        if cat_match_note:
            full_match_note = f"{match_method} | cat:{cat_match_note}"

        # ── Difference: Received Amount − (Qty × PWN) ────────────────────
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
            "Charges Category": cat,
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
# STYLED EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

def apply_roc_sheet_style(ws, df: pd.DataFrame):
    """Apply rich formatting + live Excel formulas to the Reconciliation sheet."""

    C_HEADER_BG = "1A3C5E"
    C_HEADER_FG = "FFFFFF"
    C_ALT1      = "EAF2FB"
    C_ALT2      = "FFFFFF"
    C_GREEN_BG  = "D6EFDD"
    C_RED_BG    = "FDDEDE"
    C_ZERO_BG   = "FFF9E6"
    C_TOTAL_BG  = "1A3C5E"
    C_TOTAL_FG  = "FFD700"
    C_BORDER    = "B0C4D8"

    thin       = Side(style="thin",   color=C_BORDER)
    thick      = Side(style="medium", color="1A3C5E")
    bdr        = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
    bdr_header = Border(left=thick, right=thick, top=thick, bottom=thick)

    money_names = set(MONEY_COLS)
    cols        = df.columns.tolist()

    C = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}

    col_widths = {
        "Order Id": 20, "SKU": 28,
        "Product": 40, "Brand Name": 18, "Ordered On": 14,
        "Sub-Category": 20, "Charges Category": 18, "Charge Method": 24,
        "Qty": 6, "Invoice Amount": 15, "GT (As Per Calc)": 15,
        "Selling Price": 15, "Commission": 14, "Collection Fee": 15,
        "Fixed Fee": 10, "Total Charges": 15, "GST on Charges": 15,
        "Taxable Value": 14, "TDS": 10, "TCS": 10,
        "Total Deductions": 16, "Received Amount": 16,
        "PWN": 12, "PWN Benchmark": 15, "PWN Match": 14, "Difference": 14,
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
        sp   = C.get("Selling Price",    "")
        inv  = C.get("Invoice Amount",   "")
        gt   = C.get("GT (As Per Calc)", "")
        qty  = C.get("Qty",              "")
        com  = C.get("Commission",       "")
        colf = C.get("Collection Fee",   "")
        ff   = C.get("Fixed Fee",        "")
        tc   = C.get("Total Charges",    "")
        gst  = C.get("GST on Charges",   "")
        tv   = C.get("Taxable Value",    "")
        tds  = C.get("TDS",              "")
        tcs  = C.get("TCS",              "")
        td   = C.get("Total Deductions", "")
        ra   = C.get("Received Amount",  "")
        pwn  = C.get("PWN",              "")
        pb   = C.get("PWN Benchmark",    "")
        diff = C.get("Difference",       "")

        formulas = {}

        if sp and inv and gt:
            formulas["Selling Price"] = (
                f'=IF(OR({gt}{r}="",{gt}{r}=0),"",'
                f'ROUND({inv}{r}-{gt}{r},2))'
            )
        if tc and com and colf and ff:
            formulas["Total Charges"] = (
                f'=IF({sp}{r}="","",ROUND({com}{r}+{colf}{r}+{ff}{r},2))'
            )
        if gst and tc:
            formulas["GST on Charges"] = (
                f'=IF({tc}{r}="","",ROUND({tc}{r}*0.18,2))'
            )
        if tv and sp:
            formulas["Taxable Value"] = (
                f'=IF({sp}{r}="","",ROUND({sp}{r}-{sp}{r}/105*5,2))'
            )
        if tds and tv:
            formulas["TDS"] = (
                f'=IF({tv}{r}="","",ROUND({tv}{r}*0.001,2))'
            )
        if tcs and tv:
            formulas["TCS"] = (
                f'=IF({tv}{r}="","",ROUND({tv}{r}*0.005,2))'
            )
        if td and tc and gst and tds and tcs:
            formulas["Total Deductions"] = (
                f'=IF({tc}{r}="","",ROUND({tc}{r}+{gst}{r}+{tds}{r}+{tcs}{r},2))'
            )
        if ra and sp and tc and gst and tds and tcs:
            formulas["Received Amount"] = (
                f'=IF({sp}{r}="","",ROUND({sp}{r}-{tc}{r}-{gst}{r}-{tds}{r}-{tcs}{r},2))'
            )
        if pb and pwn and qty:
            formulas["PWN Benchmark"] = (
                f'=IF({pwn}{r}="","",ROUND({pwn}{r}*{qty}{r},2))'
            )
        if diff and ra and pb:
            formulas["Difference"] = (
                f'=IF(OR({ra}{r}="",{pb}{r}=""),"",ROUND({ra}{r}-{pb}{r},2))'
            )

        return formulas

    for r_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        alt_fill = PatternFill("solid", fgColor=C_ALT1 if r_idx % 2 == 0 else C_ALT2)
        formulas = row_formulas(r_idx)
        diff_val = getattr(row_data, "Difference", None)

        for c_idx, (col_name, val) in enumerate(zip(cols, row_data), start=1):
            cell = ws.cell(row=r_idx, column=c_idx)

            if col_name in formulas:
                cell.value = formulas[col_name]
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
        cell        = ws.cell(row=total_row, column=c_idx)
        cell.fill   = PatternFill("solid", fgColor=C_TOTAL_BG)
        cell.font   = Font(bold=True, color=C_TOTAL_FG, size=10, name="Calibri")
        cell.border = bdr_header

        if c_idx == 1:
            cell.value     = "TOTALS"
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif col_name in money_names:
            col_l              = get_column_letter(c_idx)
            cell.value         = f"=SUM({col_l}2:{col_l}{last_data_row})"
            cell.number_format = '₹#,##0.00'
            cell.alignment     = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[total_row].height = 22


def apply_summary_style(ws):
    C_H = "2C3E50"; C_FG = "FFFFFF"
    C_ODD = "EBF5FB"; C_EVEN = "FFFFFF"
    thin  = Side(style="thin", color="AED6F1")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill      = PatternFill("solid", fgColor=C_H)
        cell.font      = Font(bold=True, color=C_FG, size=10, name="Calibri")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = bdr
    ws.row_dimensions[1].height = 24

    for r_idx in range(2, ws.max_row + 1):
        fill = PatternFill("solid", fgColor=C_ODD if r_idx % 2 == 0 else C_EVEN)
        for cell in ws[r_idx]:
            cell.fill      = fill
            cell.font      = Font(size=9, name="Calibri")
            cell.border    = bdr
            cell.alignment = Alignment(vertical="center")
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
            left=Side(style="medium", color="2C3E50"),
            right=Side(style="medium", color="2C3E50"),
            top=Side(style="medium", color="2C3E50"),
            bottom=Side(style="medium", color="2C3E50"),
        )
        if c_idx == 1:
            tot_cell.value = "TOTALS"
        elif isinstance(sample_cell.value, (int, float)):
            col_l           = get_column_letter(c_idx)
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
    valid  = df[df["Received Amount"].notna()]
    totals = {"Metric": [], "Value": []}
    fields = [
        ("Total Orders",              len(df)),
        ("Orders Calculated",         int(df["Received Amount"].notna().sum())),
        ("Orders NaN (no category)",  int(df["Received Amount"].isna().sum())),
        ("Brand+Cat Charges",         int((df.get("Charge Method", pd.Series(dtype=str)).str.startswith("brand+cat", na=False)).sum())),
        ("Category-only Charges",     int((df.get("Charge Method", pd.Series(dtype=str)).str.startswith("category", na=False)).sum())),
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
    ("sku_corrections", {}),
    ("manual_cat_map", {}),
    ("result_df",      None),
    ("charges_df",     None),
    ("known_brands",   []),
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
if order_files and charges_file:

    with st.spinner("🔄 Reading files…"):
        # ── Load & merge all order files ────────────────────────────────────
        order_df, file_info, file_errors = load_all_order_files(order_files)

        st.sidebar.markdown("---")
        st.sidebar.markdown("**📄 Uploaded Order Files**")
        fi_df = pd.DataFrame(file_info)
        st.sidebar.dataframe(fi_df, hide_index=True, use_container_width=True)
        total_order_rows = len(order_df)
        st.sidebar.caption(f"Total rows loaded: **{total_order_rows:,}**")

        if file_errors:
            for err in file_errors:
                st.warning(f"⚠️ {err}")

        if order_df.empty:
            st.error("❌ No valid order rows could be loaded from the uploaded file(s). Check column names.")
            st.stop()

        xl     = pd.read_excel(charges_file, sheet_name=None, header=None)
        sheets = list(xl.values())

        if len(sheets) < 3:
            st.error(f"❌ Excel must have at least 3 sheets. Found: {list(xl.keys())}")
            st.stop()

        # Sheet 0 — Brand Name + Category + Charge slabs
        charges_df = parse_charges_df(sheets[0])

        # Extract known brands from Sheet 0 for product-name matching
        known_brands = []
        if "Brand Name" in charges_df.columns:
            known_brands = charges_df["Brand Name"].dropna().unique().tolist()
        st.sidebar.success(f"✅ Known brands from Sheet 0: {', '.join(known_brands) if known_brands else 'None'}")

        # Sheet 1 — SKU → Sub-category (+ optional Brand column)
        sku_cat_dict, _ = parse_sku_cat(sheets[1])

        # Sheet 2 — PWN prices
        pwn_dict = parse_pwn_dict(sheets[2])

        # Replace SKU map (optional)
        replace_map = {}
        if replace_sku_file:
            replace_map = parse_replace_map(replace_sku_file)
            st.sidebar.success(f"✅ Replace SKU loaded: {len(replace_map):,} entries")

        charge_cats  = charges_df["Category"].unique().tolist()
        all_sub_cats = sorted(set(v for v in sku_cat_dict.values() if v and v != "nan"))
        cat_map      = build_cat_map(all_sub_cats, charge_cats)
        unmapped     = [sc for sc in all_sub_cats if sc.lower() not in cat_map]

        st.session_state.update({
            "charges_df":   charges_df,
            "known_brands": known_brands,
            "sku_cat_dict": sku_cat_dict,
            "pwn_dict":     pwn_dict,
            "order_df":     order_df,
            "cat_map":      cat_map,
            "charge_cats":  charge_cats,
            "unmapped_cats": unmapped,
            "replace_map":  replace_map,
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
                "entry in **Charges Description**. Pick the correct charges category for each."
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
        known_brands=st.session_state["known_brands"],
        replace_map=st.session_state["replace_map"],
        pwn_overrides=st.session_state["pwn_overrides"],
        sku_corrections=st.session_state["sku_corrections"],
    )
    st.session_state["result_df"] = result_df
    summary_df, cat_df = build_summary(result_df)

    replace_resolved = result_df[result_df["PWN Match"].str.startswith("replace", na=False)]
    brand_cat_charged = result_df[result_df["Charge Method"].str.startswith("brand+cat", na=False)]

    st.success(
        f"✅ Processed **{len(result_df):,}** orders  |  "
        f"**{int(result_df['Received Amount'].notna().sum()):,}** calculated  |  "
        f"**{int(result_df['Received Amount'].isna().sum()):,}** skipped (no category/GT match)"
        + (f"  |  **{len(brand_cat_charged):,}** charged via Brand+Category rates" if len(brand_cat_charged) else "")
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

        # ── SKU Correction Panel ────────────────────────────────────────────
        broken_df = result_df[
            result_df["Received Amount"].isna() | (result_df["PWN Match"] == "not_found")
        ]
        if len(broken_df):
            no_cat = int(broken_df["Received Amount"].isna().sum())
            no_pwn = int((broken_df["PWN Match"] == "not_found").sum())

            with st.expander(
                f"✏️  **{len(broken_df)} SKU(s) have lookup issues — correct SKU name to retry all lookups**",
                expanded=False,
            ):
                st.info(
                    "Type the **corrected SKU** for each row — the tool will re-run "
                    "**all** lookups (sub-category → GT → commission → collection → PWN) "
                    "using the corrected name. Leave blank to keep original."
                )
                st.caption(
                    f"🔴 No category/GT match: **{no_cat}**  |  "
                    f"🟡 No PWN found: **{no_pwn}**"
                )

                broken_skus = broken_df["SKU"].unique().tolist()
                correction_inputs = {}

                h1, h2, h3, h4 = st.columns([3, 2, 2, 3])
                h1.markdown("**Original SKU**")
                h2.markdown("**Issue**")
                h3.markdown("**Corrected SKU**")
                h4.markdown("**Live preview (after save)**")
                st.markdown("---")

                for sku in broken_skus:
                    sku_rows = broken_df[broken_df["SKU"] == sku]
                    issues = []
                    if sku_rows["Received Amount"].isna().any():
                        issues.append("❌ No category/GT")
                    if (sku_rows["PWN Match"] == "not_found").any():
                        issues.append("⚠️ No PWN")
                    issue_str = "  &  ".join(issues)

                    existing_correction = st.session_state["sku_corrections"].get(sku.upper(), "")

                    c1, c2, c3, c4 = st.columns([3, 2, 2, 3])

                    c1.markdown(
                        f"<div style='padding-top:6px;font-size:0.88rem;"
                        f"word-break:break-all'><code>{sku}</code></div>",
                        unsafe_allow_html=True,
                    )
                    c2.markdown(
                        f"<div style='padding-top:6px;font-size:0.82rem'>{issue_str}</div>",
                        unsafe_allow_html=True,
                    )
                    corrected = c3.text_input(
                        "Corrected SKU",
                        value=existing_correction,
                        placeholder="e.g. YK1234-L",
                        label_visibility="collapsed",
                        key=f"sku_corr_{sku}",
                    )
                    correction_inputs[sku] = corrected.strip()

                    if existing_correction:
                        lookup_sku = strip_vendor_prefix(existing_correction)
                        sub_cat_preview = st.session_state["sku_cat_dict"].get(lookup_sku.upper(), "")
                        pwn_v, pwn_m = lookup_pwn_with_replace(
                            lookup_sku,
                            st.session_state["pwn_dict"],
                            st.session_state["replace_map"],
                        )
                        cat_resolved = (
                            st.session_state["cat_map"].get(sub_cat_preview.strip().lower(), "")
                            or st.session_state["manual_cat_map"].get(sub_cat_preview.strip().lower(), "")
                        )
                        status_parts = []
                        if sub_cat_preview and sub_cat_preview != "nan":
                            status_parts.append(f"📦 Sub-cat: *{sub_cat_preview}*")
                        if cat_resolved:
                            status_parts.append(f"🏷️ Charges cat: *{cat_resolved}*")
                        if pd.notna(pwn_v):
                            status_parts.append(f"💰 PWN: ₹{pwn_v:,.2f}")
                        if status_parts:
                            preview_html = (
                                "<div style='padding-top:4px;font-size:0.80rem;color:#1a7a3c;line-height:1.6'>"
                                + "<br>".join(status_parts)
                                + "</div>"
                            )
                        else:
                            preview_html = (
                                "<div style='padding-top:6px;font-size:0.80rem;color:#c0392b'>"
                                "⚠️ Still unresolved — check SKU spelling"
                                "</div>"
                            )
                        c4.markdown(preview_html, unsafe_allow_html=True)
                    else:
                        c4.markdown(
                            "<div style='padding-top:6px;font-size:0.80rem;color:#aaa'>"
                            "— type a correction to preview —"
                            "</div>",
                            unsafe_allow_html=True,
                        )

                st.markdown("---")
                col_save, col_clear = st.columns([2, 1])
                if col_save.button("💾  Save SKU Corrections & Recalculate", type="primary"):
                    new_corrections = {}
                    for orig_sku, corrected_sku in correction_inputs.items():
                        if corrected_sku:
                            new_corrections[orig_sku.upper()] = corrected_sku
                    st.session_state["sku_corrections"] = new_corrections
                    st.rerun()
                if col_clear.button("🗑️  Clear All Corrections"):
                    st.session_state["sku_corrections"] = {}
                    st.rerun()

        if st.session_state["sku_corrections"]:
            with st.expander(
                f"✅  **{len(st.session_state['sku_corrections'])} SKU correction(s) active** — click to view/manage",
                expanded=False,
            ):
                corr_rows = [
                    {"Original SKU": orig, "→ Corrected SKU": corr}
                    for orig, corr in st.session_state["sku_corrections"].items()
                ]
                st.dataframe(pd.DataFrame(corr_rows), use_container_width=True, hide_index=True)
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
        k10.metric(
            "Net Difference",
            f"₹{net:,.2f}",
            delta=f"{'▲' if net >= 0 else '▼'} {abs(net):,.2f}",
            delta_color="normal" if net >= 0 else "inverse",
        )

        st.markdown("---")

        # ── Filters ─────────────────────────────────────────────────────────
        f1, f2, f3, f4 = st.columns([2, 2, 2, 3])
        all_cats_opt = ["All"] + sorted(result_df["Sub-Category"].dropna().unique().tolist())
        sel_cat   = f1.selectbox("Sub-Category", all_cats_opt)

        all_brands = ["All"] + sorted(result_df["Brand Name"].dropna().unique().tolist())
        sel_brand  = f2.selectbox("Brand", all_brands)

        diff_opt = f3.selectbox("Difference type",
                                ["All", "Positive (+)", "Negative (−)",
                                 "Zero / Matched", "No PWN data", "No Category (NaN)"])
        search = f4.text_input("🔎 Search by SKU or Order ID")

        view = result_df.copy()
        if sel_cat != "All":
            view = view[view["Sub-Category"] == sel_cat]
        if sel_brand != "All":
            view = view[view["Brand Name"] == sel_brand]
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

        # ── Display columns (no Lookup SKU / SKU Corrected / Source File) ──
        display_cols = [
            "Order Id", "SKU", "Product", "Brand Name", "Ordered On",
            "Sub-Category", "Charges Category", "Charge Method",
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
            "ℹ️  **Charge Method:** brand+cat:Brand|Category (brand+category lookup) or category:CategoryName (fallback)  \n"
            "**Received Amount** = Selling Price − Total Charges − GST on Charges − TDS − TCS  \n"
            "**Taxable Value** = Selling Price − (Selling Price / 105 × 5)  \n"
            "**TDS** = Taxable Value × 0.1%  |  **TCS** = Taxable Value × 0.5%  \n"
            "**Difference** = Received Amount − (Qty × PWN)"
        )

        st.markdown("---")
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols = [
            "Order Id", "SKU", "Brand Name", "Sub-Category", "Charges Category", "Charge Method",
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
    st.info("👈 Upload **order file(s)** and the **Data Excel** in the sidebar to begin.")
    st.markdown("""
---
### How it works

| File | Description |
|------|-------------|
| **Order File(s)** | Flipkart Seller Hub export — CSV, XLSX or XLS. Upload **one or more** files and they are merged automatically. Needs columns: `Order Id`, `SKU`, `Product`, `Ordered On`, `Invoice Amount`, `Quantity` |
| **Data Excel** | Yash Gallery workbook — 3 sheets by position (see below) |
| **Replace SKU Excel** *(optional)* | Maps Seller SKU Id → OMS SKU for PWN fallback lookup |

**Excel sheet positions:**

| Position | Sheet Name | Used For |
|----------|-----------|----------|
| Index 0 | Data For Flipkart Roc | Brand Name + Category → Commission / Collection / GT slabs |
| Index 1 | Category Discription | Seller SKU → Sub-category |
| Index 2 | Price We Need | OMS Child SKU → PWN price |

---
### ✨ Brand + Category Charge Lookup

**How it works:**

1. **Brand extracted from Product column** in the Order File  
   - Example: `"Yash Gallery Women Floral Print..."` → Brand = `"Yash Gallery"`
   - Example: `"KALINI Women Fit and Flare..."` → Brand = `"KALINI"`
   - Matches against known brands loaded from Sheet 0 (longest match first)

2. **Charge lookup uses Brand + Category together**  
   - Looks up in Sheet 0 using both Brand Name AND Category  
   - Falls back to Category-only if brand not found in Sheet 0

3. **Sheet 0 format (Data For Flipkart Roc):**

| Brand Name | Category | Lower Limit Commission | Upper Limit Commission | Commission % | Collection Lower | Collection Upper | Collection % | GT Lower | GT Upper | GT Charge |
|------------|----------|----------------------|----------------------|--------------|-----------------|-----------------|--------------|----------|----------|-----------|
| Yash Gallery | kurta | 0 | 500 | 0.09 | 0 | 500 | 0.003 | 0 | 200 | 52 |
| Yash Gallery | gown | 0 | 500 | 0.09 | 0 | 500 | 0.003 | 0 | 200 | 52 |
| KALINI | kurta | 0 | 500 | 0.12 | 0 | 500 | 0.003 | 0 | 200 | 52 |

---
### Calculation per order

```
BRAND LOOKUP:
Brand Name       = Extracted from Product column (matched against Sheet 0 brands)
Charge Method    = "brand+cat:BrandName|Category" if found, else "category:CategoryName"

GT Amount        = Fixed ₹ from GT slab (Invoice Amount → Brand+Category slab)
Selling Price    = Invoice Amount − GT Amount

Commission       = Selling Price × Commission % (from Brand+Category slab)
Collection Fee   = Selling Price × Collection % (from Brand+Category slab)
Total Charges    = Commission + Collection Fee + Fixed Fee

GST              = Total Charges × 18%

Taxable Value    = Selling Price − (Selling Price / 105 × 5)
TDS              = Taxable Value × 0.1%
TCS              = Taxable Value × 0.5%

Received Amount  = Selling Price − Total Charges − GST − TDS − TCS

PWN Benchmark    = Qty × PWN
Difference       = Received Amount − PWN Benchmark
```
""")
