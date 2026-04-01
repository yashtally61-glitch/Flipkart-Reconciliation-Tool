import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Flipkart Reconciliation – Multi-Account",
    layout="wide", page_icon="🧾",
)
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.3rem; font-weight: 700; }
.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

st.title("🧾 Flipkart Reconciliation Tool — Multi-Account")
st.caption("Yash Gallery Private Limited — Tool made by Ashu Bhatt | Finance Team")

# ─── ACCOUNT CONFIG ───────────────────────────────────────────────────────────
ACCOUNTS = {
    "Yash Gallery": {
        "brands": ["Yash Gallery", "KALINI", "Tasrika"],
        "color": "#1A3C5E",
    },
    "Pushpa Enterprises": {
        "brands": ["AKIKO", "Pushpa", "HouseOfCommon"],
        "color": "#1E6B3A",
    },
    "Aashirwad Garments": {
        "brands": ["IKRASS"],
        "color": "#7B2D00",
    },
}
ACCOUNT_NAMES = list(ACCOUNTS.keys())

def get_account_for_brand(brand: str) -> str:
    for acc, cfg in ACCOUNTS.items():
        for b in cfg["brands"]:
            if brand.lower() == b.lower():
                return acc
    return "Unmatched"

# ─── KNOWN BRANDS (all) ───────────────────────────────────────────────────────
ALL_KNOWN_BRANDS = [b for cfg in ACCOUNTS.values() for b in cfg["brands"]]

def extract_brand_from_product(product: str) -> str:
    if not product or str(product).strip().lower() == "nan":
        return ""
    p = str(product).strip()
    for brand in ALL_KNOWN_BRANDS:
        if p.lower().startswith(brand.lower()):
            return brand
    return ""

# ─── SIZE EXPAND ──────────────────────────────────────────────────────────────
SIZE_EXPAND = {
    "L-XL":["L","XL"],"S-M":["S","M"],"XXL-3XL":["XXL","3XL"],
    "F-S/XXL":["F"],"F-3xl/5xl":["F"],"XS-S":["XS","S"],
    "M-L":["M","L"],"XL-XXL":["XL","XXL"],"3XL-4XL":["3XL","4XL"],
    "5XL-6XL":["5XL","6XL"],"7XL-8XL":["7XL","8XL"],
    "4XL-5XL":["4XL","5XL"],"2XL-3XL":["2XL","3XL"],
    "XS-S-M":["XS","S","M"],"L-XL-XXL":["L","XL","XXL"],
}
ORDER_TO_PRICE_SIZE: dict = defaultdict(list)
for _ps, _os_list in SIZE_EXPAND.items():
    for _os in _os_list:
        ORDER_TO_PRICE_SIZE[_os.upper()].append(_ps)

VENDOR_PREFIXES = ["GWN-","GWN_","GWN","SPF-","SPF_","SPF","KL_","KL-","KL"]

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def strip_vendor_prefix(sku: str) -> str:
    upper = sku.upper()
    for p in VENDOR_PREFIXES:
        if upper.startswith(p.upper()):
            sku = sku[len(p):]
            break
    sku = re.sub(r"(?i)YKN","YK",sku)
    sku = re.sub(r"(?i)YKC","YK",sku)
    return sku

def get_sku_base(sku: str) -> str:
    key = re.sub(r"\s*-\s*","-",sku.strip().upper())
    parts = key.rsplit("-",1)
    return parts[0] if len(parts)==2 else key

def lookup_sub_cat(raw_sku: str, sku_info_dict: dict) -> tuple:
    key_raw = raw_sku.strip().upper()
    info = sku_info_dict.get(key_raw)
    if info and info.get("sub_cat") and str(info["sub_cat"]).lower() != "nan":
        return info["sub_cat"], "exact"
    stripped = strip_vendor_prefix(raw_sku).strip().upper()
    if stripped != key_raw:
        info2 = sku_info_dict.get(stripped)
        if info2 and info2.get("sub_cat") and str(info2["sub_cat"]).lower() != "nan":
            return info2["sub_cat"], "exact-stripped"
    base_raw = get_sku_base(key_raw)
    for csk, ci in sku_info_dict.items():
        if get_sku_base(csk)==base_raw and ci.get("sub_cat") and str(ci["sub_cat"]).lower()!="nan":
            return ci["sub_cat"], f"base({csk})"
    if stripped != key_raw:
        base_str = get_sku_base(stripped)
        for csk, ci in sku_info_dict.items():
            if get_sku_base(csk)==base_str and ci.get("sub_cat") and str(ci["sub_cat"]).lower()!="nan":
                return ci["sub_cat"], f"base-stripped({csk})"
    return "", "not_found"

def lookup_pwn(sku: str, pwn_dict: dict) -> tuple:
    key = sku.strip().upper()
    val = pwn_dict.get(key)
    if val is not None and pd.notna(val): return float(val),"direct"
    parts = key.rsplit("-",1)
    if len(parts)==2:
        base,size = parts
        for ps in ORDER_TO_PRICE_SIZE.get(size,[]):
            val = pwn_dict.get(f"{base}-{ps.upper()}")
            if val is not None and pd.notna(val): return float(val),f"size-expand({ps})"
    base = get_sku_base(key)
    if base:
        for csk,cv in pwn_dict.items():
            if get_sku_base(csk)==base and pd.notna(cv): return float(cv),f"base-match({csk})"
    return np.nan,"not_found"

def lookup_pwn_with_replace(sku, pwn_dict, replace_map):
    v,m = lookup_pwn(sku, pwn_dict)
    if m!="not_found": return v,m
    oms = replace_map.get(sku.strip().upper())
    if oms:
        v2,m2 = lookup_pwn(oms, pwn_dict)
        if m2!="not_found": return v2,f"replace→{m2}"
    return np.nan,"not_found"

# ─── SLAB LOOKUPS ─────────────────────────────────────────────────────────────
def _filter_brand_cat(charges_df, brand, cat):
    if not brand or not cat: return pd.DataFrame()
    if "Brand Name" not in charges_df.columns or "Category" not in charges_df.columns:
        return pd.DataFrame()
    bn = str(brand).strip().lower(); cn = str(cat).strip().lower()
    if bn=="nan" or cn=="nan": return pd.DataFrame()
    bm = charges_df["Brand Name"].fillna("").astype(str).str.strip().str.lower()==bn
    cm = charges_df["Category"].fillna("").astype(str).str.strip().str.lower()==cn
    return charges_df[bm & cm].copy()

def lookup_gt(brand, cat, inv_amount, charges_df):
    for _,r in _filter_brand_cat(charges_df,brand,cat).iterrows():
        lo,hi,gt = r.get("GT Lower Limit"),r.get("GT Upper Limit"),r.get("GT Charge")
        if pd.isna(lo) or pd.isna(hi) or pd.isna(gt): continue
        try:
            if float(lo)<=inv_amount<=float(hi)+0.99: return float(gt)
        except: continue
    return np.nan

def lookup_commission(brand, cat, sell_price, charges_df):
    for _,r in _filter_brand_cat(charges_df,brand,cat).iterrows():
        lo,hi,ch = r.get("Lower Limit Commision"),r.get("Upper Limit Commision"),r.get("Commision Charge")
        if pd.isna(lo) or pd.isna(hi) or pd.isna(ch): continue
        try:
            if float(lo)<=sell_price<=float(hi)+0.99: return round(float(ch)*sell_price,5)
        except: continue
    return np.nan

def lookup_collection(brand, cat, sell_price, charges_df):
    for _,r in _filter_brand_cat(charges_df,brand,cat).iterrows():
        lo_raw,hi,cf = r.get("Collection Lower Limit"),r.get("Collection Upper Limit"),r.get("Collection Charge")
        if pd.isna(hi) or pd.isna(cf): continue
        try:
            cf_val = float(cf) if pd.notna(cf) else 0.0
            lo_val = 0.0 if (pd.isna(lo_raw) or str(lo_raw).strip().startswith(">")) else float(lo_raw)
            if lo_val<sell_price<=float(hi)+0.99: return round(cf_val*sell_price,5)
        except: continue
    return np.nan

# ─── PARSERS ──────────────────────────────────────────────────────────────────
def parse_charges_df(raw):
    df = raw.copy()
    # Auto-detect header row: look for a row containing "Brand Name" or "Category"
    header_row = 0
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.tolist()]
        if "brand name" in vals or "category" in vals:
            header_row = i
            break
    df.columns = [str(c).strip() for c in df.iloc[header_row].tolist()]
    df = df.iloc[header_row+1:].reset_index(drop=True)
    if "Brand Name" in df.columns: df["Brand Name"] = df["Brand Name"].ffill()
    if "Category"   in df.columns: df["Category"]   = df["Category"].ffill()
    df = df[df["Category"].notna()].copy()
    # Strip whitespace from all string columns
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    for col in ["Lower Limit Commision","Upper Limit Commision","Commision Charge",
                "Collection Lower Limit","Collection Upper Limit",
                "GT Lower Limit","GT Upper Limit","GT Charge"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col],errors="coerce")
    if "Collection Charge" in df.columns:
        df["Collection Charge"] = (df["Collection Charge"].astype(str)
            .str.replace("₹","",regex=False).str.strip()
            .pipe(pd.to_numeric,errors="coerce"))
    return df

def parse_sku_info(raw):
    df = raw.copy()
    df.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)
    sku_info = {}
    for _,row in df.iterrows():
        sku     = str(row.get("Seller SKU Id","")).strip().upper()
        sub_cat = str(row.get("Sub-category","")).strip()
        brand   = str(row.get("Brand","")).strip() if "Brand" in df.columns else ""
        if sku:
            sku_info[sku] = {"sub_cat":sub_cat,"brand":brand}
    return sku_info

def parse_pwn_dict(raw):
    df = raw.copy()
    df.columns = [str(c).strip() for c in raw.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)
    df["OMS Child SKU"] = df["OMS Child SKU"].astype(str).str.strip()
    df["PWN+10%+50"] = pd.to_numeric(df["PWN+10%+50"],errors="coerce")
    return dict(zip(df["OMS Child SKU"].str.upper(),df["PWN+10%+50"]))

def parse_replace_map(file):
    xl = pd.read_excel(file,header=None)
    df = xl.copy()
    df.columns = [str(c).strip() for c in xl.iloc[0].tolist()]
    df = df.iloc[1:].reset_index(drop=True)
    return dict(zip(df["Seller SKU Id"].astype(str).str.strip().str.upper(),
                    df["OMS SKU"].astype(str).str.strip().str.upper()))

# ─── FILE READER ──────────────────────────────────────────────────────────────
REQUIRED_ORDER_COLS = {"Order Id","SKU","Invoice Amount","Quantity","Product"}

def read_order_file(f):
    name = f.name.lower()
    try:
        if name.endswith(".csv"):
            raw = f.read()
            for enc in ("utf-8","utf-8-sig","latin-1","cp1252"):
                try: df = pd.read_csv(BytesIO(raw),encoding=enc); break
                except UnicodeDecodeError: continue
            else: return pd.DataFrame(), f"Could not decode '{f.name}'."
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            df = pd.read_excel(f, engine="openpyxl" if name.endswith(".xlsx") else "xlrd")
        else: return pd.DataFrame(), f"Unsupported: '{f.name}'"
        df.columns = [str(c).strip() for c in df.columns]
        missing = REQUIRED_ORDER_COLS - set(df.columns)
        if missing: return pd.DataFrame(), f"'{f.name}' missing: {', '.join(sorted(missing))}"
        df["_source_file"] = f.name
        return df, ""
    except Exception as e:
        return pd.DataFrame(), f"Error reading '{f.name}': {e}"

def load_all_order_files(files):
    frames,file_info,errors = [],[],[]
    for f in files:
        df,err = read_order_file(f)
        if err: errors.append(err); file_info.append({"File":f.name,"Rows":0,"Status":"❌ Error"})
        else:   frames.append(df); file_info.append({"File":f.name,"Rows":len(df),"Status":"✅ OK"})
    combined = pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
    return combined, file_info, errors

# ─── CORE RECONCILIATION ──────────────────────────────────────────────────────
TDS_RATE = 0.001
TCS_RATE = 0.005

def run_reconciliation(order_df, charges_df, sku_info_dict, pwn_dict,
                       fixed_fee, gst_rate, replace_map=None,
                       pwn_overrides=None, sku_corrections=None):
    replace_map     = replace_map     or {}
    pwn_overrides   = pwn_overrides   or {}
    sku_corrections = sku_corrections or {}
    rows_out = []

    for _, row in order_df.iterrows():
        raw_sku    = str(row.get("SKU","")).strip()
        product    = str(row.get("Product","")).strip()
        order_id   = str(row.get("Order Id","")).strip()
        ordered_on = row.get("Ordered On","")
        inv_amount = float(row.get("Invoice Amount",0) or 0)
        quantity   = int(row.get("Quantity",1) or 1)

        corrected_raw = sku_corrections.get(raw_sku.upper(), raw_sku)
        if corrected_raw.strip().upper() in replace_map:
            corrected_raw = replace_map[corrected_raw.strip().upper()]

        brand_name = extract_brand_from_product(product)
        sub_cat, cat_match_note = lookup_sub_cat(corrected_raw, sku_info_dict)
        cat = sub_cat.strip() if sub_cat and str(sub_cat).lower() != "nan" else ""
        sku_for_pwn = strip_vendor_prefix(corrected_raw)

        # ── initialise ALL output vars so nothing is ever unbound ──
        gt_val = sell_price = commission = coll_fee = np.nan
        total_charges = gst_on_charges = taxable_value = np.nan
        tds = tcs = total_deductions = received_amount = np.nan
        charge_method = "not_found"

        if not brand_name:
            charge_method = "no_brand_in_product"
        elif not cat:
            charge_method = f"no_subcat|brand={brand_name}|sku={corrected_raw}"
        else:
            gt_val = lookup_gt(brand_name, cat, inv_amount, charges_df)
            if pd.isna(gt_val):
                charge_method = f"gt_slab_missing|{brand_name}|{cat}|inv={inv_amount}"
            else:
                sell_price = round(inv_amount - gt_val, 5)
                commission = lookup_commission(brand_name, cat, sell_price, charges_df)
                coll_fee   = lookup_collection(brand_name, cat, sell_price, charges_df)
                if pd.isna(commission) or pd.isna(coll_fee):
                    charge_method = f"comm_or_coll_missing|{brand_name}|{cat}|sp={sell_price}"
                    gt_val = sell_price = commission = coll_fee = np.nan
                else:
                    commission       = float(commission)
                    coll_fee         = float(coll_fee)
                    total_charges    = round(commission + coll_fee + float(fixed_fee), 5)
                    gst_on_charges   = round(total_charges * gst_rate, 5)
                    taxable_value    = round(sell_price - (sell_price / 105 * 5), 5)
                    tds              = round(taxable_value * TDS_RATE, 5)
                    tcs              = round(taxable_value * TCS_RATE, 5)
                    total_deductions = round(total_charges + gst_on_charges + tds + tcs, 5)
                    received_amount  = round(sell_price - total_charges - gst_on_charges - tds - tcs, 5)
                    charge_method    = f"{brand_name} | {cat}" 

        pwn_val, match_method = lookup_pwn_with_replace(sku_for_pwn, pwn_dict, replace_map)
        if sku_for_pwn.upper() in pwn_overrides:
            pwn_val, match_method = float(pwn_overrides[sku_for_pwn.upper()]), "manual"
        # manual_pwn: keyed by stripped or original SKU
        _mpwn = manual_pwn.get(sku_for_pwn.strip().upper()) or manual_pwn.get(raw_sku.strip().upper())
        if _mpwn is not None:
            try: pwn_val, match_method = float(_mpwn), "manual-pwn"
            except: pass

        full_match_note = match_method
        if cat_match_note and cat_match_note not in ("exact","exact-stripped"):
            full_match_note = f"{match_method} | cat:{cat_match_note}"

        if pd.notna(received_amount) and pd.notna(pwn_val):
            pwn_benchmark = round(pwn_val * quantity, 5)
            difference    = round(received_amount - pwn_benchmark, 5)
        else:
            pwn_benchmark = difference = np.nan

        rows_out.append({
            "Order Id": order_id,
            "SKU":raw_sku, "Product":product, "Brand Name":brand_name,
            "Ordered On":ordered_on, "Sub-Category":sub_cat,
            "Charge Method":charge_method, "Qty":quantity,
            "Invoice Amount":inv_amount, "GT (As Per Calc)":gt_val,
            "Selling Price":sell_price, "Commission":commission,
            "Collection Fee":coll_fee, "Fixed Fee":float(fixed_fee),
            "Total Charges":total_charges, "GST on Charges":gst_on_charges,
            "Taxable Value":taxable_value, "TDS":tds, "TCS":tcs,
            "Total Deductions":total_deductions, "Received Amount":received_amount,
            "PWN":pwn_val, "PWN Benchmark":pwn_benchmark,
            "PWN Match":full_match_note, "Difference":difference,
        })

    return pd.DataFrame(rows_out)

# ─── FORMATTING ───────────────────────────────────────────────────────────────
MONEY_COLS = ["Invoice Amount","GT (As Per Calc)","Selling Price","Commission",
              "Collection Fee","Fixed Fee","Total Charges","GST on Charges",
              "Taxable Value","TDS","TCS","Total Deductions","Received Amount",
              "PWN","PWN Benchmark","Difference"]

def fmt_inr(x):
    try:
        if pd.isna(x): return "—"
        return f"₹{float(x):,.2f}"
    except: return str(x)

def style_table(df, diff_col="Difference"):
    fmt_dict = {c: fmt_inr for c in df.columns if c in MONEY_COLS}
    def colour_diff(val):
        try:
            v = float(val)
            if v<0: return "color: red; font-weight: bold"
            if v>0: return "color: green; font-weight: bold"
        except: pass
        return ""
    styler = df.style.format(fmt_dict)
    if diff_col in df.columns:
        styler = styler.map(colour_diff, subset=[diff_col])
    return styler

# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
def apply_roc_sheet_style(ws, df):
    C_HEADER_BG="1A3C5E"; C_HEADER_FG="FFFFFF"
    C_ALT1="EAF2FB"; C_ALT2="FFFFFF"
    C_GREEN_BG="D6EFDD"; C_RED_BG="FDDEDE"; C_ZERO_BG="FFF9E6"
    C_TOTAL_BG="1A3C5E"; C_TOTAL_FG="FFD700"; C_BORDER="B0C4D8"
    thin  = Side(style="thin",   color=C_BORDER)
    thick = Side(style="medium", color="1A3C5E")
    bdr        = Border(left=thin, right=thin, top=thin, bottom=thin)
    bdr_header = Border(left=thick,right=thick,top=thick,bottom=thick)
    money_names = set(MONEY_COLS)
    cols = df.columns.tolist()
    C = {name: get_column_letter(i+1) for i,name in enumerate(cols)}
    col_widths = {
        "Order Id":20,"SKU":28,"Product":40,"Brand Name":18,"Ordered On":14,
        "Sub-Category":20,"Charge Method":28,"Qty":6,"Invoice Amount":15,
        "GT (As Per Calc)":15,"Selling Price":15,"Commission":14,"Collection Fee":15,
        "Fixed Fee":10,"Total Charges":15,"GST on Charges":15,"Taxable Value":14,
        "TDS":10,"TCS":10,"Total Deductions":16,"Received Amount":16,
        "PWN":12,"PWN Benchmark":15,"PWN Match":16,"Difference":14,
    }
    for i,cn in enumerate(cols,start=1):
        ws.column_dimensions[get_column_letter(i)].width = col_widths.get(cn,14)
    for cell in ws[1]:
        cell.fill=PatternFill("solid",fgColor=C_HEADER_BG)
        cell.font=Font(bold=True,color=C_HEADER_FG,size=10,name="Calibri")
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.border=bdr_header
    ws.row_dimensions[1].height=30
    diff_col_idx = cols.index("Difference")+1 if "Difference" in cols else None

    def row_formulas(r):
        sp=C.get("Selling Price",""); inv=C.get("Invoice Amount","")
        gt=C.get("GT (As Per Calc)",""); qty=C.get("Qty","")
        com=C.get("Commission",""); colf=C.get("Collection Fee","")
        ff=C.get("Fixed Fee",""); tc=C.get("Total Charges","")
        gst=C.get("GST on Charges",""); tv=C.get("Taxable Value","")
        tds=C.get("TDS",""); tcs=C.get("TCS","")
        td=C.get("Total Deductions",""); ra=C.get("Received Amount","")
        pwn=C.get("PWN",""); pb=C.get("PWN Benchmark","")
        diff=C.get("Difference","")
        fmls={}
        if sp and inv and gt:   fmls["Selling Price"]=f'=IF(OR({gt}{r}="",{gt}{r}=0),"",ROUND({inv}{r}-{gt}{r},2))'
        if tc and com and colf: fmls["Total Charges"]=f'=IF({sp}{r}="","",ROUND({com}{r}+{colf}{r}+{ff}{r},2))'
        if gst and tc:          fmls["GST on Charges"]=f'=IF({tc}{r}="","",ROUND({tc}{r}*0.18,2))'
        if tv and sp:           fmls["Taxable Value"]=f'=IF({sp}{r}="","",ROUND({sp}{r}-{sp}{r}/105*5,2))'
        if tds and tv:          fmls["TDS"]=f'=IF({tv}{r}="","",ROUND({tv}{r}*0.001,2))'
        if tcs and tv:          fmls["TCS"]=f'=IF({tv}{r}="","",ROUND({tv}{r}*0.005,2))'
        if td and tc:           fmls["Total Deductions"]=f'=IF({tc}{r}="","",ROUND({tc}{r}+{gst}{r}+{tds}{r}+{tcs}{r},2))'
        if ra and sp and tc:    fmls["Received Amount"]=f'=IF({sp}{r}="","",ROUND({sp}{r}-{tc}{r}-{gst}{r}-{tds}{r}-{tcs}{r},2))'
        if pb and pwn and qty:  fmls["PWN Benchmark"]=f'=IF({pwn}{r}="","",ROUND({pwn}{r}*{qty}{r},2))'
        if diff and ra and pb:  fmls["Difference"]=f'=IF(OR({ra}{r}="",{pb}{r}=""),"",ROUND({ra}{r}-{pb}{r},2))'
        return fmls

    for r_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        alt_fill = PatternFill("solid", fgColor=C_ALT1 if r_idx%2==0 else C_ALT2)
        fmls = row_formulas(r_idx)
        for c_idx,(col_name,val) in enumerate(zip(cols,row_data),start=1):
            cell = ws.cell(row=r_idx,column=c_idx)
            if col_name in fmls: cell.value = fmls[col_name]
            else: cell.value = None if (isinstance(val,float) and np.isnan(val)) else val
            cell.border=bdr; cell.font=Font(size=9,name="Calibri"); cell.fill=alt_fill
            if col_name in money_names:
                cell.number_format='₹#,##0.00'
                cell.alignment=Alignment(horizontal="right",vertical="center")
            elif col_name=="Qty":
                cell.alignment=Alignment(horizontal="center",vertical="center")
            else:
                cell.alignment=Alignment(horizontal="left",vertical="center")
            if c_idx==diff_col_idx:
                try:
                    v=float(val)
                    if not np.isnan(v):
                        if v<0:   cell.fill=PatternFill("solid",fgColor=C_RED_BG);   cell.font=Font(color="C0392B",bold=True,size=9,name="Calibri")
                        elif v>0: cell.fill=PatternFill("solid",fgColor=C_GREEN_BG); cell.font=Font(color="1E8449",bold=True,size=9,name="Calibri")
                        else:     cell.fill=PatternFill("solid",fgColor=C_ZERO_BG);  cell.font=Font(color="7D6608",bold=True,size=9,name="Calibri")
                except: pass
        ws.row_dimensions[r_idx].height=16

    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    last_data_row=len(df)+1; total_row=last_data_row+2
    for c_idx,col_name in enumerate(cols,start=1):
        cell=ws.cell(row=total_row,column=c_idx)
        cell.fill=PatternFill("solid",fgColor=C_TOTAL_BG)
        cell.font=Font(bold=True,color=C_TOTAL_FG,size=10,name="Calibri")
        cell.border=bdr_header
        if c_idx==1: cell.value="TOTALS"; cell.alignment=Alignment(horizontal="left",vertical="center")
        elif col_name in money_names:
            col_l=get_column_letter(c_idx)
            cell.value=f"=SUM({col_l}2:{col_l}{last_data_row})"
            cell.number_format='₹#,##0.00'
            cell.alignment=Alignment(horizontal="right",vertical="center")
    ws.row_dimensions[total_row].height=22

def apply_summary_style(ws):
    thin=Side(style="thin",color="AED6F1")
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    for cell in ws[1]:
        cell.fill=PatternFill("solid",fgColor="2C3E50")
        cell.font=Font(bold=True,color="FFFFFF",size=10,name="Calibri")
        cell.alignment=Alignment(horizontal="center",vertical="center")
        cell.border=bdr
    ws.row_dimensions[1].height=24
    for r_idx in range(2,ws.max_row+1):
        fill=PatternFill("solid",fgColor="EBF5FB" if r_idx%2==0 else "FFFFFF")
        for cell in ws[r_idx]:
            cell.fill=fill; cell.font=Font(size=9,name="Calibri")
            cell.border=bdr; cell.alignment=Alignment(vertical="center")
        ws.row_dimensions[r_idx].height=15
    for col_cells in ws.columns:
        width=max((len(str(c.value or "")) for c in col_cells),default=10)
        ws.column_dimensions[col_cells[0].column_letter].width=min(width+4,40)
    last_row=ws.max_row; total_row=last_row+2
    for c_idx in range(1,ws.max_column+1):
        sc=ws.cell(row=2,column=c_idx); tc=ws.cell(row=total_row,column=c_idx)
        tc.fill=PatternFill("solid",fgColor="2C3E50")
        tc.font=Font(bold=True,color="FFD700",size=10,name="Calibri")
        tc.border=Border(left=Side(style="medium",color="2C3E50"),right=Side(style="medium",color="2C3E50"),
                         top=Side(style="medium",color="2C3E50"),bottom=Side(style="medium",color="2C3E50"))
        if c_idx==1: tc.value="TOTALS"
        elif isinstance(sc.value,(int,float)):
            cl=get_column_letter(c_idx)
            tc.value=f"=SUM({cl}2:{cl}{last_row})"
            tc.number_format='₹#,##0.00'
            tc.alignment=Alignment(horizontal="right",vertical="center")
    ws.freeze_panes="A2"

def to_excel_multi(recon_df, summary_df, cat_df, sheet_prefix=""):
    buf=BytesIO()
    sname = lambda s: f"{sheet_prefix[:10]+' ' if sheet_prefix else ''}{s}"[:31]
    with pd.ExcelWriter(buf,engine="openpyxl") as writer:
        recon_df.to_excel(writer,index=False,sheet_name=sname("Reconciliation"))
        cat_df.to_excel(writer,index=False,sheet_name=sname("Category Breakdown"))
        summary_df.to_excel(writer,index=False,sheet_name=sname("Charges Summary"))
        apply_roc_sheet_style(writer.sheets[sname("Reconciliation")],recon_df)
        apply_summary_style(writer.sheets[sname("Category Breakdown")])
        apply_summary_style(writer.sheets[sname("Charges Summary")])
    return buf.getvalue()

def to_excel_unmatched(df):
    buf=BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as writer:
        df.to_excel(writer,index=False,sheet_name="Unmatched Orders")
        apply_summary_style(writer.sheets["Unmatched Orders"])
    return buf.getvalue()

# ─── SUMMARY ──────────────────────────────────────────────────────────────────
def build_summary(df):
    valid=df[df["Received Amount"].notna()]
    totals={"Metric":[],"Value":[]}
    for label,val in [
        ("Total Orders",len(df)),
        ("Orders Calculated",int(df["Received Amount"].notna().sum())),
        ("Orders NaN (no match)",int(df["Received Amount"].isna().sum())),
        ("Total Invoice Amount",df["Invoice Amount"].sum()),
        ("Total GT (As Per Calc)",valid["GT (As Per Calc)"].sum()),
        ("Total Selling Price",valid["Selling Price"].sum()),
        ("Total Commission",valid["Commission"].sum()),
        ("Total Collection Fee",valid["Collection Fee"].sum()),
        ("Total Fixed Fee",valid["Fixed Fee"].sum()),
        ("Total Charges (C+F+Fixed)",valid["Total Charges"].sum()),
        ("Total GST on Charges",valid["GST on Charges"].sum()),
        ("Total TDS",valid["TDS"].sum()),
        ("Total TCS",valid["TCS"].sum()),
        ("Total Deductions",valid["Total Deductions"].sum()),
        ("Total Received Amount",valid["Received Amount"].sum()),
        ("Net Difference vs PWN",valid["Difference"].sum()),
        ("Orders with -ve Diff",int((valid["Difference"]<0).sum())),
        ("Orders with +ve Diff",int((valid["Difference"]>0).sum())),
        ("Orders – No PWN found",int(df["Difference"].isna().sum())),
        ("Avg Received per Order",valid["Received Amount"].mean()),
        ("Avg Difference per Order",valid["Difference"].mean()),
    ]:
        totals["Metric"].append(label)
        totals["Value"].append(round(val,2) if isinstance(val,float) else val)
    summary_df=pd.DataFrame(totals)
    if valid.empty:
        cat_df=pd.DataFrame(columns=["Sub-Category","Orders","Invoice Total","GT Total",
            "Selling Total","Commission","Collection Fee","Fixed Fee","Total Charges",
            "GST Total","TDS Total","TCS Total","Total Deductions","Received Total",
            "Net Difference","Avg Difference"])
        return summary_df, cat_df
    cat_df=(
        valid.groupby("Sub-Category")
        .agg(Orders=("Order Id","count"),
             Invoice_Total=("Invoice Amount","sum"),
             GT_Total=("GT (As Per Calc)","sum"),
             Selling_Total=("Selling Price","sum"),
             Commission=("Commission","sum"),
             Collection=("Collection Fee","sum"),
             Fixed=("Fixed Fee","sum"),
             Total_Charges=("Total Charges","sum"),
             GST_Total=("GST on Charges","sum"),
             TDS_Total=("TDS","sum"),
             TCS_Total=("TCS","sum"),
             Deductions=("Total Deductions","sum"),
             Received_Total=("Received Amount","sum"),
             Net_Diff=("Difference","sum"),
             Avg_Diff=("Difference","mean"))
        .reset_index().sort_values("Invoice_Total",ascending=False).round(2)
    )
    cat_df.columns=["Sub-Category","Orders","Invoice Total","GT Total","Selling Total",
                    "Commission","Collection Fee","Fixed Fee","Total Charges",
                    "GST Total","TDS Total","TCS Total","Total Deductions",
                    "Received Total","Net Difference","Avg Difference"]
    return summary_df, cat_df

# ─── DISPLAY COLS ─────────────────────────────────────────────────────────────
DISPLAY_COLS = ["Order Id","SKU","Product","Brand Name","Ordered On","Sub-Category",
                "Charge Method","Qty","Invoice Amount","GT (As Per Calc)","Selling Price",
                "Commission","Collection Fee","Fixed Fee","Total Charges","GST on Charges",
                "Taxable Value","TDS","TCS","Total Deductions","Received Amount",
                "PWN","PWN Benchmark","Difference","PWN Match"]

# ─── RENDER ONE ACCOUNT TAB ───────────────────────────────────────────────────
def render_account_tab(acc_name, result_df, summary_df, cat_df, fixed_fee, gst_rate,
                       sku_info_dict, pwn_dict, replace_map, charges_df=None):
    if result_df is None or result_df.empty:
        st.info(f"No orders processed yet for **{acc_name}**. Upload Order file and Data Excel in the sidebar.")
        return

    valid = result_df[result_df["Received Amount"].notna()]

    # ── Diagnostic expander ──────────────────────────────────────────────────
    nan_count = int(result_df["Received Amount"].isna().sum())
    calc_count = int(result_df["Received Amount"].notna().sum())
    if nan_count > 0:
        with st.expander(f"🔍 **Diagnostic: {nan_count} orders NOT calculated** — click to investigate", expanded=True):
            reason_counts = result_df[result_df["Received Amount"].isna()]["Charge Method"].value_counts()
            st.markdown("**Why are rows blank? (Charge Method breakdown)**")
            st.dataframe(reason_counts.reset_index().rename(columns={"index":"Reason","Charge Method":"Count"}),
                         use_container_width=True, hide_index=True)
            st.markdown("---")
            st.markdown("**Sample of failed rows (first 10):**")
            fail_cols = ["Order Id","SKU","Product","Brand Name","Sub-Category","Charge Method","Invoice Amount"]
            fc = [c for c in fail_cols if c in result_df.columns]
            st.dataframe(result_df[result_df["Received Amount"].isna()][fc].head(10),
                         use_container_width=True, hide_index=True)
            # Show what brands + categories exist in the charges sheet
            if charges_df is not None and not charges_df.empty:
                brands_in_sheet = sorted(charges_df["Brand Name"].dropna().unique().tolist()) if "Brand Name" in charges_df.columns else []
                cats_in_sheet   = sorted(charges_df["Category"].dropna().unique().tolist())   if "Category"   in charges_df.columns else []
                st.markdown(f"**Brands in your Charges sheet (Sheet 0):** `{'` | `'.join(brands_in_sheet) or 'NONE FOUND'}`")
                st.markdown(f"**Categories in your Charges sheet (Sheet 0):** `{'` | `'.join(cats_in_sheet) or 'NONE FOUND'}`")
                st.markdown("**Charges sheet preview (first 20 rows):**")
                st.dataframe(charges_df.head(20), use_container_width=True)
            st.info(
                "**How to fix:**\n"
                "- `no_brand_in_product` → Product name doesn't start with a known brand. Fix the Product column or add the brand prefix.\n"
                "- `no_subcat|brand=...` → SKU not found in Sheet 1. Check SKU spelling or add it to the Sub-category sheet.\n"
                "- `gt_slab_missing|...` → Invoice Amount not covered by any GT slab in Sheet 0. Check your GT slab ranges.\n"
                "- `comm_or_coll_missing|...` → Selling Price not in Commission/Collection slab range."
            )
    else:
        st.success(f"✅ All **{calc_count:,}** orders calculated successfully.")

    # ── Missing Sub-Category / PWN panel ────────────────────────────────────
    missing_subcat = result_df[result_df["Sub-Category"].fillna("").str.strip() == ""]
    missing_pwn    = result_df[result_df["PWN Match"] == "not_found"]
    missing_either = pd.concat([missing_subcat, missing_pwn]).drop_duplicates(subset=["SKU"])

    if not missing_either.empty:
        with st.expander(
            f"✏️ **Fix Missing Data — {len(missing_either)} unique SKU(s) missing Sub-Category or PWN**",
            expanded=False
        ):
            st.markdown("### Fix Missing Sub-Category & PWN Values")
            st.caption(
                "For each SKU below: type the correct **Sub-Category** and/or **PWN** value. "
                "Save to recalculate. You can also re-upload corrected Sheet 1 or Sheet 2 Excel files below."
            )

            # ── Manual entry table ────────────────────────────────────────────
            st.markdown("#### ✏️ Manual Entry")
            existing_sc  = st.session_state["manual_sub_cats"].get(acc_name, {})
            existing_pwn = st.session_state["manual_pwn"].get(acc_name, {})
            new_sc_vals  = {}
            new_pwn_vals = {}

            hc1,hc2,hc3,hc4,hc5 = st.columns([3,2,2,2,2])
            hc1.markdown("**SKU**"); hc2.markdown("**Issue**")
            hc3.markdown("**Current Sub-Cat**"); hc4.markdown("**Enter Sub-Category**"); hc5.markdown("**Enter PWN (₹)**")
            st.markdown("---")

            for _, mrow in missing_either.iterrows():
                sku = str(mrow["SKU"]).strip()
                issues = []
                if str(mrow.get("Sub-Category","")).strip() in ("","nan"): issues.append("❌ No Sub-Cat")
                if mrow.get("PWN Match","") == "not_found":                issues.append("⚠️ No PWN")

                c1,c2,c3,c4,c5 = st.columns([3,2,2,2,2])
                c1.markdown(f"`{sku}`")
                c2.markdown(" & ".join(issues))
                c3.markdown(f"`{mrow.get('Sub-Category','—')}`")
                new_sc  = c4.text_input("Sub-Cat",  value=existing_sc.get(sku.upper(),""),
                                         placeholder="e.g. Kurta", label_visibility="collapsed",
                                         key=f"msc_{acc_name}_{sku}")
                new_pwn = c5.text_input("PWN",      value=str(existing_pwn.get(sku.upper(),"")) if existing_pwn.get(sku.upper()) else "",
                                         placeholder="e.g. 250.00", label_visibility="collapsed",
                                         key=f"mpwn_{acc_name}_{sku}")
                new_sc_vals[sku.upper()]  = new_sc.strip()
                new_pwn_vals[sku.upper()] = new_pwn.strip()

            st.markdown("---")
            btn_save, btn_clear = st.columns([2,1])
            if btn_save.button(f"💾 Save & Recalculate — {acc_name}", key=f"save_manual_{acc_name}", type="primary"):
                st.session_state["manual_sub_cats"][acc_name] = {k:v for k,v in new_sc_vals.items() if v}
                st.session_state["manual_pwn"][acc_name]      = {}
                for k,v in new_pwn_vals.items():
                    if v:
                        try: st.session_state["manual_pwn"][acc_name][k] = float(v)
                        except: st.warning(f"PWN value for {k} is not a valid number: '{v}'")
                # Recalculate this account immediately
                _od = st.session_state["order_dfs"].get(acc_name)
                _cd = st.session_state["charges_dfs"].get(acc_name)
                _si = st.session_state["sku_info_dicts"].get(acc_name, {})
                _pd = st.session_state["pwn_dicts"].get(acc_name, {})
                if _od is not None and _cd is not None:
                    _r = run_reconciliation(
                        _od, _cd, _si, _pd, fixed_fee, gst_rate,
                        replace_map=st.session_state["replace_map"],
                        pwn_overrides=st.session_state["pwn_overrides"],
                        sku_corrections=st.session_state["sku_corrections"],
                        manual_sub_cats=st.session_state["manual_sub_cats"].get(acc_name, {}),
                        manual_pwn=st.session_state["manual_pwn"].get(acc_name, {}),
                    )
                    _s, _c = build_summary(_r)
                    st.session_state["results"][acc_name]   = _r
                    st.session_state["summaries"][acc_name] = _s
                    st.session_state["cat_dfs"][acc_name]   = _c
                st.rerun()
            if btn_clear.button(f"🗑️ Clear Manual Entries — {acc_name}", key=f"clear_manual_{acc_name}"):
                st.session_state["manual_sub_cats"][acc_name] = {}
                st.session_state["manual_pwn"][acc_name]      = {}
                st.rerun()

            # ── Re-upload corrected sheets ────────────────────────────────────
            st.markdown("---")
            st.markdown("#### 📤 Re-upload Corrected Data Sheets")
            st.caption(
                "Upload a corrected **Sheet 1** (SKU → Sub-category) or **Sheet 2** (OMS Child SKU → PWN+10%+50) "
                "to merge new rows into the existing lookup tables without replacing your other data."
            )
            ru1, ru2 = st.columns(2)
            new_sheet1 = ru1.file_uploader(
                "Re-upload Sheet 1 (Sub-category)", type=["xlsx"],
                key=f"reup_sheet1_{acc_name}"
            )
            new_sheet2 = ru2.file_uploader(
                "Re-upload Sheet 2 (PWN)", type=["xlsx"],
                key=f"reup_sheet2_{acc_name}"
            )
            if new_sheet1 or new_sheet2:
                if st.button(f"🔄 Merge & Recalculate — {acc_name}", key=f"merge_sheets_{acc_name}", type="primary"):
                    if new_sheet1:
                        try:
                            raw1 = pd.read_excel(new_sheet1, header=None)
                            new_info = parse_sku_info(raw1)
                            merged_info = {**st.session_state["sku_info_dicts"].get(acc_name, {}), **new_info}
                            st.session_state["sku_info_dicts"][acc_name] = merged_info
                            st.success(f"✅ Sheet 1 merged: {len(new_info):,} SKU entries added/updated.")
                        except Exception as e:
                            st.error(f"Sheet 1 error: {e}")
                    if new_sheet2:
                        try:
                            raw2 = pd.read_excel(new_sheet2, header=None)
                            new_pwn = parse_pwn_dict(raw2)
                            merged_pwn = {**st.session_state["pwn_dicts"].get(acc_name, {}), **new_pwn}
                            st.session_state["pwn_dicts"][acc_name] = merged_pwn
                            st.success(f"✅ Sheet 2 merged: {len(new_pwn):,} PWN entries added/updated.")
                        except Exception as e:
                            st.error(f"Sheet 2 error: {e}")
                    # Recalculate after merge
                    _od = st.session_state["order_dfs"].get(acc_name)
                    _cd = st.session_state["charges_dfs"].get(acc_name)
                    _si = st.session_state["sku_info_dicts"].get(acc_name, {})
                    _pd = st.session_state["pwn_dicts"].get(acc_name, {})
                    if _od is not None and _cd is not None:
                        _r = run_reconciliation(
                            _od, _cd, _si, _pd, fixed_fee, gst_rate,
                            replace_map=st.session_state["replace_map"],
                            pwn_overrides=st.session_state["pwn_overrides"],
                            sku_corrections=st.session_state["sku_corrections"],
                            manual_sub_cats=st.session_state["manual_sub_cats"].get(acc_name, {}),
                            manual_pwn=st.session_state["manual_pwn"].get(acc_name, {}),
                        )
                        _s, _c = build_summary(_r)
                        st.session_state["results"][acc_name]   = _r
                        st.session_state["summaries"][acc_name] = _s
                        st.session_state["cat_dfs"][acc_name]   = _c
                    st.rerun()

    # ── Metrics row ──────────────────────────────────────────────────────────
    st.markdown("### 📊 Summary")
    k1,k2,k3,k4,k5,k6,k7,k8,k9,k10 = st.columns(10)
    k1.metric("Orders", f"{len(result_df):,}")
    k2.metric("Invoice Total", f"₹{result_df['Invoice Amount'].sum():,.0f}")
    k3.metric("GT Total", f"₹{valid['GT (As Per Calc)'].sum():,.0f}")
    k4.metric("Selling Total", f"₹{valid['Selling Price'].sum():,.0f}")
    k5.metric("Total Charges", f"₹{valid['Total Charges'].sum():,.0f}")
    k6.metric("GST on Charges", f"₹{valid['GST on Charges'].sum():,.0f}")
    k7.metric("Total TDS", f"₹{valid['TDS'].sum():,.2f}")
    k8.metric("Total TCS", f"₹{valid['TCS'].sum():,.2f}")
    k9.metric("Received Total", f"₹{valid['Received Amount'].sum():,.0f}")
    net = valid["Difference"].sum()
    k10.metric("Net Difference", f"₹{net:,.2f}",
               delta=f"{'▲' if net>=0 else '▼'} {abs(net):,.2f}",
               delta_color="normal" if net>=0 else "inverse")
    st.markdown("---")

    # ── Sub-tabs ─────────────────────────────────────────────────────────────
    t1, t2, t3 = st.tabs(["📋 Reconciliation", "💰 Charges Summary", "📊 Category Breakdown"])

    with t1:
        f1,f2,f3,f4 = st.columns([2,2,2,3])
        sel_cat   = f1.selectbox("Sub-Category", ["All"]+sorted(result_df["Sub-Category"].dropna().unique().tolist()), key=f"cat_{acc_name}")
        sel_brand = f2.selectbox("Brand", ["All"]+sorted(result_df["Brand Name"].dropna().unique().tolist()), key=f"brand_{acc_name}")
        diff_opt  = f3.selectbox("Difference type", ["All","Positive (+)","Negative (−)","Zero / Matched","No PWN data","No Category (NaN)"], key=f"diff_{acc_name}")
        search    = f4.text_input("🔎 Search by SKU or Order ID", key=f"search_{acc_name}")

        view = result_df.copy()
        if sel_cat   !="All": view=view[view["Sub-Category"]==sel_cat]
        if sel_brand !="All": view=view[view["Brand Name"]==sel_brand]
        if diff_opt=="Positive (+)":    view=view[view["Difference"]>0]
        elif diff_opt=="Negative (−)":  view=view[view["Difference"]<0]
        elif diff_opt=="Zero / Matched":view=view[view["Difference"]==0]
        elif diff_opt=="No PWN data":   view=view[view["PWN Match"]=="not_found"]
        elif diff_opt=="No Category (NaN)": view=view[view["Received Amount"].isna()]
        if search.strip():
            mask=(view["SKU"].str.contains(search.strip(),case=False,na=False)|
                  view["Order Id"].str.contains(search.strip(),case=False,na=False))
            view=view[mask]

        st.caption(f"Showing **{len(view):,}** of **{len(result_df):,}** orders")
        cols_show = [c for c in DISPLAY_COLS if c in view.columns]
        st.dataframe(style_table(view[cols_show], diff_col="Difference"), use_container_width=True, height=500)

        st.markdown("### 📥 Download")
        d1,d2 = st.columns(2)
        d1.download_button(
            "⬇  Full Reconciliation (Excel – 3 sheets, styled)",
            data=to_excel_multi(result_df[cols_show], summary_df, cat_df, sheet_prefix=acc_name[:10]),
            file_name=f"recon_{acc_name.replace(' ','_').lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_full_{acc_name}")
        d2.download_button(
            "⬇  Filtered View (Excel, styled)",
            data=to_excel_multi(view[cols_show].reset_index(drop=True), summary_df, cat_df, sheet_prefix=acc_name[:10]),
            file_name=f"recon_{acc_name.replace(' ','_').lower()}_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_filt_{acc_name}")

    with t2:
        st.markdown("### 💰 Total Charges Summary")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("#### 📤 Flipkart Deductions")
            a1,a2=st.columns(2)
            a1.metric("Commission",     f"₹{valid['Commission'].sum():,.2f}")
            a2.metric("Collection Fee", f"₹{valid['Collection Fee'].sum():,.2f}")
            a1.metric("Fixed Fee",      f"₹{valid['Fixed Fee'].sum():,.2f}")
            a2.metric("GST on Charges", f"₹{valid['GST on Charges'].sum():,.2f}")
            a1.metric("TDS (0.1%)",     f"₹{valid['TDS'].sum():,.2f}")
            a2.metric("TCS (0.5%)",     f"₹{valid['TCS'].sum():,.2f}")
            st.metric("🔴 Total Deductions", f"₹{valid['Total Deductions'].sum():,.2f}")
        with col_b:
            st.markdown("#### 📥 What You Receive")
            b1,b2=st.columns(2)
            b1.metric("Total Invoice", f"₹{result_df['Invoice Amount'].sum():,.2f}")
            b2.metric("GT Total",      f"₹{valid['GT (As Per Calc)'].sum():,.2f}")
            b1.metric("Selling Total", f"₹{valid['Selling Price'].sum():,.2f}")
            b2.metric("Total Received",f"₹{valid['Received Amount'].sum():,.2f}")
            net2=valid["Difference"].sum()
            b1.metric("Net Diff vs PWN",f"₹{net2:,.2f}",
                      delta=f"{'▲' if net2>=0 else '▼'} {abs(net2):,.2f}",
                      delta_color="normal" if net2>=0 else "inverse")
            b2.metric("Orders −ve Diff",int((valid["Difference"]<0).sum()))
        st.markdown("---")
        st.markdown("#### 📋 Per-Order Charges Detail")
        charge_cols=["Order Id","SKU","Brand Name","Sub-Category","Charge Method",
                     "Invoice Amount","GT (As Per Calc)","Selling Price","Commission","Collection Fee",
                     "Fixed Fee","Total Charges","GST on Charges","Taxable Value","TDS","TCS",
                     "Total Deductions","Received Amount"]
        cc = [c for c in charge_cols if c in result_df.columns]
        st.dataframe(style_table(result_df[cc]), use_container_width=True, height=480)
        st.markdown("#### 🧾 Grand Summary Table")
        st.dataframe(summary_df, use_container_width=True)

    with t3:
        st.markdown("### 📊 Sub-Category-wise Breakdown")
        cat_money=[c for c in cat_df.columns if c not in ("Sub-Category","Orders")]
        if not cat_df.empty:
            st.dataframe(style_table(cat_df, diff_col="Net Difference")
                         .format({c:"₹{:.2f}" for c in cat_money}),
                         use_container_width=True)
            st.markdown("---")
            st.markdown("#### 🔢 Charge Components (per Sub-Category)")
            comp_cols=["Sub-Category","Orders","GT Total","Commission","Collection Fee",
                       "Fixed Fee","GST Total","TDS Total","TCS Total","Total Deductions"]
            comp_money=[c for c in comp_cols if c not in ("Sub-Category","Orders")]
            st.dataframe(cat_df[comp_cols].style.format({c:"₹{:.2f}" for c in comp_money}),
                         use_container_width=True)
        else:
            st.info("No calculated orders to break down.")

# ─── SESSION STATE ────────────────────────────────────────────────────────────
defaults = {
    "pwn_overrides": {},
    "sku_corrections": {},
    "replace_map": {},
    # per-account manual overrides  {acc: {SKU.upper(): value}}
    "manual_sub_cats": {acc: {} for acc in ACCOUNT_NAMES},
    "manual_pwn":      {acc: {} for acc in ACCOUNT_NAMES},
    # per-account results
    "results": {acc: None for acc in ACCOUNT_NAMES},
    "summaries": {acc: None for acc in ACCOUNT_NAMES},
    "cat_dfs": {acc: None for acc in ACCOUNT_NAMES},
    "unmatched_df": None,
    # per-account data files
    "charges_dfs": {acc: None for acc in ACCOUNT_NAMES},
    "sku_info_dicts": {acc: {} for acc in ACCOUNT_NAMES},
    "pwn_dicts": {acc: {} for acc in ACCOUNT_NAMES},
    # per-account order dfs
    "order_dfs": {acc: None for acc in ACCOUNT_NAMES},
}
for k,v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Upload Files")

    st.markdown("#### 1️⃣ Yash Gallery")
    yg_orders  = st.file_uploader("Order File(s) – Yash Gallery",  type=["csv","xlsx","xls"], accept_multiple_files=True, key="up_ord_yg")
    yg_data    = st.file_uploader("Data Excel – Yash Gallery",     type=["xlsx"], key="up_data_yg")

    st.markdown("#### 2️⃣ Pushpa Enterprises")
    pe_orders  = st.file_uploader("Order File(s) – Pushpa Enterprises", type=["csv","xlsx","xls"], accept_multiple_files=True, key="up_ord_pe")
    pe_data    = st.file_uploader("Data Excel – Pushpa Enterprises",    type=["xlsx"], key="up_data_pe")

    st.markdown("#### 3️⃣ Aashirwad Garments")
    ag_orders  = st.file_uploader("Order File(s) – Aashirwad Garments", type=["csv","xlsx","xls"], accept_multiple_files=True, key="up_ord_ag")
    ag_data    = st.file_uploader("Data Excel – Aashirwad Garments",    type=["xlsx"], key="up_data_ag")

    st.markdown("#### 🔄 Shared")
    replace_sku_file = st.file_uploader("Replace SKU Excel (shared, optional)", type=["xlsx"], key="up_replace")

    st.markdown("---")
    st.subheader("⚙️ Settings")
    fixed_fee = st.number_input("Fixed Fee per order (₹)", value=5,  min_value=0, step=1)
    gst_rate  = st.number_input("GST on charges (%)",      value=18, min_value=0, step=1) / 100

    st.markdown("---")
    st.subheader("✏️ Global SKU Corrections")
    st.caption("Apply corrections to SKUs across all accounts before recalculating.")

    all_results = [st.session_state["results"].get(a) for a in ACCOUNT_NAMES]
    all_dfs     = [r for r in all_results if r is not None and not r.empty]
    if all_dfs:
        combined_broken = pd.concat([
            r[r["Received Amount"].isna() | (r["PWN Match"]=="not_found")]
            for r in all_dfs
        ]).drop_duplicates(subset=["SKU"])
        broken_skus = combined_broken["SKU"].unique().tolist()
        if broken_skus:
            correction_inputs = {}
            for sku in broken_skus:
                existing = st.session_state["sku_corrections"].get(sku.upper(),"")
                corrected = st.text_input(
                    f"SKU: {sku}", value=existing,
                    placeholder="Corrected SKU",
                    key=f"global_corr_{sku}"
                )
                correction_inputs[sku] = corrected.strip()

            if st.button("💾 Save Corrections & Recalculate", type="primary"):
                st.session_state["sku_corrections"] = {
                    o.upper(): c for o, c in correction_inputs.items() if c
                }
                st.rerun()
            if st.button("🗑️ Clear All Corrections"):
                st.session_state["sku_corrections"] = {}
                st.rerun()
        else:
            st.success("✅ No broken SKUs found across all accounts.")
    else:
        st.info("Run reconciliation first to see SKU issues here.")

# ─── PROCESS UPLOADS ──────────────────────────────────────────────────────────
ACCOUNT_UPLOADS = {
    "Yash Gallery":        (yg_orders,  yg_data),
    "Pushpa Enterprises":  (pe_orders,  pe_data),
    "Aashirwad Garments":  (ag_orders,  ag_data),
}

# Load replace map
if replace_sku_file:
    try:
        st.session_state["replace_map"] = parse_replace_map(replace_sku_file)
        st.sidebar.success(f"✅ Replace SKU loaded: {len(st.session_state['replace_map']):,} entries")
    except Exception as e:
        st.sidebar.error(f"Replace SKU error: {e}")

replace_map = st.session_state["replace_map"]

any_loaded = False
for acc_name, (order_files, data_file) in ACCOUNT_UPLOADS.items():
    if order_files and data_file:
        any_loaded = True
        with st.spinner(f"⏳ Processing {acc_name}…"):
            order_df, file_info, file_errors = load_all_order_files(order_files)
            for err in file_errors:
                st.warning(f"⚠️ [{acc_name}] {err}")
            if order_df.empty:
                st.error(f"❌ [{acc_name}] No valid order rows loaded.")
                continue

            # Parse Data Excel
            try:
                xl = pd.read_excel(data_file, sheet_name=None, header=None)
                sheets = list(xl.values())
                if len(sheets) < 3:
                    st.error(f"❌ [{acc_name}] Data Excel needs ≥3 sheets. Found: {list(xl.keys())}")
                    continue
                charges_df    = parse_charges_df(sheets[0])
                sku_info_dict = parse_sku_info(sheets[1])
                pwn_dict      = parse_pwn_dict(sheets[2])
            except Exception as e:
                st.error(f"❌ [{acc_name}] Error reading Data Excel: {e}")
                continue

            st.session_state["charges_dfs"][acc_name]    = charges_df
            st.session_state["sku_info_dicts"][acc_name] = sku_info_dict
            st.session_state["pwn_dicts"][acc_name]      = pwn_dict
            st.session_state["order_dfs"][acc_name]      = order_df

            result_df = run_reconciliation(
                order_df, charges_df, sku_info_dict, pwn_dict,
                fixed_fee, gst_rate,
                replace_map=replace_map,
                pwn_overrides=st.session_state["pwn_overrides"],
                sku_corrections=st.session_state["sku_corrections"],
                manual_sub_cats=st.session_state["manual_sub_cats"].get(acc_name, {}),
                manual_pwn=st.session_state["manual_pwn"].get(acc_name, {}),
            )
            summary_df, cat_df = build_summary(result_df)
            st.session_state["results"][acc_name]  = result_df
            st.session_state["summaries"][acc_name] = summary_df
            st.session_state["cat_dfs"][acc_name]   = cat_df

# ─── UNMATCHED: merge re-uploaded corrections ─────────────────────────────────
# Build unmatched from all results (brand = "" or "Unmatched")
all_result_dfs = [r for r in st.session_state["results"].values() if r is not None and not r.empty]
if all_result_dfs:
    combined_all = pd.concat(all_result_dfs, ignore_index=True)
    unmatched_df = combined_all[combined_all["Brand Name"]==""].copy()
    st.session_state["unmatched_df"] = unmatched_df
else:
    unmatched_df = st.session_state.get("unmatched_df") or pd.DataFrame()

# ─── MAIN TABS ────────────────────────────────────────────────────────────────
if any_loaded or any(r is not None for r in st.session_state["results"].values()):
    tab_yg, tab_pe, tab_ag, tab_um = st.tabs([
        "🏷️ Yash Gallery",
        "🌸 Pushpa Enterprises",
        "🧵 Aashirwad Garments",
        f"⚠️ Unmatched ({len(unmatched_df) if unmatched_df is not None else 0})",
    ])

    with tab_yg:
        render_account_tab(
            "Yash Gallery",
            st.session_state["results"]["Yash Gallery"],
            st.session_state["summaries"]["Yash Gallery"],
            st.session_state["cat_dfs"]["Yash Gallery"],
            fixed_fee, gst_rate,
            st.session_state["sku_info_dicts"]["Yash Gallery"],
            st.session_state["pwn_dicts"]["Yash Gallery"],
            replace_map,
            charges_df=st.session_state["charges_dfs"].get("Yash Gallery"),
        )

    with tab_pe:
        render_account_tab(
            "Pushpa Enterprises",
            st.session_state["results"]["Pushpa Enterprises"],
            st.session_state["summaries"]["Pushpa Enterprises"],
            st.session_state["cat_dfs"]["Pushpa Enterprises"],
            fixed_fee, gst_rate,
            st.session_state["sku_info_dicts"]["Pushpa Enterprises"],
            st.session_state["pwn_dicts"]["Pushpa Enterprises"],
            replace_map,
            charges_df=st.session_state["charges_dfs"].get("Pushpa Enterprises"),
        )

    with tab_ag:
        render_account_tab(
            "Aashirwad Garments",
            st.session_state["results"]["Aashirwad Garments"],
            st.session_state["summaries"]["Aashirwad Garments"],
            st.session_state["cat_dfs"]["Aashirwad Garments"],
            fixed_fee, gst_rate,
            st.session_state["sku_info_dicts"]["Aashirwad Garments"],
            st.session_state["pwn_dicts"]["Aashirwad Garments"],
            replace_map,
            charges_df=st.session_state["charges_dfs"].get("Aashirwad Garments"),
        )

    with tab_um:
        st.markdown("### ⚠️ Unmatched Orders")
        st.caption("These orders had no recognized brand prefix in their Product name.")

        if unmatched_df is not None and not unmatched_df.empty:
            st.error(f"**{len(unmatched_df):,}** unmatched orders found.")
            ucols = [c for c in DISPLAY_COLS if c in unmatched_df.columns]
            st.dataframe(unmatched_df[ucols], use_container_width=True, height=400)

            st.markdown("#### 📥 Step 1 – Download Unmatched Orders")
            st.download_button(
                "⬇  Download Unmatched Orders (Excel)",
                data=to_excel_unmatched(unmatched_df[ucols]),
                file_name="unmatched_orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_unmatched"
            )

            st.markdown("---")
            st.markdown("#### 📤 Step 2 – Re-upload Corrected File")
            st.info(
                "Fix the **Product** column in the downloaded file so the brand prefix is recognized "
                "(e.g. `Yash Gallery`, `KALINI`, `Tasrika`, `AKIKO`, `Pushpa`, `HouseOfCommon`, `IKRASS`), "
                "then re-upload below. The corrected orders will be merged back into the correct account tab."
            )
            corrected_upload = st.file_uploader(
                "Upload Corrected Unmatched File (CSV / XLSX)",
                type=["csv","xlsx","xls"],
                key="up_corrected_unmatched"
            )
            if corrected_upload:
                corr_df, corr_err = read_order_file(corrected_upload)
                if corr_err:
                    st.error(f"Error reading corrected file: {corr_err}")
                else:
                    st.success(f"✅ Corrected file loaded: **{len(corr_df):,}** rows")
                    # Re-run reconciliation for each corrected row using correct account's data
                    merge_counts = defaultdict(int)
                    for acc_name in ACCOUNT_NAMES:
                        acc_brands_lower = [b.lower() for b in ACCOUNTS[acc_name]["brands"]]
                        acc_rows = corr_df[corr_df["Product"].apply(
                            lambda p: any(str(p).strip().lower().startswith(b) for b in acc_brands_lower)
                        )]
                        if acc_rows.empty:
                            continue
                        charges_df    = st.session_state["charges_dfs"].get(acc_name)
                        sku_info_dict = st.session_state["sku_info_dicts"].get(acc_name, {})
                        pwn_dict      = st.session_state["pwn_dicts"].get(acc_name, {})
                        if charges_df is None:
                            st.warning(f"⚠️ No Data Excel loaded for **{acc_name}** — cannot merge corrected rows.")
                            continue
                        new_result = run_reconciliation(
                            acc_rows.reset_index(drop=True),
                            charges_df, sku_info_dict, pwn_dict,
                            fixed_fee, gst_rate,
                            replace_map=replace_map,
                            pwn_overrides=st.session_state["pwn_overrides"],
                            sku_corrections=st.session_state["sku_corrections"],
                            manual_sub_cats=st.session_state["manual_sub_cats"].get(acc_name, {}),
                            manual_pwn=st.session_state["manual_pwn"].get(acc_name, {}),
                        )
                        existing = st.session_state["results"].get(acc_name)
                        if existing is not None and not existing.empty:
                            merged = pd.concat([existing, new_result], ignore_index=True)
                        else:
                            merged = new_result
                        summary_df, cat_df = build_summary(merged)
                        st.session_state["results"][acc_name]   = merged
                        st.session_state["summaries"][acc_name] = summary_df
                        st.session_state["cat_dfs"][acc_name]   = cat_df
                        merge_counts[acc_name] += len(new_result)

                    still_unmatched = corr_df[corr_df["Product"].apply(
                        lambda p: extract_brand_from_product(str(p)) == ""
                    )]
                    for acc_name, cnt in merge_counts.items():
                        st.success(f"✅ Merged **{cnt}** corrected orders into **{acc_name}**")
                    if not still_unmatched.empty:
                        st.warning(f"⚠️ **{len(still_unmatched)}** rows still unmatched after correction.")
                    if merge_counts:
                        st.button("🔄 Refresh to see updated tabs", on_click=st.rerun)
        else:
            st.success("🎉 No unmatched orders! All products were recognized.")

else:
    st.info("👈 Upload **Order files** and **Data Excel** for each account in the sidebar to begin.")
    st.markdown("""
---
### How it works

| File | Description |
|------|-------------|
| **Order File(s)** | Flipkart Seller Hub export per account (CSV / XLSX / XLS) |
| **Data Excel** | Per-account workbook — 3 sheets (Charges, SKU Info, PWN) |
| **Replace SKU Excel** *(optional, shared)* | Seller SKU → OMS SKU mapping |

---
### ✅ Account → Brand Mapping

| Account | Product name starts with |
|---|---|
| **Yash Gallery** | `Yash Gallery`, `KALINI`, `Tasrika` |
| **Pushpa Enterprises** | `AKIKO`, `Pushpa`, `HouseOfCommon` |
| **Aashirwad Garments** | `IKRASS` |

---
### ✅ Calculation Logic

**Step 1** — Brand from Product name  
**Step 2** — Sub-category from Sheet 1 (SKU lookup)  
**Step 3** — Slab lookups from Sheet 0 using Brand + Sub-category

| Charge | Input |
|--------|-------|
| GT | Invoice Amount |
| Commission | Selling Price (= Invoice − GT) |
| Collection | Selling Price |

**Step 4** — Final amounts:
```
Total Charges   = Commission + Collection Fee + Fixed Fee
GST on Charges  = Total Charges × 18%
Taxable Value   = Selling Price − (Selling Price / 105 × 5)
TDS             = Taxable Value × 0.1%
TCS             = Taxable Value × 0.5%
Received Amount = Selling Price − Total Charges − GST − TDS − TCS
Difference      = Received Amount − (Qty × PWN)
```
""")
