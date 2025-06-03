import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
import logging

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("TRSM_Cost_Adjustment")

# --- Streamlit Config & UI Styling ---
st.set_page_config(page_title="TRSM Cost Adjustment Tool", layout="wide")
st.markdown("""
    <style>
    .title    { text-align: center; color: #0e1e33; font-size: 2.8em; font-weight: 800;}
    .subtitle { text-align: center; color: #234; font-size: 1.3em; margin-bottom: 1.5em;}
    .header   { color: #1a1a1a; font-size: 1.4em; margin-top: 1.2em; font-weight: 700;}
    .footer   { text-align: center; color: #888; font-size: 1.1em; margin-top: 3em;}
    .note     { color: #345; font-size: 1em;}
    .stButton>button { background-color: #1d72b8; color: white; border-radius: 6px; padding: 0.7em 1.4em; font-weight: 700;}
    .stButton>button:hover { background-color: #164e7a; }
    .summary-good { color: green; font-weight: bold; }
    .summary-bad  { color: #d00; }
    </style>
""", unsafe_allow_html=True)

VERSION = "1.3.1"

# --- Constants ---
DEFAULT_RECOVERY = 1.0
DEFAULT_TRIM = 0.0
DEFAULT_LABOUR = 0.0
DEFAULT_STICKER = 0.0
FREIGHT_RATES = {
    "PRAN": 0.07,
    "Local Pickup": 0.11,
    "Alberta": 0.205,
    "Ontario/Quebec": 0.25
}
DEFAULT_FREIGHT = 0.0
DEFAULT_BASE_MARGIN = 0.17
DEFAULT_LIST_MARGIN = 0.25

try:
    import xlsxwriter
    EXCEL_ENGINE = "xlsxwriter"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

# --- Helper Functions ---
def clean_trsm_code(code) -> str:
    code_str = str(code).strip()
    try:
        val = float(code_str.replace(",", ""))
        if val.is_integer():
            code_str = str(int(val))
        else:
            code_str = str(val)
    except Exception:
        code_str = code_str.replace(",", "")
    return code_str

def safe_float(value, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default

def get_freight_cost(vendor: str) -> float:
    vendor = str(vendor).strip()
    if not vendor:
        return DEFAULT_FREIGHT
    if "PRAN" in vendor:
        return FREIGHT_RATES["PRAN"]
    elif "Local" in vendor or "Pickup" in vendor:
        return FREIGHT_RATES["Local Pickup"]
    elif "Alberta" in vendor:
        return FREIGHT_RATES["Alberta"]
    elif "Ontario" in vendor or "Quebec" in vendor:
        return FREIGHT_RATES["Ontario/Quebec"]
    return DEFAULT_FREIGHT

def calculate_actual_inv_cost(vendor_invoice_price: float, lb_per_billing_uom: float) -> float:
    return vendor_invoice_price if lb_per_billing_uom == 0 else vendor_invoice_price / lb_per_billing_uom

def calculate_market_cost(actual_inv_cost: float, adj: float) -> float:
    return actual_inv_cost + adj

def calculate_landed_cost(market_cost: float, freight: float) -> float:
    return market_cost + freight

def calculate_recovery_input(market_cost: float, freight: float, recovery_percent: float) -> float:
    return (market_cost + freight) / recovery_percent if recovery_percent != 0 else 0.0

def compute_totals_for_inputs(row: pd.Series) -> pd.Series:
    for i in range(1, 5):
        qty = safe_float(row.get(f"Qty-{i}"), 0.0)
        unit_cost = safe_float(row.get(f"Unit $-{i}"), 0.0)
        row[f"Total $-{i}"] = qty * unit_cost
    return row

def calculate_raw_material_per_lb_cost(row: pd.Series) -> float:
    return sum(safe_float(row.get(f"Total $-{i}"), 0.0) for i in range(1, 5))

def calculate_waste_output(raw_material_cost: float, recovery_percent: float) -> float:
    return (raw_material_cost / recovery_percent - raw_material_cost) if recovery_percent != 0 else 0.0

def calculate_recovery(trim_cost_lb: float, trim_percent: float, recovery_percent: float) -> float:
    return (trim_cost_lb * trim_percent) / recovery_percent if recovery_percent != 0 else 0.0

def calculate_price_from_margin(cost: float, margin_percent: float) -> float:
    if margin_percent >= 1:
        logger.warning(f"Margin % ≥ 100% ({margin_percent}), using cost.")
        return cost
    return cost / (1 - margin_percent)

def calculate_margin_dollars(base_price: float, final_cost: float) -> float:
    return base_price - final_cost

def auto_fill_missing_columns(df: pd.DataFrame, required_cols=None, verbose=True) -> pd.DataFrame:
    df.columns = [col.strip().replace('\n', ' ') for col in df.columns]
    fills = []
    # 1. Market Cost = Actual Inv Cost(lb) + Adj
    if "Market Cost" not in df.columns and all(col in df.columns for col in ["Actual Inv Cost(lb)", "Adj"]):
        df["Market Cost"] = df["Actual Inv Cost(lb)"] + df["Adj"]
        fills.append("Market Cost")
    # 2. Landed Cost = Market Cost + Freight
    if "Landed Cost" not in df.columns and all(col in df.columns for col in ["Market Cost", "Freight"]):
        df["Landed Cost"] = df["Market Cost"] + df["Freight"]
        fills.append("Landed Cost")
    # 3. Recovery Input = (Market Cost + Freight) / Recovery %
    if ("Recovery %" in df.columns and "Recovery Input" not in df.columns and 
        all(col in df.columns for col in ["Market Cost", "Freight"])):
        recovery = df["Recovery %"].replace(0, np.nan).apply(lambda x: x/100 if x > 1 else x)
        df["Recovery Input"] = (df["Market Cost"] + df["Freight"]) / recovery
        fills.append("Recovery Input")
    # 4. Raw Material Per LB Cost = sum Total $-1 to Total $-4
    total_cols = [f"Total $-{i}" for i in range(1, 5)]
    if "Raw Material Per LB Cost" not in df.columns and all(col in df.columns for col in total_cols):
        df["Raw Material Per LB Cost"] = df[total_cols].sum(axis=1)
        fills.append("Raw Material Per LB Cost")
    # 5. Waste Output $ = (Raw Material Per LB Cost / Recovery %) - Raw Material Per LB Cost
    if ("Waste Output $" not in df.columns and
        "Raw Material Per LB Cost" in df.columns and "Recovery %" in df.columns):
        recovery = df["Recovery %"].replace(0, np.nan).apply(lambda x: x/100 if x > 1 else x)
        df["Waste Output $"] = (df["Raw Material Per LB Cost"] / recovery) - df["Raw Material Per LB Cost"]
        fills.append("Waste Output $")
    # 6. Final Cost = Billling UOM Cost + Priced Sticker
    if "Final Cost" not in df.columns and all(col in df.columns for col in ["Billling UOM Cost", "Priced Sticker"]):
        df["Final Cost"] = df["Billling UOM Cost"] + df["Priced Sticker"]
        fills.append("Final Cost")
    # 7. Margin $ = Base Price - Final Cost
    if "Margin $" not in df.columns and all(col in df.columns for col in ["Base Price", "Final Cost"]):
        df["Margin $"] = df["Base Price"] - df["Final Cost"]
        fills.append("Margin $")
    # 8. List Margin % = (List Price - Final Cost) / List Price
    if ("List Margin %" not in df.columns and
        "List Price" in df.columns and "Final Cost" in df.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            df["List Margin %"] = (df["List Price"] - df["Final Cost"]) / df["List Price"]
        fills.append("List Margin %")
    if verbose and fills:
        st.info(f"🔧 Filled missing columns automatically: {', '.join(fills)}")
    still_missing = []
    if required_cols:
        for col in required_cols:
            if col not in df.columns:
                still_missing.append(col)
        if verbose and still_missing:
            st.warning(f"⚠️ These required columns are missing and could not be auto-filled: {', '.join(still_missing)}")
    return df

def read_files(cost_file, export_file, cost_sheet_name: str, export_sheet_name: str):
    try:
        excel_cost = pd.ExcelFile(cost_file)
        df_cost = pd.read_excel(cost_file, sheet_name=cost_sheet_name, engine="openpyxl")
        other_sheets = {
            sheet: excel_cost.parse(sheet)
            for sheet in excel_cost.sheet_names
            if sheet != cost_sheet_name
        }
        df_export = pd.read_excel(export_file, sheet_name=export_sheet_name, engine="openpyxl")
        df_cost.columns = [col.strip().replace('\n', ' ') for col in df_cost.columns]
        df_export.columns = [col.strip().replace('\n', ' ') for col in df_export.columns]
        if "TRSM Code" in df_cost.columns:
            df_cost["TRSM Code"] = df_cost["TRSM Code"].apply(clean_trsm_code)
        if "Product Code" in df_export.columns:
            df_export["Product Code"] = df_export["Product Code"].apply(clean_trsm_code)
        if df_cost.empty or df_export.empty:
            st.error("One or both sheets are empty.")
            return None, None, None
        return df_cost, df_export, other_sheets
    except Exception as e:
        st.error(f"Error reading files: {e}")
        return None, None, None

def validate_columns(df: pd.DataFrame, required_cols: set, sheet_name: str) -> bool:
    available_cols = set(df.columns)
    missing = required_cols - available_cols
    if missing:
        st.error(f"{sheet_name} missing columns: {missing}")
        return False
    return True

def update_cost_row(row: pd.Series, new_cost_price: float = None, original_row: pd.Series = None) -> pd.Series:
    old_vendor_invoice = safe_float(original_row.get("Vendor Invoice Price") if original_row is not None else row.get("Vendor Invoice Price"))
    old_final_cost = safe_float(original_row.get("Final Cost") if original_row is not None else row.get("Final Cost"))
    old_base_price = safe_float(original_row.get("Base Price") if original_row is not None else row.get("Base Price"))
    old_list_price = safe_float(original_row.get("List Price") if original_row is not None else row.get("List Price"))
    row = compute_totals_for_inputs(row)
    lb_per_billing_uom = safe_float(row.get("lb Per Billling UOM", 1), 1)
    vendor_invoice_price = new_cost_price if new_cost_price is not None else safe_float(row.get("Vendor Invoice Price"))
    actual_inv_cost = calculate_actual_inv_cost(vendor_invoice_price, lb_per_billing_uom)
    adj = safe_float(row.get("Adj", 0))
    market_cost = calculate_market_cost(actual_inv_cost, adj)
    freight = get_freight_cost(row.get("Supplier S Name", ""))
    landed_cost = calculate_landed_cost(market_cost, freight)
    raw_recovery_val = safe_float(row.get("Recovery %", DEFAULT_RECOVERY), DEFAULT_RECOVERY)
    if raw_recovery_val > 1.0:
        raw_recovery_val = raw_recovery_val / 100.0
    if raw_recovery_val <= 0.0:
        raw_recovery_val = DEFAULT_RECOVERY
    recovery_percent = raw_recovery_val
    recovery_input = calculate_recovery_input(market_cost, freight, recovery_percent)
    raw_material_cost = calculate_raw_material_per_lb_cost(row)
    waste_output = calculate_waste_output(raw_material_cost, recovery_percent)
    trim_percent = safe_float(row.get("Trim %", DEFAULT_TRIM))
    trim_cost_lb = safe_float(row.get("Trim Cost/LB", 0.0))
    recovery = calculate_recovery(trim_cost_lb, trim_percent, recovery_percent)
    input_cost = recovery_input + raw_material_cost + waste_output
    net_input_cost = input_cost - recovery
    labour = safe_float(row.get("Labour $", DEFAULT_LABOUR))
    sticker = safe_float(row.get("Normal Sticker", DEFAULT_STICKER))
    new_final_cost_lb = net_input_cost + labour + sticker
    column1_value = safe_float(row.get("Column1", 0))
    billling_uom_cost = new_final_cost_lb * lb_per_billing_uom + column1_value
    priced_sticker = safe_float(row.get("Priced Sticker", 0))
    final_cost = billling_uom_cost + priced_sticker
    base_margin_value = safe_float(row.get("Base Margin %", 0.0))
    if base_margin_value == 0.0:
        base_margin_value = DEFAULT_BASE_MARGIN
    if base_margin_value > 1.0:
        base_margin_decimal = base_margin_value / 100.0
    else:
        base_margin_decimal = base_margin_value
    base_price = calculate_price_from_margin(final_cost, base_margin_decimal)
    list_margin_value = safe_float(row.get("List Margin %", 0.0))
    if list_margin_value == 0.0:
        list_margin_value = DEFAULT_LIST_MARGIN
    elif list_margin_value > 1.0:
        list_margin_value = list_margin_value / 100.0
    list_price = calculate_price_from_margin(final_cost, list_margin_value)
    margin_dollars = calculate_margin_dollars(base_price, final_cost)
    row["Price Change Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["Old Vendor Invoice Price"] = old_vendor_invoice
    row["Vendor Invoice Price"] = vendor_invoice_price
    row["Actual Inv Cost(lb)"] = actual_inv_cost
    row["Adj"] = adj
    row["Market Cost"] = market_cost
    row["Freight"] = freight
    row["Landed Cost"] = landed_cost
    row["Recovery %"] = recovery_percent * 100  # as %
    row["Recovery Input"] = recovery_input
    row["Raw Material Per LB Cost"] = raw_material_cost
    row["Waste Output $"] = waste_output
    row["Trim %"] = trim_percent
    row["Trim Cost/LB"] = trim_cost_lb
    row["Recovery"] = recovery
    row["Input Cost"] = input_cost
    row["Net Input Cost"] = net_input_cost
    row["Labour $"] = labour
    row["Normal Sticker"] = sticker
    row["Material + Labour"] = new_final_cost_lb
    row["New Final Cost (Lb)"] = new_final_cost_lb
    row["Column1"] = column1_value
    row["Billling UOM Cost"] = billling_uom_cost
    row["Priced Sticker"] = priced_sticker
    row["Final Cost"] = final_cost
    row["Old Final Cost"] = old_final_cost
    row["Old Base Price"] = old_base_price
    row["Base Price"] = base_price
    row["Old List Price"] = old_list_price
    row["List Price"] = list_price
    row["Margin $"] = margin_dollars
    return row

def update_cost_sheet(df_cost: pd.DataFrame, trsm_code: str, new_cost_price: float):
    df_updated = df_cost.copy()
    trsm_code_clean = clean_trsm_code(trsm_code)
    updated_flag = False
    updated_trsm_codes = set()
    df_updated["TRSM Code"] = df_updated["TRSM Code"].apply(clean_trsm_code)
    for col in ["Old Vendor Invoice Price", "Old Final Cost", "Old Base Price", "Old List Price", "Price Change Date"]:
        if col not in df_updated.columns:
            df_updated[col] = None
    mask_main = df_updated["TRSM Code"] == trsm_code_clean
    if mask_main.any():
        for idx in df_updated[mask_main].index:
            original_row = df_cost.loc[idx]
            df_updated.loc[idx] = update_cost_row(
                row=df_updated.loc[idx],
                new_cost_price=new_cost_price,
                original_row=original_row
            )
        updated_trsm_codes.update(df_updated.loc[mask_main, "TRSM Code"])
        updated_flag = True
    item_cols = [c for c in df_updated.columns if c.startswith("Item-")]
    to_update = True
    iteration = 0
    while to_update and iteration < 10:
        to_update = False
        composite_mask = pd.Series(False, index=df_updated.index)
        for item_col in item_cols:
            unit_col = f"Unit $-{item_col.split('-')[1]}"
            if unit_col not in df_updated.columns:
                continue
            df_updated[item_col] = df_updated[item_col].apply(clean_trsm_code)
            mask_item = df_updated[item_col].isin(updated_trsm_codes)
            if mask_item.any():
                for idx in df_updated[mask_item].index:
                    trsm_code_item = df_updated.loc[idx, item_col]
                    matching_row = df_updated[df_updated["TRSM Code"] == trsm_code_item]
                    if not matching_row.empty:
                        # ←─── UPDATED LINE ────
                        df_updated.loc[idx, unit_col] = safe_float(matching_row.iloc[0]["Net Input Cost"])
                        # ────────────────────────
                composite_mask |= mask_item
                to_update = True
        if composite_mask.any():
            updated_trsm_codes.update(df_updated.loc[composite_mask, "TRSM Code"])
            for idx in df_updated[composite_mask].index:
                original_row = df_cost.loc[idx]
                df_updated.loc[idx] = update_cost_row(
                    row=df_updated.loc[idx],
                    new_cost_price=None,
                    original_row=original_row
                )
            updated_flag = True
        iteration += 1
    return df_updated, updated_flag, updated_trsm_codes

def update_export_sheet(df_export: pd.DataFrame, df_cost_updated: pd.DataFrame, updated_trsm_codes: set):
    df_export_updated = df_export.copy()
    updated_flag = False
    df_export_updated["Product Code"] = df_export_updated["Product Code"].apply(clean_trsm_code)
    for trsm_code in updated_trsm_codes:
        cost_row = df_cost_updated[df_cost_updated["TRSM Code"] == trsm_code]
        if cost_row.empty:
            continue
        final_cost = safe_float(cost_row.iloc[0]["Final Cost"])
        final_base = safe_float(cost_row.iloc[0]["Base Price"])
        final_list = safe_float(cost_row.iloc[0]["List Price"])
        mask_export = df_export_updated["Product Code"] == trsm_code
        if mask_export.any():
            df_export_updated.loc[mask_export, "Cost Price"] = final_cost
            df_export_updated.loc[mask_export, "Base Price"] = final_base
            df_export_updated.loc[mask_export, "Suggested Price"] = final_list
            updated_flag = True
    return df_export_updated, updated_flag

# --- MAIN APP Logic ---
st.markdown(f'<div class="title">TRSM Product Cost Adjustment Tool</div>', unsafe_allow_html=True)
st.markdown(f'<div class="subtitle">Bulk pricing update· Version {VERSION}</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("🔧 Instructions")
    st.markdown("""
    - Upload both the **Cost Sheet** and **Export Sheet** Excel files.
    - If using different sheet names, specify them in the main page.
    - You may update prices by editing the table below, or by uploading a CSV/Excel of new prices.
    - Review all changes, then click **Apply All Cost Changes**.
    - Download both updated sheets as a single ZIP.
    """)

st.markdown('<div class="header">Step 1: Upload Excel Files</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    cost_file = st.file_uploader("Upload Cost Sheet (XLSX)", type=["xlsx"], key="cost")
with col2:
    export_file = st.file_uploader("Upload Export Sheet (XLSX)", type=["xlsx"], key="export")

if cost_file and export_file:
    cost_sheet_name = st.text_input("Cost Sheet Name:", value="Final Copy").strip()
    export_sheet_name = st.text_input("Export Sheet Name:", value="AllProducts").strip()
    df_cost, df_export, other_sheets = read_files(cost_file, export_file, cost_sheet_name, export_sheet_name)
    if df_cost is None or df_export is None:
        st.stop()
    cost_required = {
        "TRSM Code", "lb Per Billling UOM", "Supplier S Name", "Vendor Invoice Price",
        "Actual Inv Cost(lb)", "Adj", "Market Cost", "Freight", "Landed Cost",
        "Recovery %", "Recovery Input", "Raw Material Input Qty", "Raw Material Per LB Cost",
        "Trim %", "Trim Cost/LB", "Recovery", "Input Cost", "Waste Output $", 
        "Net Input Cost", "Labour $", "Normal Sticker", "Material + Labour",
        "New Final Cost (Lb)", "Column1", "Billling UOM Cost", "Priced Sticker", "Final Cost",
        "Base Margin %", "Margin $", "Base Price", "List Price", "List Margin %"
    }
    for i in range(1, 5):
        cost_required.update({f"Item-{i}", f"Qty-{i}", f"Unit $-{i}", f"Total $-{i}"})
    export_required = {"Product Code", "Cost Price", "Base Price", "Suggested Price"}
    df_cost = auto_fill_missing_columns(df_cost, required_cols=cost_required)
    if not validate_columns(df_cost, cost_required, "Cost Sheet"):
        st.stop()
    if not validate_columns(df_export, export_required, "Export Sheet"):
        st.stop()
    st.success("✅ Files uploaded successfully!")
    st.markdown('<div class="header">Preview Sheets</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(df_cost.head(8), use_container_width=True)
    with c2:
        st.dataframe(df_export.head(8), use_container_width=True)

    # --- Cost Update Section ---
    st.markdown('<div class="header">Step 2: Bulk Update Product Costs</div>', unsafe_allow_html=True)
    st.info("🔄 You can edit prices below, **or** upload a simple two-column file (TRSM Code, New Cost Price).")
    sample_data = pd.DataFrame({"TRSM Code": ["", "", ""], "New Cost Price": [None, None, None]})
    edit_df = st.data_editor(
        sample_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="bulk_cost_editor",
        column_order=["TRSM Code", "New Cost Price"]
    )
    # --- File upload alternative ---
    st.markdown('<div class="note">Upload CSV/Excel for bulk pricing: <b>TRSM Code</b>, <b>New Cost Price</b></div>', unsafe_allow_html=True)
    uploaded_price_file = st.file_uploader("Bulk Price Update File", type=["csv", "xlsx"], key="pricefile")
    if uploaded_price_file:
        try:
            if uploaded_price_file.name.endswith('.csv'):
                price_df = pd.read_csv(uploaded_price_file)
            else:
                price_df = pd.read_excel(uploaded_price_file)
            st.success("Bulk price update file loaded!")
            st.dataframe(price_df)
            # Clean and convert input
            price_df["TRSM Code"] = price_df["TRSM Code"].astype(str).apply(clean_trsm_code)
            price_df["New Cost Price"] = pd.to_numeric(price_df["New Cost Price"], errors="coerce")
            price_df = price_df[price_df["TRSM Code"].str.strip() != ""]
            price_df = price_df[price_df["New Cost Price"].notna()]
            price_df = price_df[price_df["New Cost Price"] > 0]
            changes_to_apply = price_df
        except Exception as e:
            st.error(f"Could not read uploaded file: {e}")
            changes_to_apply = pd.DataFrame()
    else:
        # Use editable table
        changes_to_apply = edit_df.dropna(subset=["TRSM Code", "New Cost Price"])
        changes_to_apply["New Cost Price"] = pd.to_numeric(changes_to_apply["New Cost Price"], errors="coerce")
        changes_to_apply = changes_to_apply[changes_to_apply["TRSM Code"].str.strip() != ""]
        changes_to_apply = changes_to_apply[changes_to_apply["New Cost Price"].notna()]
        changes_to_apply = changes_to_apply[changes_to_apply["New Cost Price"] > 0]

    # --- Apply All Changes ---
    if changes_to_apply.empty:
        st.info("Add at least one valid TRSM Code and Cost Price above or upload a file to enable the Update button.")
    if st.button("✅ Apply All Cost Changes", type="primary", disabled=changes_to_apply.empty):
        summary_rows = []
        all_updated_codes = set()
        df_cost_updated = df_cost.copy()
        df_export_updated = df_export.copy()
        try:
            for i, row in changes_to_apply.iterrows():
                code, price = str(row["TRSM Code"]).strip(), float(row["New Cost Price"])
                df_cost_updated, cost_updated, updated_codes = update_cost_sheet(df_cost_updated, code, price)
                df_export_updated, export_updated = update_export_sheet(df_export_updated, df_cost_updated, updated_codes)
                all_updated_codes.update(updated_codes)
                summary_rows.append({
                    "TRSM Code": code,
                    "New Cost Price": price,
                    "Updated?": "Yes" if cost_updated or export_updated else "No"
                })
            st.markdown('<div class="header">Summary of Updates</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)
            if all_updated_codes:
                st.success(f"✅ Updated pricing for {len(all_updated_codes)} TRSM Code(s).")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown('<div class="header">Updated Cost Sheet Rows</div>', unsafe_allow_html=True)
                    st.dataframe(df_cost_updated[df_cost_updated["TRSM Code"].isin(all_updated_codes)])
                with col2:
                    st.markdown('<div class="header">Updated Export Sheet Rows</div>', unsafe_allow_html=True)
                    st.dataframe(df_export_updated[df_export_updated["Product Code"].isin(all_updated_codes)])
            else:
                st.warning("No rows were updated. Please check your input.")
            # --- Download Excel/ZIP ---
            cost_buf = io.BytesIO()
            export_buf = io.BytesIO()
            with pd.ExcelWriter(cost_buf, engine=EXCEL_ENGINE) as writer:
                df_cost_updated.to_excel(writer, sheet_name=cost_sheet_name, index=False)
                for sheet_name, df in other_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            cost_buf.seek(0)
            with pd.ExcelWriter(export_buf, engine=EXCEL_ENGINE) as writer:
                df_export_updated.to_excel(writer, sheet_name=export_sheet_name, index=False)
            export_buf.seek(0)
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr(f"Updated_Cost_Sheet_BULK.xlsx", cost_buf.getvalue())
                zip_file.writestr(f"Updated_Export_Sheet_BULK.xlsx", export_buf.getvalue())
            zip_buf.seek(0)
            st.markdown('<div class="header">Download Updated Files</div>', unsafe_allow_html=True)
            st.download_button(
                label="⬇️ Download Updated Files (ZIP)",
                data=zip_buf.getvalue(),
                file_name=f"Updated_Files_BULK.zip",
                mime="application/zip"
            )
        except Exception as e:
            st.error(f"Error during update: {e}")
    st.markdown(f'<div class="footer">Powered by Kush | {datetime.now().strftime("%B %d, %Y")} | Version {VERSION}</div>', unsafe_allow_html=True)
else:
    st.info("Please upload both Cost Sheet and Export Sheet to proceed.")
