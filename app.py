import streamlit as st
import pandas as pd
import io
import zipfile
from typing import Tuple, Optional, Set
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("TRSM_Cost_Adjustment")

# Page config
st.set_page_config(page_title="TRSM Cost Adjustment Tool", layout="wide")

# Try xlsxwriter first, else openpyxl
try:
    import xlsxwriter
    EXCEL_ENGINE = "xlsxwriter"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

# Constants
VERSION = "1.0.12"
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
DEFAULT_BASE_MARGIN = 0.17   # 17% in decimal
DEFAULT_LIST_MARGIN = 0.25   # 25% in decimal

# ------------------------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------------------------
def clean_trsm_code(code: str) -> str:
    """
    Clean TRSM code by removing commas and trailing .0
    """
    code_str = str(code).strip()
    try:
        val = float(code_str.replace(",", ""))
        if val.is_integer():
            code_str = str(int(val))
        else:
            code_str = str(val)
    except ValueError:
        code_str = code_str.replace(",", "")
    return code_str

def safe_float(value, default: float = 0.0) -> float:
    """Safely convert a value to float; return default if it fails."""
    try:
        return float(value) if not pd.isna(value) else default
    except (ValueError, TypeError):
        return default

def get_freight_cost(vendor: str) -> float:
    """Return the freight cost based on vendor name."""
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

def calculate_actual_inv_cost(vendor_invoice_price: float, lb_per_billling_uom: float) -> float:
    """Actual Inv Cost(lb) = vendor_invoice_price / lb_per_billling_uom."""
    return vendor_invoice_price if lb_per_billling_uom == 0 else vendor_invoice_price / lb_per_billling_uom

def calculate_market_cost(actual_inv_cost: float, adj: float) -> float:
    """Market Cost = actual_inv_cost + adj."""
    return actual_inv_cost + adj

def calculate_landed_cost(market_cost: float, freight: float) -> float:
    """Landed Cost = market_cost + freight."""
    return market_cost + freight

def calculate_recovery_input(market_cost: float, freight: float, recovery_percent: float) -> float:
    """
    Recovery Input = (market_cost + freight) / recovery_percent
    Avoid division by zero if recovery_percent = 0
    """
    return (market_cost + freight) / recovery_percent if recovery_percent != 0 else 0.0

def compute_totals_for_inputs(row: pd.Series) -> pd.Series:
    """
    For each of the 4 item columns, recalc Total $-i = Qty-i * Unit $-i.
    """
    for i in range(1, 5):
        qty = safe_float(row.get(f"Qty-{i}"), 0.0)
        unit_cost = safe_float(row.get(f"Unit $-{i}"), 0.0)
        row[f"Total $-{i}"] = qty * unit_cost
    return row

def calculate_raw_material_per_lb_cost(row: pd.Series) -> float:
    """
    Raw Material Per LB Cost = sum(Total $-1..4).
    """
    total = 0.0
    for i in range(1, 5):
        total += safe_float(row.get(f"Total $-{i}"), 0.0)
    return total

def calculate_waste_output(raw_material_cost: float, recovery_percent: float) -> float:
    """
    Waste Output = (raw_material_cost / recovery_percent) - raw_material_cost
    """
    if recovery_percent == 0:
        return 0.0
    return (raw_material_cost / recovery_percent) - raw_material_cost

def calculate_recovery(trim_cost_lb: float, trim_percent: float, recovery_percent: float) -> float:
    """
    Recovery = (trim_cost_lb * trim_percent) / recovery_percent
    """
    if recovery_percent == 0:
        return 0.0
    return (trim_cost_lb * trim_percent) / recovery_percent

def calculate_price_from_margin(cost: float, margin_percent: float) -> float:
    """Compute price from cost & margin. If margin≥1 => fallback to cost."""
    if margin_percent >= 1:
        logger.warning(f"Margin % ≥ 100% ({margin_percent}), using cost.")
        return cost
    return cost / (1 - margin_percent)

def calculate_margin_dollars(base_price: float, final_cost: float) -> float:
    """Margin $ = base_price - final_cost."""
    return base_price - final_cost

# ------------------------------------------------------------------------------
# Reading Excel
# ------------------------------------------------------------------------------
def read_files(cost_file, export_file, cost_sheet_name: str, export_sheet_name: str):
    """Read cost and export Excel files, do basic cleaning."""
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

    except ValueError as e:
        st.error(f"Sheet not found: {e}.")
        return None, None, None
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

# ------------------------------------------------------------------------------
# The main update function for a single row
# ------------------------------------------------------------------------------
def update_cost_row(
    row: pd.Series,
    new_cost_price: float = None,
    original_row: pd.Series = None
) -> pd.Series:
    """
    Update a single row with the logic:

      1) Recompute Totals for items
      2) raw_material_cost
      3) actual_inv_cost
      4) market_cost
      5) landed_cost
      6) recovery_input
      7) waste_output
      8) recovery
      9) input_cost
      10) net_input_cost
      11) new_final_cost_lb
      12) final_cost
      13) base_price from Base Margin %, list_price from List Margin %
    """
    # Keep old columns for reference
    old_vendor_invoice = safe_float(original_row.get("Vendor Invoice Price") if original_row is not None else row.get("Vendor Invoice Price"))
    old_final_cost = safe_float(original_row.get("Final Cost") if original_row is not None else row.get("Final Cost"))
    old_base_price = safe_float(original_row.get("Base Price") if original_row is not None else row.get("Base Price"))
    old_list_price = safe_float(original_row.get("List Price") if original_row is not None else row.get("List Price"))

    # 0) Recompute Totals
    row = compute_totals_for_inputs(row)

    # 1) Actual Inv Cost
    lb_per_billling_uom = safe_float(row.get("lb Per Billling UOM", 1), 1)
    vendor_invoice_price = new_cost_price if new_cost_price is not None else safe_float(row.get("Vendor Invoice Price"))
    actual_inv_cost = calculate_actual_inv_cost(vendor_invoice_price, lb_per_billling_uom)

    # 2) Market / Landed
    adj = safe_float(row.get("Adj", 0))
    market_cost = calculate_market_cost(actual_inv_cost, adj)
    freight = get_freight_cost(row.get("Supplier S Name", ""))
    landed_cost = calculate_landed_cost(market_cost, freight)

    # 3) Recovery %
    raw_recovery_val = safe_float(row.get("Recovery %", DEFAULT_RECOVERY), DEFAULT_RECOVERY)
    if raw_recovery_val > 1.0:
        raw_recovery_val = raw_recovery_val / 100.0
    if raw_recovery_val <= 0.0:
        raw_recovery_val = DEFAULT_RECOVERY
    recovery_percent = raw_recovery_val

    # 4) Recovery Input
    recovery_input = calculate_recovery_input(market_cost, freight, recovery_percent)

    # 5) Raw Material Per LB Cost
    raw_material_cost = calculate_raw_material_per_lb_cost(row)

    # 6) Waste Output
    waste_output = calculate_waste_output(raw_material_cost, recovery_percent)

    # 7) Recovery
    trim_percent = safe_float(row.get("Trim %", DEFAULT_TRIM))
    trim_cost_lb = safe_float(row.get("Trim Cost/LB", 0.0))
    recovery = calculate_recovery(trim_cost_lb, trim_percent, recovery_percent)

    # 8) Input Cost
    input_cost = recovery_input + raw_material_cost + waste_output

    # 9) Net Input Cost
    net_input_cost = input_cost - recovery

    # 10) New Final Cost (Lb)
    labour = safe_float(row.get("Labour $", DEFAULT_LABOUR))
    sticker = safe_float(row.get("Normal Sticker", DEFAULT_STICKER))
    new_final_cost_lb = net_input_cost + labour + sticker

    # 11) Billling UOM Cost
    column1_value = safe_float(row.get("Column1", 0))
    billling_uom_cost = new_final_cost_lb * lb_per_billling_uom + column1_value

    # 12) Final Cost
    priced_sticker = safe_float(row.get("Priced Sticker", 0))
    final_cost = billling_uom_cost + priced_sticker

    # 13) Base Price from "Base Margin %"
    base_margin_value = safe_float(row.get("Base Margin %", 0.0))
    if base_margin_value == 0.0:
        base_margin_value = DEFAULT_BASE_MARGIN
    if base_margin_value > 1.0:
        base_margin_decimal = base_margin_value / 100.0
    else:
        base_margin_decimal = base_margin_value
    base_price = calculate_price_from_margin(final_cost, base_margin_decimal)

    # Now read "List Margin %" from the cost sheet
    list_margin_value = safe_float(row.get("List Margin %", 0.0))
    if list_margin_value == 0.0:
        list_margin_value = DEFAULT_LIST_MARGIN
    elif list_margin_value > 1.0:
        list_margin_value = list_margin_value / 100.0

    list_price = calculate_price_from_margin(final_cost, list_margin_value)
    margin_dollars = calculate_margin_dollars(base_price, final_cost)

    # Logging
    logger.info(f"Row TRSM Code: {row.get('TRSM Code', 'N/A')}")
    logger.info(f"Final Cost: {final_cost:.2f}, Base Price: {base_price:.2f}, List Price: {list_price:.2f}, Margin: {margin_dollars:.2f}")

    # Store updated fields
    row["Price Change Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["Old Vendor Invoice Price"] = old_vendor_invoice
    row["Vendor Invoice Price"] = vendor_invoice_price
    row["Actual Inv Cost(lb)"] = actual_inv_cost
    row["Adj"] = adj
    row["Market Cost"] = market_cost
    row["Freight"] = freight
    row["Landed Cost"] = landed_cost
    row["Recovery %"] = recovery_percent * 100  # store as integer-based %
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

# ------------------------------------------------------------------------------
# Updating the cost sheet for a given TRSM code
# ------------------------------------------------------------------------------
def update_cost_sheet(
    df_cost: pd.DataFrame,
    trsm_code: str,
    new_cost_price: float
) -> Tuple[pd.DataFrame, bool, Set[str]]:
    """
    We do NOT ask for list_margin_percent from the UI. We read "List Margin %" from each row in df_cost.
    """
    df_updated = df_cost.copy()
    trsm_code_clean = clean_trsm_code(trsm_code)
    updated_flag = False
    updated_trsm_codes = set()

    df_updated["TRSM Code"] = df_updated["TRSM Code"].apply(clean_trsm_code)

    # Initialize columns for history if missing
    for col in ["Old Vendor Invoice Price", "Old Final Cost", "Old Base Price", "Old List Price", "Price Change Date"]:
        if col not in df_updated.columns:
            df_updated[col] = None

    # 1) Update the main TRSM code
    mask_main = df_updated["TRSM Code"] == trsm_code_clean
    if mask_main.any():
        st.write(f"Updating TRSM Code: {trsm_code_clean} with new cost: {new_cost_price}")
        for idx in df_updated[mask_main].index:
            original_row = df_cost.loc[idx]
            df_updated.loc[idx] = update_cost_row(
                row=df_updated.loc[idx],
                new_cost_price=new_cost_price,
                original_row=original_row
            )
        updated_trsm_codes.update(df_updated.loc[mask_main, "TRSM Code"])
        updated_flag = True

    # 2) Iteratively update items that use the updated TRSM codes
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
                        df_updated.loc[idx, unit_col] = safe_float(matching_row.iloc[0]["Vendor Invoice Price"])
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

    if not updated_flag:
        st.warning(f"No matches found for TRSM Code: {trsm_code_clean}")

    return df_updated, updated_flag, updated_trsm_codes

# ------------------------------------------------------------------------------
# Update the export sheet
# ------------------------------------------------------------------------------
def update_export_sheet(
    df_export: pd.DataFrame,
    df_cost_updated: pd.DataFrame,
    updated_trsm_codes: Set[str]
) -> Tuple[pd.DataFrame, bool]:
    df_export_updated = df_export.copy()
    updated_flag = False

    df_export_updated["Product Code"] = df_export_updated["Product Code"].apply(clean_trsm_code)
    
    for trsm_code in updated_trsm_codes:
        cost_row = df_cost_updated[df_cost_updated["TRSM Code"] == trsm_code]
        if cost_row.empty:
            st.warning(f"No matching TRSM Code {trsm_code} in updated Cost Sheet.")
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
        else:
            st.warning(f"No matches in Export Sheet for TRSM Code: {trsm_code}")
            
    return df_export_updated, updated_flag

# ------------------------------------------------------------------------------
# UI styling
# ------------------------------------------------------------------------------
st.write("""
    <style>
    .title { text-align: center; color: #333; font-size: 2.5em; }
    .header { color: #555; font-size: 1.5em; margin-top: 1em; }
    .footer { text-align: center; color: #777; font-size: 0.9em; margin-top: 2em; }
    .stButton>button { background-color: #007bff; color: white; border-radius: 5px; padding: 0.6em 1.2em; }
    .stButton>button:hover { background-color: #0056b3; }
    </style>
""", unsafe_allow_html=True)

st.write(f'<div class="title">TRSM Product Cost Adjustment Tool (v{VERSION})</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# Main Application
# ------------------------------------------------------------------------------
st.write('<div class="header">1. Upload Excel Files</div>', unsafe_allow_html=True)
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

    # Validate columns
    cost_required = {
        "TRSM Code", "lb Per Billling UOM", "Supplier S Name", "Vendor Invoice Price",
        "Actual Inv Cost(lb)", "Adj", "Market Cost", "Freight", "Landed Cost",
        "Recovery %", "Recovery Input", "Raw Material Input Qty", "Raw Material Per LB Cost",
        "Trim %", "Trim Cost/LB", "Recovery", "Input Cost", "Waste Output $", 
        "Net Input Cost", "Labour $", "Normal Sticker", "Material + Labour",
        "New Final Cost (Lb)", "Column1", "Billling UOM Cost", "Priced Sticker", "Final Cost",
        "Base Margin %", "Margin $", "Base Price", "List Price",
        # NEW: We must have "List Margin %" in the cost sheet
        "List Margin %"
    }
    for i in range(1, 5):
        cost_required.update({f"Item-{i}", f"Qty-{i}", f"Unit $-{i}", f"Total $-{i}"})
    export_required = {"Product Code", "Cost Price", "Base Price", "Suggested Price"}

    if not validate_columns(df_cost, cost_required, "Cost Sheet"):
        st.stop()
    if not validate_columns(df_export, export_required, "Export Sheet"):
        st.stop()

    st.success("Files uploaded successfully!")
    col1, col2 = st.columns(2)
    with col1:
        st.write('<div class="header">Cost Sheet Preview</div>', unsafe_allow_html=True)
        st.dataframe(df_cost.head(10))
    with col2:
        st.write('<div class="header">Export Sheet Preview</div>', unsafe_allow_html=True)
        st.dataframe(df_export.head(10))

    # 2. Update
    st.write('<div class="header">2. Update Product Cost</div>', unsafe_allow_html=True)
    trsm_code = st.text_input("TRSM Code to Update", "").strip()
    new_cost_price = st.number_input("New Cost Price", min_value=0.0, step=0.01, format="%.2f")
    st.write("Note: Both Base Margin % and List Margin % come from the Cost Sheet columns. "
             "No separate input for List Margin is required.")
    
    if st.button("Apply Cost Changes"):
        if not trsm_code:
            st.error("Please enter a valid TRSM Code.")
            st.stop()

        try:
            df_cost_updated, cost_updated, updated_trsm_codes = update_cost_sheet(
                df_cost,
                trsm_code,
                new_cost_price
            )
            df_export_updated, export_updated = update_export_sheet(df_export, df_cost_updated, updated_trsm_codes)

            if cost_updated or export_updated:
                st.success(f"✅ Updated pricing for TRSM Code(s): {', '.join(updated_trsm_codes)}")

                # Show updated previews
                col1, col2 = st.columns(2)
                with col1:
                    st.write('<div class="header">Updated Cost Sheet Preview</div>', unsafe_allow_html=True)
                    mask_main = df_cost_updated["TRSM Code"].isin(updated_trsm_codes)
                    item_cols = [c for c in df_cost_updated.columns if c.startswith("Item-")]
                    mask_item = pd.concat([df_cost_updated[col].isin(updated_trsm_codes) for col in item_cols], axis=1).any(axis=1)
                    st.dataframe(df_cost_updated[mask_main | mask_item])

                with col2:
                    st.write('<div class="header">Updated Export Sheet Preview</div>', unsafe_allow_html=True)
                    export_mask = df_export_updated["Product Code"].isin(updated_trsm_codes)
                    st.dataframe(df_export_updated[export_mask])

                # Create Excel files in-memory
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

                # Zip both
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    zip_file.writestr(f"Updated_Cost_Sheet_{trsm_code}.xlsx", cost_buf.getvalue())
                    zip_file.writestr(f"Updated_Export_Sheet_{trsm_code}.xlsx", export_buf.getvalue())
                zip_buf.seek(0)

                st.write('<div class="header">Download Updated Files</div>', unsafe_allow_html=True)
                st.download_button(
                    label="Download Updated Files (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name=f"Updated_Files_{trsm_code}.zip",
                    mime="application/zip"
                )
            else:
                st.warning(f"No updates applied for TRSM Code: {trsm_code}.")
        except Exception as e:
            st.error(f"Error during update: {e}")
            st.stop()
else:
    st.info("Please upload both Cost Sheet and Export Sheet to proceed.")

st.write(f'<div class="footer">Powered by Kush | March 19, 2025 | Version {VERSION}</div>', unsafe_allow_html=True)
