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

# Page configuration
st.set_page_config(page_title="TRSM Cost Adjustment Tool", layout="wide")

# Excel engine selection
try:
    import xlsxwriter
    EXCEL_ENGINE = "xlsxwriter"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

# Constants
VERSION = "1.0.12"  # Updated version
DEFAULT_RECOVERY = 1.0
DEFAULT_TRIM = 0.0
DEFAULT_LABOUR = 0.0
DEFAULT_STICKER = 0.0
FREIGHT_RATES = {"PRAN": 0.07, "Local Pickup": 0.11, "Alberta": 0.205, "Ontario/Quebec": 0.25}
DEFAULT_FREIGHT = 0.0
DEFAULT_BASE_MARGIN = 0.17  # 17% fallback if not in file (stored as decimal)
DEFAULT_LIST_MARGIN = 0.25  # 25% default for List Margin

# Helper Functions
def clean_trsm_code(code: str) -> str:
    """
    Clean TRSM code by:
      1) Removing trailing '.0'
      2) Removing commas
    """
    code_str = str(code).rstrip('.0')
    code_str = code_str.replace(",", "")  # remove commas entirely
    return code_str

def safe_float(value, default: float = 0.0) -> float:
    """Convert value to float safely, returning default if conversion fails."""
    try:
        return float(value) if not pd.isna(value) else default
    except (ValueError, TypeError):
        return default

def get_freight_cost(vendor: str) -> float:
    """Determine freight cost based on vendor name."""
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
    """Calculate actual invoice cost per billing UOM."""
    return vendor_invoice_price if lb_per_billling_uom == 0 else vendor_invoice_price / lb_per_billling_uom

def calculate_market_cost(actual_inv_cost: float, adj: float) -> float:
    """Calculate market cost by adding adjustment."""
    return actual_inv_cost + adj

def calculate_landed_cost(market_cost: float, freight: float) -> float:
    """Calculate landed cost by adding freight."""
    return market_cost + freight

def calculate_recovery_input(landed_cost: float, recovery_percent: float) -> float:
    """Calculate recovery input based on percentage."""
    return landed_cost * recovery_percent

def calculate_input_costs(row: pd.Series) -> float:
    """Calculate total input costs from quantities and unit costs."""
    total = 0.0
    for i in range(1, 5):
        qty = safe_float(row.get(f"Qty-{i}"))
        unit_cost = safe_float(row.get(f"Unit $-{i}"))
        total += qty * unit_cost
        row[f"Total $-{i}"] = qty * unit_cost
    return total

def calculate_raw_material_per_lb_cost(input_cost: float, raw_material_input_qty: float) -> float:
    """Calculate raw material cost per pound."""
    return input_cost if raw_material_input_qty == 0 else input_cost / raw_material_input_qty

def calculate_recovery(raw_material_per_lb: float, trim_percent: float) -> float:
    """Calculate recovery considering trim percentage."""
    return raw_material_per_lb if trim_percent >= 1 else raw_material_per_lb / (1 - trim_percent)

def calculate_material_labour(net_input_cost: float, labour: float, sticker: float) -> float:
    """Calculate total material and labour cost."""
    return net_input_cost + labour + sticker

def calculate_billling_uom_cost(new_final_cost_lb: float, lb_per_billling_uom: float) -> float:
    """Calculate billing UOM cost."""
    return new_final_cost_lb * lb_per_billling_uom

def calculate_final_cost(billling_uom_cost: float, priced_sticker: float) -> float:
    """Calculate final cost including priced sticker."""
    return billling_uom_cost + priced_sticker

def calculate_price_from_margin(cost: float, margin_percent: float) -> float:
    """
    Calculate price based on cost and margin percentage (decimal).
    E.g. if margin_percent=0.17 => price = cost / (1 - 0.17).
    """
    # If margin is 100% or more, fallback to cost to avoid division by zero or negative
    if margin_percent >= 1:
        logger.warning(f"Margin % ≥ 100% ({margin_percent}), using Cost.")
        return cost
    return cost / (1 - margin_percent)

def calculate_margin_dollars(base_price: float, final_cost: float) -> float:
    """Calculate Margin $ as Base Price - Final Cost."""
    return base_price - final_cost

def read_files(cost_file, export_file, cost_sheet_name: str, export_sheet_name: str) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], Optional[dict]]:
    """Read cost and export Excel files."""
    try:
        excel_cost = pd.ExcelFile(cost_file)
        df_cost = pd.read_excel(cost_file, sheet_name=cost_sheet_name, engine="openpyxl")
        other_sheets = {sheet: excel_cost.parse(sheet) for sheet in excel_cost.sheet_names if sheet != cost_sheet_name}
        df_export = pd.read_excel(export_file, sheet_name=export_sheet_name, engine="openpyxl")

        # Strip extra spaces/newlines from columns
        df_cost.columns = [col.strip().replace('\n', ' ') for col in df_cost.columns]
        df_export.columns = [col.strip().replace('\n', ' ') for col in df_export.columns]

        # IMPORTANT: Clean TRSM codes and Product codes immediately
        if "TRSM Code" in df_cost.columns:
            df_cost["TRSM Code"] = df_cost["TRSM Code"].astype(str).apply(clean_trsm_code)
        if "Product Code" in df_export.columns:
            df_export["Product Code"] = df_export["Product Code"].astype(str).apply(clean_trsm_code)

        if df_cost.empty or df_export.empty:
            st.error("One or both sheets are empty.")
            return None, None, None
        return df_cost, df_export, other_sheets
    except ValueError as e:
        st.error(f"Sheet not found: {e}.")
        return None, None, None
    except Exception as e:
        st.error(f"Error reading files: {e}.")
        return None, None, None

def validate_columns(df: pd.DataFrame, required_cols: set, sheet_name: str) -> bool:
    """Validate that required columns exist in the DataFrame."""
    available_cols = set(df.columns)
    missing = required_cols - available_cols
    if missing:
        st.error(f"{sheet_name} missing columns: {missing}")
        return False
    return True

def update_cost_row(
    row: pd.Series,
    new_cost_price: float = None,
    original_row: pd.Series = None,
    list_margin_percent: float = DEFAULT_LIST_MARGIN
) -> pd.Series:
    """
    Update a single row in the cost sheet with new calculations.
    The final List Price is based on Final Cost (not Base Price).
    """
    old_vendor_invoice = safe_float(original_row.get("Vendor Invoice Price") if original_row is not None else row.get("Vendor Invoice Price"))
    old_final_cost = safe_float(original_row.get("Final Cost") if original_row is not None else row.get("Final Cost"))
    old_base_price = safe_float(original_row.get("Base Price") if original_row is not None else row.get("Base Price"))
    old_list_price = safe_float(original_row.get("List Price") if original_row is not None else row.get("List Price"))

    # Step 1: Update vendor invoice price and calculate market cost
    lb_per_billling_uom = safe_float(row.get("lb Per Billling UOM", 1), 1)
    vendor_invoice_price = new_cost_price if new_cost_price is not None else safe_float(row.get("Vendor Invoice Price"))
    actual_inv_cost = calculate_actual_inv_cost(vendor_invoice_price, lb_per_billling_uom)
    adj = safe_float(row.get("Adj", 0))
    market_cost = calculate_market_cost(actual_inv_cost, adj)
    
    # Calculate freight and landed cost
    freight = get_freight_cost(row.get("Supplier S Name", ""))
    landed_cost = calculate_landed_cost(market_cost, freight)
    
    # Step 2: Calculate recovery input and raw material costs
    recovery_percent = safe_float(row.get("Recovery %", DEFAULT_RECOVERY), DEFAULT_RECOVERY)
    recovery_input = calculate_recovery_input(landed_cost, recovery_percent)
    
    # If we are updating the cost from a single new_cost_price, we skip the item-by-item cost
    # and treat the row's "Vendor Invoice Price" as the cost. Otherwise, we do the normal logic.
    if new_cost_price is not None:
        input_cost = vendor_invoice_price
    else:
        input_cost = calculate_input_costs(row)
    
    raw_material_input_qty = safe_float(row.get("Raw Material Input Qty", 1), 1)
    raw_material_per_lb = calculate_raw_material_per_lb_cost(input_cost, raw_material_input_qty)
    
    # Calculate trim and recovery
    trim_percent = safe_float(row.get("Trim %", 0.0), 0.0)
    recovery = calculate_recovery(raw_material_per_lb, trim_percent)
    net_input_cost = recovery
    
    # Step 3: Add labour and calculate new final cost
    labour = safe_float(row.get("Labour $", DEFAULT_LABOUR), DEFAULT_LABOUR)
    sticker = safe_float(row.get("Normal Sticker", DEFAULT_STICKER), DEFAULT_STICKER)
    material_labour = calculate_material_labour(net_input_cost, labour, sticker)
    new_final_cost_lb = material_labour
    
    # Step 4: Calculate billing UOM cost, final cost, and prices
    billling_uom_cost = calculate_billling_uom_cost(new_final_cost_lb, lb_per_billling_uom)
    priced_sticker = safe_float(row.get("Priced Sticker", 0))
    final_cost = calculate_final_cost(billling_uom_cost, priced_sticker)

    # ---- FIX FOR BASE MARGIN % ----
    # If the sheet has "17" we treat it as 17%. If it has "0.17" we treat that as 17% in decimal form.
    base_margin_value = safe_float(row.get("Base Margin %", 0.0))
    if base_margin_value == 0.0:
        base_margin_value = DEFAULT_BASE_MARGIN  # 0.17

    # If above 1, interpret as an integer-based percentage (e.g. 17 => 0.17).
    if base_margin_value > 1:
        base_margin_decimal = base_margin_value / 100.0
    else:
        base_margin_decimal = base_margin_value
    
    # Calculate base price (using decimal, e.g. 0.17)
    base_price = calculate_price_from_margin(final_cost, base_margin_decimal)

    # Calculate list price from Final Cost (not from Base Price)
    list_price = calculate_price_from_margin(final_cost, list_margin_percent)

    # Calculate Margin $ (based on Base Price - Final Cost)
    margin_dollars = calculate_margin_dollars(base_price, final_cost)

    # Log calculations for verification
    logger.info(f"Row TRSM Code: {row.get('TRSM Code', 'N/A')}, Base Margin (decimal): {base_margin_decimal}")
    logger.info(f"Final Cost: {final_cost:.2f}, Base Price: {base_price:.2f}, List Price: {list_price:.2f}, Margin $: {margin_dollars:.2f}")

    # Update row with new values
    row["Price Change Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["Old Vendor Invoice Price"] = old_vendor_invoice
    row["Vendor Invoice Price"] = vendor_invoice_price
    row["Actual Inv Cost(lb)"] = actual_inv_cost
    row["Market Cost"] = market_cost
    row["Freight"] = freight
    row["Landed Cost"] = landed_cost
    row["Recovery Input"] = recovery_input
    row["Input Cost"] = input_cost
    row["Raw Material Per LB Cost"] = raw_material_per_lb
    row["Recovery"] = recovery
    row["Net Input Cost"] = net_input_cost
    row["Material + Labour"] = material_labour
    row["New Final Cost (Lb)"] = new_final_cost_lb
    row["Old Final Cost"] = old_final_cost
    row["Billling UOM Cost"] = billling_uom_cost
    row["Final Cost"] = final_cost
    row["Old Base Price"] = old_base_price
    row["Base Price"] = base_price
    row["Old List Price"] = old_list_price
    row["List Price"] = list_price
    row["Waste Output $"] = 0.0
    row["Trim Cost/LB"] = 0.0

    # We store 'Base Margin %' back as an integer-based percentage for clarity, e.g. 17.0 for 17%
    row["Base Margin %"] = base_margin_decimal * 100
    row["Margin $"] = margin_dollars 
    return row

def update_cost_sheet(
    df_cost: pd.DataFrame,
    trsm_code: str,
    new_cost_price: float,
    list_margin_percent: float
) -> Tuple[pd.DataFrame, bool, Set[str]]:
    """Update the cost sheet for a given TRSM code."""
    df_updated = df_cost.copy()
    trsm_code_clean = clean_trsm_code(trsm_code)
    updated_flag = False
    updated_trsm_codes = set()

    # Ensure TRSM codes are properly formatted
    df_updated["TRSM Code"] = df_updated["TRSM Code"].astype(str).apply(clean_trsm_code)
    
    # Initialize history columns if they don't exist
    for col in ["Old Vendor Invoice Price", "Old Final Cost", "Old Base Price", "Old List Price", "Price Change Date"]:
        if col not in df_updated.columns:
            df_updated[col] = None

    # First, update the main TRSM code
    mask_main = df_updated["TRSM Code"] == trsm_code_clean
    if mask_main.any():
        st.write(f"Updating TRSM Code: {trsm_code_clean} with new cost: {new_cost_price}")
        for idx in df_updated[mask_main].index:
            original_row = df_cost.loc[idx]
            df_updated.loc[idx] = update_cost_row(
                df_updated.loc[idx],
                new_cost_price=new_cost_price,
                original_row=original_row,
                list_margin_percent=list_margin_percent
            )
        updated_trsm_codes.update(df_updated.loc[mask_main, "TRSM Code"])
        updated_flag = True

    # Then iteratively update any items that use the updated TRSM codes as inputs
    item_cols = [c for c in df_updated.columns if c.startswith("Item-")]
    to_update = True
    iteration = 0
    while to_update and iteration < 10:  # Limit iterations to prevent infinite loops
        to_update = False
        composite_mask = pd.Series(False, index=df_updated.index)
        
        for item_col in item_cols:
            unit_col = f"Unit $-{item_col.split('-')[1]}"
            if unit_col not in df_updated.columns:
                continue
                
            # Clean TRSM codes in item columns
            df_updated[item_col] = df_updated[item_col].astype(str).apply(clean_trsm_code)
            
            # Find items that use updated TRSM codes as inputs
            mask_item = df_updated[item_col].isin(updated_trsm_codes)
            
            if mask_item.any():
                for idx in df_updated[mask_item].index:
                    # Get the TRSM code used as input
                    trsm_code_item = df_updated.loc[idx, item_col]
                    
                    # Find the matching row with updated cost
                    matching_row = df_updated[df_updated["TRSM Code"] == trsm_code_item]
                    
                    if not matching_row.empty:
                        # Update the unit cost with the new vendor invoice price
                        df_updated.loc[idx, unit_col] = safe_float(matching_row.iloc[0]["Vendor Invoice Price"])
                
                composite_mask |= mask_item
                to_update = True

        # Update all rows that use the updated TRSM codes
        if composite_mask.any():
            updated_trsm_codes.update(df_updated.loc[composite_mask, "TRSM Code"])
            for idx in df_updated[composite_mask].index:
                original_row = df_cost.loc[idx]
                df_updated.loc[idx] = update_cost_row(
                    df_updated.loc[idx],
                    new_cost_price=None,
                    original_row=original_row,
                    list_margin_percent=list_margin_percent
                )
            updated_flag = True
            
        iteration += 1

    if not updated_flag:
        st.warning(f"No matches found for TRSM Code: {trsm_code_clean}")
        
    return df_updated, updated_flag, updated_trsm_codes

def update_export_sheet(
    df_export: pd.DataFrame,
    df_cost_updated: pd.DataFrame,
    updated_trsm_codes: Set[str]
) -> Tuple[pd.DataFrame, bool]:
    """Update the export sheet based on updated cost sheet."""
    df_export_updated = df_export.copy()
    updated_flag = False

    # Ensure product codes are properly formatted
    df_export_updated["Product Code"] = df_export_updated["Product Code"].astype(str).apply(clean_trsm_code)
    
    for trsm_code in updated_trsm_codes:
        # Find the updated cost information
        cost_row = df_cost_updated[df_cost_updated["TRSM Code"] == trsm_code]
        
        if cost_row.empty:
            st.warning(f"No matching TRSM Code {trsm_code} in updated Cost Sheet.")
            continue
            
        # Get the updated prices
        new_final_cost_lb = safe_float(cost_row.iloc[0]["New Final Cost (Lb)"])
        final_cost = safe_float(cost_row.iloc[0]["Final Cost"])
        final_base = safe_float(cost_row.iloc[0]["Base Price"])
        final_list = safe_float(cost_row.iloc[0]["List Price"])
        
        # Update the export sheet
        mask_export = df_export_updated["Product Code"] == trsm_code
        if mask_export.any():
            df_export_updated.loc[mask_export, "Cost Price"] = final_cost
            df_export_updated.loc[mask_export, "Base Price"] = final_base
            df_export_updated.loc[mask_export, "Suggested Price"] = final_list
            updated_flag = True
        else:
            st.warning(f"No matches in Export Sheet for TRSM Code: {trsm_code}")
            
    return df_export_updated, updated_flag

# UI Styling
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

# Main Application
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
        "Trim %", "Recovery", "Labour $", "Normal Sticker", "Material + Labour",
        "New Final Cost (Lb)", "Billling UOM Cost", "Priced Sticker", "Final Cost",
        "Base Price", "List Price", "Base Margin %"
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

    st.write('<div class="header">2. Update Product Cost</div>', unsafe_allow_html=True)
    trsm_code = st.text_input("TRSM Code to Update", "").strip()
    new_cost_price = st.number_input("New Cost Price", min_value=0.0, step=0.01, format="%.2f")
    st.write("Note: Base Price uses 'Base Margin %' from the input cost file. List Price now uses the 'List Margin%' below *based on Final Cost*.")
    list_margin_percent_input = st.number_input("List Margin % (e.g., 25 for 25%)", min_value=0.0, max_value=99.99, value=25.0, step=0.1)
    list_margin_percent = list_margin_percent_input / 100.0  # convert e.g. 25 → 0.25

    if st.button("Apply Cost Changes"):
        if not trsm_code:
            st.error("Please enter a valid TRSM Code.")
            st.stop()

        try:
            df_cost_updated, cost_updated, updated_trsm_codes = update_cost_sheet(
                df_cost,
                trsm_code,
                new_cost_price,
                list_margin_percent
            )
            df_export_updated, export_updated = update_export_sheet(df_export, df_cost_updated, updated_trsm_codes)

            if cost_updated or export_updated:
                st.success(f"✅ Updated pricing for TRSM Code(s): {', '.join(updated_trsm_codes)}")
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

                # Create a zip file containing both Excel files
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
