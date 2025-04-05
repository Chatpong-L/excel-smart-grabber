import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime
import logging # Added for better error logging

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

st.set_page_config(page_title="Excel Smart Grabber", layout="wide")

st.title("📊 Excel Smart Grabber 3000 (Audit-Ready Version)")
st.markdown("Upload Excel files and extract specified columns from all sheets. Deleted rows and mismatched columns will be flagged and auditable.")

# ==== Sidebar Inputs ====
st.sidebar.header("Configuration")
# --- Updated Labels Below ---
value_label = st.sidebar.text_input("Main: Column Name to Grab", "รวมเงิน") # UPDATED LABEL
trans_label = st.sidebar.text_input("Secondary: Column Name to Grab", "รายการ") # UPDATED LABEL
typical_letter = st.sidebar.text_input("Main: Expected Column Letter (e.g. A-Z)", "M").upper() # UPDATED LABEL & Ensure uppercase
# --- End of Updated Labels ---
extra_cols_raw = st.sidebar.text_area("Extra Columns (one per line)", "")
remove_phrases_raw = st.sidebar.text_area("Remove Row Phrases (one per line)", "TOTAL")
max_scan = st.sidebar.number_input("Header Scan Limit (Rows)", 1, 30, 10)
audit_mode = st.sidebar.checkbox("Generate Audit File ZIP?", value=True)
output_filename = st.sidebar.text_input("Output Excel Name", "Master_Report.xlsx")

# Process multiline inputs
extra_cols = [col.strip() for col in extra_cols_raw.strip().splitlines() if col.strip()]
remove_phrases = [phrase.strip().lower() for phrase in remove_phrases_raw.strip().splitlines() if phrase.strip()] # Lowercase for comparison

uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

# ===== Helper Functions =====
def col_letter_to_index(letter):
    """Converts an Excel column letter (A, B, ..., Z, AA, etc.) to a zero-based index."""
    result = 0
    for char in letter: # Already ensured uppercase in sidebar input
        if 'A' <= char <= 'Z':
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            # Handle invalid input gracefully
            st.warning(f"Invalid character '{char}' in Expected Column Letter '{letter}'. Cannot convert.")
            return -1 # Return an invalid index
    return result - 1

def find_column(df, label_to_find, max_rows_to_scan):
    """
    Tries to find a column matching the label_to_find.
    1. Checks existing column names (case-insensitive).
    2. If not found, scans the first 'max_rows_to_scan' rows for the label.
    3. If found in a row, sets that row as the header, drops preceding rows,
       and recursively calls itself to find the column name in the new header.
    Returns the actual column name found in the DataFrame or None.
    """
    # Ensure label is lowercase for comparison
    label_lower = label_to_find.lower()

    # 1. Check current column names
    for col in df.columns:
        if label_lower in str(col).lower():
            logging.info(f"Found column '{col}' matching label '{label_to_find}' in header.")
            return col

    # 2. Scan rows if not found in header
    if max_rows_to_scan > 0:
        logging.info(f"Label '{label_to_find}' not in initial header. Scanning up to {max_rows_to_scan} rows...")
        for i in range(min(max_rows_to_scan, len(df))):
            try:
                row_values = df.iloc[i].astype(str).str.lower()
                if any(label_lower in cell for cell in row_values):
                    logging.info(f"Found label '{label_to_find}' in row {i}. Setting as new header.")
                    # Set the found row as header
                    df.columns = df.iloc[i]
                    # Drop rows above and including the new header row
                    df.drop(index=df.index[:i+1], inplace=True)
                    # Reset index after dropping rows
                    df.reset_index(drop=True, inplace=True)
                    # Clean up column names (e.g., strip whitespace)
                    df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed: {idx}" for idx, c in enumerate(df.columns)]
                    # Recursively call with max_rows_to_scan = 0 to search the *new* header
                    return find_column(df, label_to_find, 0)
            except Exception as e:
                logging.warning(f"Error processing row {i} while searching for header: {e}")
                continue # Try next row

    logging.warning(f"Label '{label_to_find}' not found in header or scanned rows.")
    return None

def safe_cell_write(worksheet, row, col, value, cell_format=None):
    """Writes cell value to worksheet, handling potential type errors."""
    try:
        # Handle specific pandas/numpy types that cause issues
        if pd.isna(value):
            worksheet.write_blank(row, col, None, cell_format)
        # Handle datetimes (Excel doesn't like timezone-aware)
        elif isinstance(value, (datetime, pd.Timestamp)):
             # If timezone-aware, convert to naive UTC or local time
            if getattr(value, 'tzinfo', None) is not None:
                 value = value.tz_convert(None) # Convert to naive UTC
            # xlsxwriter can handle naive datetime objects directly
            worksheet.write_datetime(row, col, value, cell_format)
        elif isinstance(value, (int, float)):
             worksheet.write_number(row, col, value, cell_format)
        elif isinstance(value, (bool)):
             worksheet.write_boolean(row, col, value, cell_format)
        else:
            # Default to string for anything else
            worksheet.write_string(row, col, str(value), cell_format)
    except TypeError as te:
        # Fallback: Try writing as string if specific types failed
        logging.warning(f"TypeError writing value '{value}' (type: {type(value)}) at row {row}, col {col}. Writing as string. Error: {te}")
        try:
            worksheet.write_string(row, col, str(value), cell_format)
        except Exception as e:
            logging.error(f"FAILED to write value '{value}' as string at row {row}, col {col}. Skipping cell. Error: {e}")
            worksheet.write_blank(row, col, None, cell_format) # Write blank on severe error
    except Exception as e:
         logging.error(f"Unexpected error writing value '{value}' (type: {type(value)}) at row {row}, col {col}. Skipping cell. Error: {e}")
         worksheet.write_blank(row, col, None, cell_format) # Write blank on severe error

# ===== Main Logic =====
if st.button("▶️ Run Excel Grabber") and uploaded_files:
    st.info("Processing started... Please wait.")
    progress_bar = st.progress(0)
    start_time = datetime.now()

    master_data = []
    all_deleted_rows = []
    not_typical_col_loc = []
    skipped_sheets_info = []
    error_logs = []

    # Prepare in-memory zip file for audit reports
    audit_zip_buffer = io.BytesIO()
    # Use 'a' mode if adding files iteratively, 'w' if creating fresh each time
    with zipfile.ZipFile(audit_zip_buffer, "w", zipfile.ZIP_DEFLATED) as audit_bundle:

        total_files = len(uploaded_files)
        for file_index, uploaded_file in enumerate(uploaded_files):
            st.text(f"Processing file: {uploaded_file.name}...")
            logging.info(f"Processing file: {uploaded_file.name}")

            # Use an in-memory buffer for the individual audit Excel file
            audit_excel_buffer = io.BytesIO()
            try:
                # Use try-except for file opening and sheet parsing
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names

                # Create ExcelWriter for the audit file in memory
                with pd.ExcelWriter(audit_excel_buffer, engine='xlsxwriter') as audit_writer:
                    workbook = audit_writer.book
                    # Define formats once
                    deleted_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"}) # Light red fill, dark red text
                    extracted_fmt = workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"}) # Yellow fill, dark yellow text

                    for sheet_index, sheet_name in enumerate(sheet_names):
                        logging.info(f"--- Processing sheet: {sheet_name}")
                        try:
                            # Read sheet
                            df = xls.parse(sheet_name, header=None) # Read without assuming header first
                            df.reset_index(drop=True, inplace=True)
                            raw_df = df.copy() # Keep a copy of the raw data *before* header manipulation

                            # --- Find Header and Key Columns ---
                            # Important: Pass a *copy* to find_column if you don't want the original df modified yet
                            # Or accept that df will be modified by find_column
                            df_processed = df.copy() # Work on a copy for finding columns/processing

                            # Use the variable 'value_label' which holds the *user input* (e.g., "รวมเงิน")
                            # The *display label* (e.g., "Main: Column Name to Grab") is only for the UI
                            value_col_name = find_column(df_processed, value_label, max_scan)

                            if not value_col_name:
                                st.warning(f"Skipping sheet '{sheet_name}' in file '{uploaded_file.name}': Required Main Column '{value_label}' not found within {max_scan} rows.")
                                skipped_sheets_info.append({"File": uploaded_file.name, "Sheet": sheet_name, "Reason": f"Main column '{value_label}' not found"})
                                logging.warning(f"Skipped sheet '{sheet_name}' in '{uploaded_file.name}': Main column '{value_label}' not found.")
                                continue # Skip this sheet

                            # Find other columns using their respective input variables
                            trans_col_name = find_column(df_processed, trans_label, 0) # Search only header now
                            extra_col_map = { # Maps user label to actual column name found
                                label: find_column(df_processed, label, 0)
                                for label in extra_cols
                            }
                            found_extra_cols = {k:v for k,v in extra_col_map.items() if v is not None}
                            missing_extra_cols = {k for k,v in extra_col_map.items() if v is None}
                            if missing_extra_cols:
                                st.warning(f"In '{uploaded_file.name}/{sheet_name}', could not find extra columns: {', '.join(missing_extra_cols)}")
                                logging.warning(f"In '{uploaded_file.name}/{sheet_name}', could not find extra columns: {', '.join(missing_extra_cols)}")


                            # --- Identify and Separate Deleted Rows ---
                            # Note: This checks the phrase anywhere in the string representation of the row.
                            # Consider making this check more specific if needed (e.g., only in transaction column).
                            if remove_phrases:
                                row_mask_to_delete = df_processed.apply(
                                    lambda row: any(phrase in str(cell).lower()
                                                    for cell in row.astype(str) # Check each cell individually
                                                    for phrase in remove_phrases),
                                    axis=1
                                )
                                deleted_df = df_processed[row_mask_to_delete].copy()
                                df_processed = df_processed[~row_mask_to_delete]
                            else:
                                deleted_df = pd.DataFrame() # Empty dataframe if no phrases

                            # --- Extract Data for Master Report ---
                            for index, row in df_processed.iterrows():
                                # Use the *user-provided labels* as keys in the output dict
                                entry = {
                                    "SourceFile": uploaded_file.name,
                                    "SourceSheet": sheet_name,
                                    value_label: row.get(value_col_name), # Use the variable holding the user's input label
                                    trans_label: row.get(trans_col_name) if trans_col_name else None # Use the variable holding the user's input label
                                }
                                for label, actual_col_name in found_extra_cols.items():
                                    entry[label] = row.get(actual_col_name)
                                # Add placeholders for missing extra columns if desired
                                # for label in missing_extra_cols:
                                #    entry[label] = None
                                master_data.append(entry)

                            # --- Collect Deleted Rows Info ---
                            for index, row in deleted_df.iterrows():
                                rec = row.to_dict()
                                rec["SourceFile"] = uploaded_file.name
                                rec["SourceSheet"] = sheet_name
                                all_deleted_rows.append(rec)

                            # --- Check Typical Column Location ---
                            # Use the 'typical_letter' variable which holds the user input (e.g., "M")
                            expected_col_index = col_letter_to_index(typical_letter)
                            if expected_col_index != -1: # Only check if letter was valid
                                try:
                                    actual_col_index = df_processed.columns.get_loc(value_col_name)
                                    if actual_col_index != expected_col_index:
                                        not_typical_col_loc.append({
                                            "File": uploaded_file.name,
                                            "Sheet": sheet_name,
                                            "MainColumnFound": value_col_name, # Adjusted key name for clarity
                                            "FoundIndex": actual_col_index,
                                            "ExpectedLetter": typical_letter, # Use the variable
                                            "ExpectedIndex": expected_col_index
                                        })
                                        logging.warning(f"Column '{value_col_name}' in '{uploaded_file.name}/{sheet_name}' was at index {actual_col_index}, not expected {expected_col_index} ('{typical_letter}').")
                                except KeyError:
                                     st.warning(f"Could not re-locate column '{value_col_name}' after processing sheet '{sheet_name}'. Skipping typical location check for this sheet.")
                                     logging.warning(f"KeyError finding location for '{value_col_name}' in processed columns for sheet '{sheet_name}'.")


                            # --- Generate Audit Sheet (using raw_df) ---
                            if audit_mode:
                                try:
                                    ws = workbook.add_worksheet(sheet_name[:31]) # Sheet name limit is 31 chars

                                    # Need to map original raw_df indices to know which rows were deleted
                                    # The indices in deleted_df correspond to the state *after* header rows were dropped by find_column
                                    # This makes direct highlighting complex if find_column modified the df significantly.
                                    # A simpler audit might just highlight columns in the *processed* df.
                                    # For now, let's highlight based on the *processed* data and deleted rows found *from it*.

                                    # Write Header
                                    for col_idx, col_name in enumerate(df_processed.columns):
                                         safe_cell_write(ws, 0, col_idx, col_name)

                                    # Write Processed Data (Highlight extracted columns)
                                    processed_row_offset = 1
                                    for row_idx, row_data in df_processed.iterrows():
                                        for col_idx, col_name in enumerate(df_processed.columns):
                                            cell_value = row_data[col_name]
                                            fmt = None
                                            if col_name == value_col_name or col_name == trans_col_name or col_name in found_extra_cols.values():
                                                fmt = extracted_fmt
                                            safe_cell_write(ws, row_idx + processed_row_offset, col_idx, cell_value, fmt)

                                    # Write Deleted Data (Highlight whole row)
                                    deleted_row_offset = processed_row_offset + len(df_processed) + 1 # Add gap
                                    if not deleted_df.empty:
                                        safe_cell_write(ws, deleted_row_offset -1 , 0, "--- Deleted Rows ---")
                                        # Write deleted header
                                        for col_idx, col_name in enumerate(deleted_df.columns):
                                            safe_cell_write(ws, deleted_row_offset, col_idx, col_name, deleted_fmt)
                                        # Write deleted rows
                                        for row_idx, row_data in deleted_df.iterrows():
                                            for col_idx, col_name in enumerate(deleted_df.columns):
                                                cell_value = row_data[col_name]
                                                safe_cell_write(ws, deleted_row_offset + row_idx + 1, col_idx, cell_value, deleted_fmt)

                                except Exception as e:
                                    st.error(f"Error creating audit sheet for '{uploaded_file.name}/{sheet_name}': {e}")
                                    error_logs.append(f"Audit Sheet Error ({uploaded_file.name}/{sheet_name}): {e}")
                                    logging.error(f"Error creating audit sheet for '{uploaded_file.name}/{sheet_name}': {e}", exc_info=True)

                        except Exception as e:
                            st.error(f"Error processing sheet '{sheet_name}' in file '{uploaded_file.name}': {e}")
                            skipped_sheets_info.append({"File": uploaded_file.name, "Sheet": sheet_name, "Reason": f"Processing error: {e}"})
                            error_logs.append(f"Sheet Processing Error ({uploaded_file.name}/{sheet_name}): {e}")
                            logging.error(f"Error processing sheet '{sheet_name}' in file '{uploaded_file.name}': {e}", exc_info=True)

                # After processing all sheets for the *current file*:
                # Close the in-memory ExcelWriter to save the audit data to the buffer
                # audit_writer is closed automatically by the 'with' statement

                # Add the content of the in-memory audit file buffer to the zip archive
                if audit_mode:
                     audit_bundle.writestr(f"{uploaded_file.name}_audit.xlsx", audit_excel_buffer.getvalue())
                audit_excel_buffer.close() # Close the buffer

            except Exception as e:
                st.error(f"Error opening or reading file '{uploaded_file.name}': {e}. Skipping this file.")
                error_logs.append(f"File Reading Error ({uploaded_file.name}): {e}")
                logging.error(f"Error opening or reading file '{uploaded_file.name}': {e}", exc_info=True)
                # Add entries for all potential sheets to skipped list? Maybe not necessary.

            # Update progress bar
            progress_bar.progress((file_index + 1) / total_files)

    # --- Final Export ---
    st.success(f"✅ Processing complete in {datetime.now() - start_time}!")

    # Export Master Report to an in-memory buffer
    output_excel_buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(output_excel_buffer, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors": True}}) as writer:
            # Use the user-provided labels as column headers in the output Excel
            master_df = pd.DataFrame(master_data)
            # Ensure the main columns exist even if missing in some files/sheets
            if value_label not in master_df.columns: master_df[value_label] = None
            if trans_label not in master_df.columns: master_df[trans_label] = None
            # Define column order, putting main ones first
            cols_order = ["SourceFile", "SourceSheet", value_label, trans_label] +                          [col for col in extra_cols if col in master_df.columns] +                          [col for col in master_df.columns if col not in ["SourceFile", "SourceSheet", value_label, trans_label] + extra_cols]
            master_df[cols_order].to_excel(writer, sheet_name="ExtractedData", index=False)

            if all_deleted_rows: # Only write if there are deleted rows
                pd.DataFrame(all_deleted_rows).to_excel(writer, sheet_name="DeletedRows", index=False)
            if not_typical_col_loc:
                 # Rename key for better readability in Excel output
                 df_not_typical = pd.DataFrame(not_typical_col_loc).rename(columns={"MainColumnFound": "ValueColumnFound"})
                 df_not_typical.to_excel(writer, sheet_name="ValueColNotTypical", index=False)
            if skipped_sheets_info:
                pd.DataFrame(skipped_sheets_info).to_excel(writer, sheet_name="SkippedSheets", index=False)
            if error_logs:
                 pd.DataFrame(error_logs, columns=["ErrorLog"]).to_excel(writer, sheet_name="ProcessingErrors", index=False)

        st.download_button(
            label="📥 Download Master Report",
            data=output_excel_buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error creating the final Master Report Excel file: {e}")
        logging.error(f"Error creating final Master Report: {e}", exc_info=True)

    # Provide download for the Audit ZIP file
    if audit_mode:
        audit_zip_buffer.seek(0) # Rewind the buffer before reading
        st.download_button(
            label="📥 Download Audit ZIP",
            data=audit_zip_buffer.getvalue(),
            file_name="audit_files.zip",
            mime="application/zip"
        )
    audit_zip_buffer.close() # Close the zip buffer

else:
    if uploaded_files:
        st.info("Click the '▶️ Run Excel Grabber' button to process the uploaded files.")
    else:
        st.info("☝️ Upload one or more Excel files using the uploader above, configure settings in the sidebar, then click '▶️ Run Excel Grabber'.")