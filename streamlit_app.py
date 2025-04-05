
import streamlit as st
import pandas as pd
import io
import zipfile
import datetime
import logging
import xlsxwriter # Ensure xlsxwriter is explicitly imported

# --- Page Configuration ---
st.set_page_config(
    page_title="MaxCode | Smart Excel Grabber",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- App Title ---
st.title("📊 MaxCode | Smart Excel Grabber")
st.markdown("Upload Excel files, extract specified columns, flag deleted rows & column mismatches.")

# --- Logging Setup ---
# Basic logging configuration - will log to console where Streamlit runs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ========================
# Helper Functions (Copied from Colab version)
# ========================

def col_letter_to_index(letter):
    """Convert Excel column letter to zero-based index."""
    letter = str(letter).strip().upper()
    result = 0
    if not letter: return -1 # Handle empty input
    for char in letter:
        if 'A' <= char <= 'Z':
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            logging.warning(f"Invalid character '{char}' in column letter '{letter}'. Cannot convert.")
            st.warning(f"Invalid character '{char}' in column letter '{letter}'. Cannot convert.")
            return -1 # Invalid index
    return result - 1

def find_header_and_columns(df_raw, required_labels, optional_labels, max_scan_rows):
    """
    Scans df_raw up to max_scan_rows to find a row containing required_labels.
    Returns:
        - header_row_index: Index (0-based) of the found header row, or None.
        - column_mapping: Dict mapping {user_label: actual_col_name_in_header}, or {}.
        - cleaned_columns: List of cleaned column names from the header row, or None.
    """
    if df_raw.empty:
        return None, {}, None

    num_rows_to_scan = min(max_scan_rows, len(df_raw))
    all_labels_to_find = [lbl for lbl in (required_labels + optional_labels) if lbl] # Filter out empty strings
    if not all_labels_to_find: # Nothing to search for
         # Assume first row is header if no labels are provided to search for
         header_row_index = 0
         header_values = df_raw.iloc[header_row_index].tolist() if len(df_raw) > 0 else []

         # Clean column names (convert to string, strip, handle duplicates)
         cleaned_columns = []
         counts = {}
         for idx, h in enumerate(header_values):
             col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
             if not col_name: col_name = f"Unnamed: {idx}"

             if col_name in counts:
                 counts[col_name] += 1
                 cleaned_columns.append(f"{col_name}.{counts[col_name]}")
             else:
                 counts[col_name] = 0
                 cleaned_columns.append(col_name)
         # No specific mapping needed if no labels were searched for
         return header_row_index if len(df_raw) > 0 else None, {}, cleaned_columns

    # Scan rows if specific labels are given
    for i in range(num_rows_to_scan):
        try:
            row_values_series = df_raw.iloc[i]
            # Convert to strings for searching, handle potential errors during conversion
            row_values_lower = [str(v).lower() if pd.notna(v) else '' for v in row_values_series]
        except Exception as e:
            logging.warning(f"Could not process row {i} as header candidate: {e}")
            continue

        # Check if *all* required labels are present (case-insensitive partial match)
        found_required_count = 0
        required_labels_lower = [req.lower() for req in required_labels if req]
        for req_label_lower in required_labels_lower:
            if any(req_label_lower in cell_lower for cell_lower in row_values_lower):
                found_required_count += 1

        # If all required labels are found in the row, consider it the header
        if found_required_count == len(required_labels_lower):
            header_row_index = i
            header_values = df_raw.iloc[header_row_index].tolist()

            # Clean column names (convert to string, strip, handle duplicates)
            cleaned_columns = []
            counts = {}
            for idx, h in enumerate(header_values):
                col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
                if not col_name: col_name = f"Unnamed: {idx}"

                if col_name in counts:
                    counts[col_name] += 1
                    cleaned_columns.append(f"{col_name}.{counts[col_name]}")
                else:
                    counts[col_name] = 0
                    cleaned_columns.append(col_name)

            # Map user labels to actual column names found in this header
            column_mapping = {}
            # Use original header values (converted to lower string) for mapping search
            header_values_lower_for_search = [str(v).lower() if pd.notna(v) else '' for v in header_values]
            for user_label in all_labels_to_find:
                user_label_lower = user_label.lower()
                found_match = None
                for idx, header_cell_lower in enumerate(header_values_lower_for_search):
                    if user_label_lower in header_cell_lower:
                        found_match = cleaned_columns[idx] # Use the cleaned name
                        break # Take first match
                column_mapping[user_label] = found_match

            logging.info(f"Header found at row index {header_row_index}. Mapping: {column_mapping}")
            return header_row_index, column_mapping, cleaned_columns

    logging.warning("Header row containing all required labels not found within scan limit.")
    return None, {}, None

def safe_cell_write(worksheet, row, col, value, cell_format=None):
    """Writes cell value to worksheet, handling potential type errors."""
    try:
        if pd.isna(value):
            worksheet.write_blank(row, col, None, cell_format)
        elif isinstance(value, (datetime.datetime, pd.Timestamp)):
            naive_datetime = value.tz_localize(None) if getattr(value, 'tzinfo', None) is not None else value
            try:
                 worksheet.write_datetime(row, col, naive_datetime, cell_format)
            except ValueError:
                 logging.warning(f"Could not write datetime '{value}' at ({row},{col}). Writing as string.")
                 worksheet.write_string(row, col, str(value), cell_format)
        elif isinstance(value, (int, float)):
            if pd.isna(value): # Handle numpy NaNs missed by first check
                 worksheet.write_blank(row, col, None, cell_format)
            else:
                 worksheet.write_number(row, col, value, cell_format)
        elif isinstance(value, bool):
            worksheet.write_boolean(row, col, value, cell_format)
        else:
            worksheet.write_string(row, col, str(value), cell_format)
    except Exception as e:
        logging.error(f"Error writing value '{value}' (type: {type(value)}) at ({row},{col}). Writing as blank. Error: {e}")
        try:
            worksheet.write_blank(row, col, None, cell_format) # Fallback to blank on error
        except Exception as final_e:
             logging.critical(f"FAILED even to write blank at ({row},{col}): {final_e}")


# ========================
# Streamlit UI Elements
# ========================

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Configuration")
value_label = st.sidebar.text_input("Main: Column Name to Grab", "รวมเงิน")
trans_label = st.sidebar.text_input("Secondary: Column Name to Grab", "รายการ")
typical_letter = st.sidebar.text_input("Main: Expected Column Letter", "M").upper()
extra_cols_raw = st.sidebar.text_area("Extra Columns (one per line)", "")
remove_phrases_raw = st.sidebar.text_area("Remove Rows if cell contains (one per line)", "TOTAL")
max_scan = st.sidebar.number_input("Header Scan Limit (Rows)", 1, 50, 10) # Increased max limit slightly
audit_mode = st.sidebar.checkbox("Generate Audit File ZIP?", value=True)
output_filename = st.sidebar.text_input("Master Report Filename", "Master_Report.xlsx")
audit_zip_filename = st.sidebar.text_input("Audit ZIP Filename", "Audit_Files.zip")

# Process multiline inputs
extra_cols_list = [col.strip() for col in extra_cols_raw.strip().splitlines() if col.strip()]
remove_phrases_list = [phrase.strip() for phrase in remove_phrases_raw.strip().splitlines() if phrase.strip()]

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "📂 Upload Excel Files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

# --- Main Execution Logic ---
if st.button("▶️ Process Uploaded Files") and uploaded_files:
    start_time = datetime.datetime.now()
    st.info(f"Processing started at {start_time.strftime('%Y-%m-%d %H:%M:%S')}...")

    # Placeholders for results
    master_data = []
    skipped_sheets_info = []
    not_typical_col_info = []
    deleted_rows_data = []
    error_logs = []
    audit_data_structure = {} # { orig_filename: { sheet_name: { df_raw: df, ... } } }

    # Prepare in-memory buffers
    master_output_buffer = io.BytesIO()
    audit_zip_buffer = io.BytesIO() # For the final ZIP

    # --- Input Validation ---
    if not value_label:
         st.error("❗ 'Main: Column Name to Grab' cannot be empty.")
         st.stop() # Stop execution if main label is missing

    typical_value_index = col_letter_to_index(typical_letter)
    if typical_value_index == -1:
        st.warning(f"Invalid typical column letter '{typical_letter}'. Typical check disabled.")
    value_col_not_letter_sheet_name = f"ValueColNot{typical_letter.upper()}" if typical_value_index != -1 else "ValueColTypicalCheck"
    remove_phrases_lower = [p.lower() for p in remove_phrases_list if p]

    # --- Progress Bar ---
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)
    files_processed = 0

    # --- File Processing Loop ---
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        files_processed += 1
        progress_text = f"Processing file {files_processed}/{total_files}: {file_name}"
        st.text(progress_text)
        logging.info(f"Processing file: {file_name}")

        try:
            # Read the entire Excel file from the uploaded buffer
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            audit_data_structure[file_name] = {} # Initialize audit structure for this file
        except Exception as e:
            st.error(f"❌ Error reading Excel file '{file_name}'. Skipping. Error: {e}")
            error_logs.append(f"File Read Error ({file_name}): {e}")
            logging.error(f"Error reading file '{file_name}': {e}", exc_info=True)
            progress_bar.progress(files_processed / total_files)
            continue # Skip to next file

        # --- Sheet Processing Loop ---
        for sheet_name in sheet_names:
            logging.info(f"-- Processing sheet: {sheet_name}")
            try:
                # Read raw data, don't assume header yet
                # Use dtype=object to prevent pandas from guessing types too early, esp. for headers
                df_raw = xls.parse(sheet_name, header=None, dtype=object)
                df_raw.index.name = 'OriginalRowIndex'
                df_raw.reset_index(inplace=True)

                # Find header and map columns
                required = [value_label] if value_label else [] # Should always have value_label here due to check above
                optional = ([trans_label] if trans_label else []) + extra_cols_list
                header_row_idx, column_map, cleaned_header = find_header_and_columns(
                    df_raw, required, optional, max_scan
                )

                # Check if essential value column was found
                actual_value_col = column_map.get(value_label)
                if header_row_idx is None or not actual_value_col:
                    reason = f"Required column '{value_label}' not found in header scan"
                    st.warning(f"⚠️ Skipping sheet '{sheet_name}' in '{file_name}': {reason}")
                    skipped_sheets_info.append({"File": file_name, "Sheet": sheet_name, "Reason": reason})
                    audit_data_structure[file_name][sheet_name] = {
                        'df_raw': df_raw.copy(), 'header_row': header_row_idx, 'deleted_indices': set(),
                        'col_map': {}, 'cleaned_header': None, 'typical_mismatch': False, 'skipped': True
                    }
                    continue # Skip this sheet

                # Identify rows to delete
                deleted_original_indices = set()
                if remove_phrases_lower:
                    data_rows_start_index = header_row_idx + 1
                    for r_idx in range(data_rows_start_index, len(df_raw)):
                        row_series = df_raw.iloc[r_idx]
                        # Important: Convert row to strings for robust checking across types
                        row_strs_lower = [str(cell).lower() if pd.notna(cell) else '' for cell in row_series]
                        if any(phrase in cell_str for cell_str in row_strs_lower for phrase in remove_phrases_lower):
                           original_index = df_raw.loc[r_idx, 'OriginalRowIndex']
                           deleted_original_indices.add(original_index)
                           # Store full deleted row data
                           deleted_rec = df_raw.iloc[r_idx].to_dict()
                           deleted_rec['FileName'] = file_name
                           deleted_rec['SheetName'] = sheet_name
                           renamed_deleted_rec = {'FileName': file_name, 'SheetName': sheet_name, 'OriginalRowIndex': original_index}
                           for h_idx, h_name in enumerate(cleaned_header):
                               # Access raw data by position using h_idx, map to cleaned_header name
                               renamed_deleted_rec[h_name] = deleted_rec.get(h_idx, None)
                           deleted_rows_data.append(renamed_deleted_rec)


                # Check typical column location
                typical_mismatch = False
                if actual_value_col and typical_value_index != -1 and cleaned_header:
                    try:
                        actual_value_col_index = cleaned_header.index(actual_value_col)
                        if actual_value_col_index != typical_value_index:
                            typical_mismatch = True
                            not_typical_col_info.append({
                                "File": file_name, "Sheet": sheet_name,
                                "ValueColumnFound": actual_value_col, "FoundIndex": actual_value_col_index,
                                "ExpectedLetter": typical_letter, "ExpectedIndex": typical_value_index
                            })
                            logging.warning(f"Typical mismatch in {file_name}/{sheet_name}: '{actual_value_col}' found at index {actual_value_col_index}, expected {typical_value_index} ('{typical_letter}')")
                    except (ValueError, IndexError):
                         logging.warning(f"Could not find '{actual_value_col}' in cleaned header list for typical check.")


                # Extract data for master report
                data_rows_start_index = header_row_idx + 1
                actual_trans_col = column_map.get(trans_label)
                extra_col_map = {lbl: column_map.get(lbl) for lbl in extra_cols_list}

                for r_idx in range(data_rows_start_index, len(df_raw)):
                    original_index = df_raw.loc[r_idx, 'OriginalRowIndex']
                    if original_index not in deleted_original_indices:
                        row_data = df_raw.iloc[r_idx] # This is a Series using original 0, 1, ... index
                        entry = {
                            "SourceFile": file_name,
                            "SourceSheet": sheet_name,
                            "OriginalRowIndex": original_index
                        }
                        try:
                             # Map cleaned header names back to their index in the cleaned_header list
                             # Use that index to get data from the raw row Series
                            val_col_idx_in_header = cleaned_header.index(actual_value_col) if actual_value_col else -1
                            entry[value_label] = row_data.iloc[val_col_idx_in_header] if val_col_idx_in_header != -1 else None

                            trans_col_idx_in_header = cleaned_header.index(actual_trans_col) if actual_trans_col else -1
                            entry[trans_label] = row_data.iloc[trans_col_idx_in_header] if trans_col_idx_in_header != -1 else None

                            for label, actual_col in extra_col_map.items():
                                ex_col_idx_in_header = cleaned_header.index(actual_col) if actual_col else -1
                                entry[label] = row_data.iloc[ex_col_idx_in_header] if ex_col_idx_in_header != -1 else None

                            master_data.append(entry)
                        except (ValueError, IndexError) as lookup_error:
                             err_msg = f"Column Mapping/Lookup Error ({file_name}/{sheet_name}/Row {original_index}): {lookup_error}"
                             logging.error(err_msg)
                             error_logs.append(err_msg)
                             # Add placeholders for this row in master data? Or skip row? Skipping for now.
                             st.warning(f"⚠️ {err_msg}. Skipping row.")


                # Store data needed for audit highlighting
                audit_data_structure[file_name][sheet_name] = {
                    'df_raw': df_raw.copy(),
                    'header_row': header_row_idx,
                    'deleted_indices': deleted_original_indices,
                    'col_map': column_map,
                    'cleaned_header': cleaned_header,
                    'typical_mismatch': typical_mismatch,
                    'skipped': False
                }

            except Exception as e:
                st.error(f"❌ Error processing sheet '{sheet_name}' in file '{file_name}'. Skipping sheet. Error: {e}")
                skipped_sheets_info.append({"File": file_name, "Sheet": sheet_name, "Reason": f"Processing error: {e}"})
                error_logs.append(f"Sheet Processing Error ({file_name}/{sheet_name}): {e}")
                logging.error(f"Error processing sheet '{sheet_name}' in file '{file_name}': {e}", exc_info=True)
                audit_data_structure[file_name][sheet_name] = {
                     'df_raw': pd.DataFrame(), 'header_row': None, 'deleted_indices': set(),
                     'col_map': {}, 'cleaned_header': None, 'typical_mismatch': False, 'skipped': True, 'error': str(e)
                }
        # --- End Sheet Loop ---
        progress_bar.progress(files_processed / total_files)
    # --- End File Loop ---

    st.success(f"✅ File processing completed in {datetime.datetime.now() - start_time}.")

    # ============================
    # Combine & Analyze Master Data
    # ============================
    master_df = pd.DataFrame(master_data) if master_data else pd.DataFrame()
    skipped_df = pd.DataFrame(skipped_sheets_info) if skipped_sheets_info else pd.DataFrame()
    not_typical_df = pd.DataFrame(not_typical_col_info) if not_typical_col_info else pd.DataFrame()
    deleted_df = pd.DataFrame(deleted_rows_data) if deleted_rows_data else pd.DataFrame()
    errors_df = pd.DataFrame(error_logs, columns=["ErrorLog"]) if error_logs else pd.DataFrame()

    analysis_df = pd.DataFrame()
    if not master_df.empty and value_label and value_label in master_df.columns:
        numeric_col_name = f"{value_label}_numeric"
        # Convert to string first to handle mixed types before replacing comma
        master_df[numeric_col_name] = pd.to_numeric(
            master_df[value_label].astype(str).str.replace(',', '', regex=False),
            errors='coerce'
        )
        if numeric_col_name in master_df.columns and master_df[numeric_col_name].notna().any():
             analysis_df = pd.DataFrame(master_df[numeric_col_name].describe())

    # ============================
    # Write Master Report to Buffer
    # ============================
    master_report_generated = False
    try:
        with pd.ExcelWriter(master_output_buffer, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors": True}}) as writer:
            if not master_df.empty:
                 cols_order = ["SourceFile", "SourceSheet", "OriginalRowIndex"]
                 if value_label: cols_order.append(value_label)
                 if trans_label: cols_order.append(trans_label)
                 cols_order.extend(extra_cols_list)
                 if f"{value_label}_numeric" in master_df.columns: cols_order.append(f"{value_label}_numeric")
                 cols_order.extend([col for col in master_df.columns if col not in cols_order])
                 master_df.to_excel(writer, sheet_name='AllData', index=False, columns=cols_order)
            else: # Write empty sheet if no data extracted
                 pd.DataFrame().to_excel(writer, sheet_name='AllData', index=False)

            # Write other info sheets only if they have data
            if not skipped_df.empty: skipped_df.to_excel(writer, sheet_name='SkippedSheets', index=False)
            if not not_typical_df.empty: not_typical_df.to_excel(writer, sheet_name=value_col_not_letter_sheet_name, index=False)
            if not deleted_df.empty:
                deleted_cols = ['FileName', 'SheetName', 'OriginalRowIndex'] + [col for col in deleted_df.columns if col not in ['FileName', 'SheetName', 'OriginalRowIndex']]
                deleted_df.to_excel(writer, sheet_name='DeletedRows', index=False, columns=deleted_cols)
            if not analysis_df.empty: analysis_df.to_excel(writer, sheet_name='ValueColumnAnalysis')
            if not errors_df.empty: errors_df.to_excel(writer, sheet_name='ProcessingErrors', index=False)
        master_report_generated = True
        logging.info("Master report buffer generated successfully.")
    except Exception as e:
        st.error(f"❌ Error creating Master Report Excel buffer: {e}")
        logging.error(f"Error creating Master Report buffer: {e}", exc_info=True)


    # ============================
    # Generate Audit Files + ZIP (if enabled)
    # ============================
    audit_zip_generated = False
    if audit_mode:
        files_audited = 0
        st.info("⚙️ Generating Audit Files ZIP...")
        audit_progress = st.progress(0)
        audit_file_count = sum(len(sheets) for sheets in audit_data_structure.values())
        audit_sheets_processed = 0

        try:
            with zipfile.ZipFile(audit_zip_buffer, "w", zipfile.ZIP_DEFLATED) as audit_zip:
                for orig_filename, sheets_data in audit_data_structure.items():
                    # Use an in-memory buffer for each individual audit Excel file
                    audit_excel_buffer = io.BytesIO()
                    audit_file_has_content = False
                    try:
                        with pd.ExcelWriter(audit_excel_buffer, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors":True}}) as writer:
                            workbook = writer.book
                            # Define formats
                            deleted_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Red fill
                            extracted_fmt = workbook.add_format({'bg_color': '#FFEB9C'}) # Yellow fill
                            mismatch_fmt = workbook.add_format({'bg_color': '#E4DFEC'}) # Purple fill (value col mismatch)
                            header_fmt = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True}) # Grey fill, bold
                            skipped_fmt = workbook.add_format({'font_color': '#A6A6A6'}) # Grey text for skipped

                            for sheet_name, audit_info in sheets_data.items():
                                audit_sheets_processed += 1
                                audit_file_has_content = True # Assume content unless df is empty AND not skipped/errored
                                df_audit = audit_info.get('df_raw', pd.DataFrame())

                                ws = workbook.add_worksheet(sheet_name[:31])
                                writer.sheets[sheet_name[:31]] = ws # Associate sheet object

                                if df_audit.empty and not audit_info.get('skipped'):
                                    audit_file_has_content = False # No real content
                                    safe_cell_write(ws, 0, 0, "(Sheet was empty in source)", skipped_fmt)
                                    continue # Skip writing empty sheet unless it was skipped/errored

                                # Extract highlight info
                                header_row = audit_info.get('header_row')
                                deleted_indices = audit_info.get('deleted_indices', set())
                                col_map = audit_info.get('col_map', {})
                                cleaned_header = audit_info.get('cleaned_header')
                                typical_mismatch = audit_info.get('typical_mismatch', False)
                                is_skipped = audit_info.get('skipped', False)
                                error_msg = audit_info.get('error')

                                if is_skipped or error_msg:
                                     reason = f"Sheet Skipped ({audit_info.get('reason', 'Unknown')})" if is_skipped else f"Sheet Errored: {error_msg}"
                                     safe_cell_write(ws, 0, 0, reason, skipped_fmt)
                                     # Optionally write raw data below if available
                                     if not df_audit.empty:
                                         for c_idx_raw, col_name_raw in enumerate(df_audit.columns):
                                              safe_cell_write(ws, 2, c_idx_raw, col_name_raw, header_fmt)
                                              for r_idx_raw in range(len(df_audit)):
                                                  safe_cell_write(ws, r_idx_raw + 3, c_idx_raw, df_audit.iloc[r_idx_raw, c_idx_raw], skipped_fmt)
                                     continue # Don't apply other highlights

                                # Get actual column names to highlight
                                actual_value_col = col_map.get(value_label)
                                actual_trans_col = col_map.get(trans_label)
                                actual_extra_cols = [col_map.get(lbl) for lbl in extra_cols_list if col_map.get(lbl)]

                                # Find indices in the *cleaned header* for mapping highlights
                                col_indices_in_cleaned_header = {}
                                if cleaned_header:
                                     try:
                                         if actual_value_col: col_indices_in_cleaned_header[actual_value_col] = cleaned_header.index(actual_value_col)
                                         if actual_trans_col: col_indices_in_cleaned_header[actual_trans_col] = cleaned_header.index(actual_trans_col)
                                         for col in actual_extra_cols: col_indices_in_cleaned_header[col] = cleaned_header.index(col)
                                     except (ValueError, IndexError):
                                         logging.warning(f"Audit Highlight Warning: Could not map all column names back to cleaned header indices in {orig_filename}/{sheet_name}")

                                # Write header (original column names/indices from df_raw)
                                raw_cols = df_audit.columns.tolist()
                                for c_idx, col_name in enumerate(raw_cols):
                                     safe_cell_write(ws, 0, c_idx, col_name, header_fmt if header_row==0 else None)

                                # Write data rows
                                for r_idx in range(len(df_audit)):
                                    row_data = df_audit.iloc[r_idx]
                                    original_row_idx = row_data.get('OriginalRowIndex', -1) # Get original index if column exists
                                    current_format = None

                                    # Determine base row format
                                    if original_row_idx in deleted_indices:
                                        current_format = deleted_fmt
                                    elif header_row is not None and r_idx == header_row:
                                        current_format = header_fmt

                                    # Write cells for this row
                                    for c_idx, col_name in enumerate(raw_cols):
                                        cell_value = row_data.iloc[c_idx]
                                        cell_format_to_use = current_format # Start with row format

                                        # Apply column-specific highlights (only if it's a data row)
                                        is_data_row = (header_row is not None and r_idx > header_row and original_row_idx not in deleted_indices)
                                        if is_data_row and cleaned_header and c_idx < len(cleaned_header): # Check index bounds
                                            current_cleaned_col_name = cleaned_header[c_idx] # Get cleaned name for this column index
                                            col_is_value = (actual_value_col == current_cleaned_col_name)
                                            col_is_trans = (actual_trans_col == current_cleaned_col_name)
                                            col_is_extra = (current_cleaned_col_name in actual_extra_cols)

                                            if col_is_value and typical_mismatch:
                                                cell_format_to_use = mismatch_fmt # Purple overrides yellow
                                            elif col_is_value or col_is_trans or col_is_extra:
                                                cell_format_to_use = extracted_fmt # Yellow

                                        safe_cell_write(ws, r_idx + 1, c_idx, cell_value, cell_format_to_use) # +1 for header offset

                                # Update audit progress bar
                                if audit_file_count > 0:
                                     audit_progress.progress(audit_sheets_processed / audit_file_count)

                        # Add the generated Excel file (in memory) to the zip archive
                        audit_filename_in_zip = f"{orig_filename.rsplit('.', 1)[0]}_audit.xlsx"
                        if audit_file_has_content:
                            audit_zip.writestr(audit_filename_in_zip, audit_excel_buffer.getvalue())
                            files_audited += 1
                            logging.info(f"Added audit file to ZIP: {audit_filename_in_zip}")
                        else:
                             logging.info(f"Skipped adding empty/errored audit file to ZIP: {audit_filename_in_zip}")

                    except Exception as sheet_audit_error:
                         st.error(f"❌ Error creating audit Excel buffer for '{orig_filename}': {sheet_audit_error}")
                         error_logs.append(f"Audit File Creation Error ({orig_filename}): {sheet_audit_error}")
                         logging.error(f"Error creating audit Excel buffer for '{orig_filename}': {sheet_audit_error}", exc_info=True)
                    finally:
                        audit_excel_buffer.close() # Close the individual buffer

            audit_zip_generated = (files_audited > 0)
            audit_progress.progress(1.0) # Mark as complete
            if audit_zip_generated:
                 st.success(f"✅ Audit ZIP generated with {files_audited} file(s).")
                 logging.info(f"Audit ZIP buffer generated with {files_audited} file(s).")
            else:
                 st.warning("⚠️ No valid audit files were generated.")

        except Exception as zip_error:
            st.error(f"❌ Error creating Audit ZIP file: {zip_error}")
            logging.error(f"Error creating Audit ZIP: {zip_error}", exc_info=True)

    # ============================
    # Display Download Buttons
    # ============================
    st.markdown("---")
    st.subheader("⬇️ Download Results")

    col1, col2 = st.columns(2)

    with col1:
        if master_report_generated:
            st.download_button(
                label="📥 Download Master Report",
                data=master_output_buffer.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_master" # Unique key
            )
            st.caption(f"Filename: `{output_filename}`")
        else:
            st.error("Master Report generation failed.")

    with col2:
        if audit_mode:
            if audit_zip_generated:
                st.download_button(
                    label="📥 Download Audit ZIP",
                    data=audit_zip_buffer.getvalue(),
                    file_name=audit_zip_filename,
                    mime="application/zip",
                    key="download_audit" # Unique key
                )
                st.caption(f"Filename: `{audit_zip_filename}`")
            else:
                st.warning("Audit ZIP generation failed or was skipped.")
        else:
            st.info("Audit file generation was disabled.")


else:
    # Initial state or after reset
    st.info("☝️ Configure settings in the sidebar, upload Excel file(s), and click 'Process Uploaded Files'.")

# Add a footer or link (optional)
st.markdown("---")
st.caption("MaxCode Smart Excel Grabber")

