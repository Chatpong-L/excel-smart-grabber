
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
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ========================
# Helper Functions
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
    all_labels_to_find = [lbl for lbl in (required_labels + optional_labels) if lbl]

    if not all_labels_to_find: # Assume first row is header if no labels provided
         header_row_index = 0 if len(df_raw) > 0 else None
         if header_row_index is None: return None, {}, []

         header_values = df_raw.iloc[header_row_index].tolist()
         cleaned_columns = []
         counts = {}
         for idx, h in enumerate(header_values):
             col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
             if not col_name: col_name = f"Unnamed: {idx}"
             if col_name in counts: counts[col_name] += 1; cleaned_columns.append(f"{col_name}.{counts[col_name]}")
             else: counts[col_name] = 0; cleaned_columns.append(col_name)
         return header_row_index, {}, cleaned_columns

    # Scan rows if specific labels are given
    for i in range(num_rows_to_scan):
        try:
            row_values_series = df_raw.iloc[i]
            row_values_lower = [str(v).lower() if pd.notna(v) else '' for v in row_values_series]
        except Exception as e:
            logging.warning(f"Could not process row {i} as header candidate: {e}")
            continue

        required_labels_lower = [req.lower() for req in required_labels if req]
        found_required_count = sum(1 for req_label_lower in required_labels_lower if any(req_label_lower in cell_lower for cell_lower in row_values_lower))

        if found_required_count == len(required_labels_lower):
            header_row_index = i
            header_values = df_raw.iloc[header_row_index].tolist()

            cleaned_columns = []
            counts = {}
            for idx, h in enumerate(header_values):
                col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
                if not col_name: col_name = f"Unnamed: {idx}"
                if col_name in counts: counts[col_name] += 1; cleaned_columns.append(f"{col_name}.{counts[col_name]}")
                else: counts[col_name] = 0; cleaned_columns.append(col_name)

            column_mapping = {}
            header_values_lower_for_search = [str(v).lower() if pd.notna(v) else '' for v in header_values]
            for user_label in all_labels_to_find:
                user_label_lower = user_label.lower()
                found_match = None
                for idx, header_cell_lower in enumerate(header_values_lower_for_search):
                    # Find first occurrence in the header row
                    if user_label_lower in header_cell_lower:
                        found_match = cleaned_columns[idx] # Use the cleaned name
                        break
                column_mapping[user_label] = found_match

            logging.info(f"Header found at row index {header_row_index}. Mapping: {column_mapping}")
            return header_row_index, column_mapping, cleaned_columns

    logging.warning("Header row containing all required labels not found within scan limit.")
    return None, {}, None


def safe_cell_write(worksheet, row, col, value, cell_format=None):
    """Writes cell value to worksheet, handling potential type errors. Needed for audit highlighting."""
    try:
        if pd.isna(value):
            worksheet.write_blank(row, col, None, cell_format)
        elif isinstance(value, (datetime.datetime, pd.Timestamp)):
            naive_datetime = value.tz_localize(None) if getattr(value, 'tzinfo', None) is not None else value
            try:
                 worksheet.write_datetime(row, col, naive_datetime, cell_format)
            except ValueError: # Handle dates outside Excel's range
                 worksheet.write_string(row, col, str(value), cell_format)
        elif isinstance(value, (int, float)):
             if pd.isna(value): worksheet.write_blank(row, col, None, cell_format)
             else: worksheet.write_number(row, col, value, cell_format)
        elif isinstance(value, bool):
            worksheet.write_boolean(row, col, value, cell_format)
        else:
            # Ensure the string is not too long for Excel
            str_value = str(value)
            max_len = 32767
            if len(str_value) > max_len:
                worksheet.write_string(row, col, str_value[:max_len-3] + "...", cell_format)
            else:
                 worksheet.write_string(row, col, str_value, cell_format)
    except Exception as e:
        logging.error(f"Error writing value '{str(value)[:50]}' (type: {type(value)}) at ({row},{col}). Writing as blank. Error: {e}")
        try: worksheet.write_blank(row, col, None, cell_format)
        except: pass # Ignore error writing blank


# ========================
# Streamlit UI Elements
# ========================

st.sidebar.header("⚙️ Configuration")
value_label = st.sidebar.text_input("Main: Column Name to Grab", "รวมเงิน", key="val_label")
trans_label = st.sidebar.text_input("Secondary: Column Name to Grab", "รายการ", key="trans_label")
typical_letter = st.sidebar.text_input("Main: Expected Column Letter", "M", key="typ_letter").upper()
extra_cols_raw = st.sidebar.text_area("Extra Columns (one per line)", "", key="extra_cols")
remove_phrases_raw = st.sidebar.text_area("Remove Rows if cell contains (one per line)", "TOTAL", key="remove_phrases")
max_scan = st.sidebar.number_input("Header Scan Limit (Rows)", 1, 50, 10, key="max_scan")
audit_mode = st.sidebar.checkbox("Generate Audit File?", value=True, key="audit_mode")
output_filename = st.sidebar.text_input("Master Report Filename", "Master_Report.xlsx", key="out_fname")
audit_zip_filename_internal = "Audit_Files.zip" # Internal name for audit zip
combined_zip_filename = st.sidebar.text_input("Combined Download Filename (if Audit)", "Processing_Results.zip", key="comb_fname")

extra_cols_list = [col.strip() for col in extra_cols_raw.strip().splitlines() if col.strip()]
remove_phrases_list = [phrase.strip() for phrase in remove_phrases_raw.strip().splitlines() if phrase.strip()]

uploaded_files = st.file_uploader(
    "📂 Upload Excel Files (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="uploader"
)

# Initialize session state for download data
if 'download_data' not in st.session_state:
    st.session_state.download_data = None
if 'download_filename' not in st.session_state:
    st.session_state.download_filename = None
if 'download_mime' not in st.session_state:
    st.session_state.download_mime = None
if 'download_ready' not in st.session_state:
    st.session_state.download_ready = False

# --- Main Execution Logic ---
if st.button("▶️ Process Uploaded Files", key="run_button") and uploaded_files:
    # Reset download state on new run
    st.session_state.download_ready = False
    st.session_state.download_data = None

    start_time = datetime.datetime.now()
    st.info(f"Processing started at {start_time.strftime('%Y-%m-%d %H:%M:%S')}...")

    # Placeholders
    master_data, skipped_sheets_info, not_typical_col_info, deleted_rows_data, error_logs = [], [], [], [], []
    audit_data_structure = {}

    # Buffers
    master_output_buffer = io.BytesIO()
    audit_zip_buffer = io.BytesIO() # Only used if audit_mode is True
    final_zip_buffer = io.BytesIO() # Used for combined download

    # --- Input Validation ---
    if not value_label:
         st.error("❗ 'Main: Column Name to Grab' cannot be empty.")
         st.stop()

    typical_value_index = col_letter_to_index(typical_letter)
    if typical_value_index == -1 and typical_letter: # Only warn if they provided an invalid letter
        st.warning(f"Invalid typical column letter '{typical_letter}'. Typical check disabled.")
    value_col_not_letter_sheet_name = f"ValueColNot{typical_letter.upper()}" if typical_value_index != -1 else "ValueColTypicalCheck"
    remove_phrases_lower = [p.lower() for p in remove_phrases_list if p]

    # --- Progress Bar ---
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_files = len(uploaded_files)
    files_processed = 0

    # --- File Processing Loop ---
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        files_processed += 1
        progress_text = f"Processing file {files_processed}/{total_files}: {file_name}"
        status_text.text(progress_text)
        logging.info(f"Processing file: {file_name}")

        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            audit_data_structure[file_name] = {}
        except Exception as e:
            st.error(f"❌ Error reading Excel file '{file_name}'. Skipping. Error: {e}")
            error_logs.append(f"File Read Error ({file_name}): {e}")
            logging.error(f"Error reading file '{file_name}': {e}", exc_info=True)
            progress_bar.progress(files_processed / total_files)
            continue

        # --- Sheet Processing Loop ---
        for sheet_name in sheet_names:
            logging.info(f"-- Processing sheet: {sheet_name}")
            try:
                df_raw = xls.parse(sheet_name, header=None, dtype=object)
                # IMPORTANT: Keep original index before resetting
                df_raw['OriginalIndexInDataframe'] = df_raw.index
                df_raw.index.name = 'OriginalExcelRow' # Index reflects Excel row number (starting 0)
                df_raw.reset_index(inplace=True) # Make Excel row number a column

                required = [value_label]
                optional = ([trans_label] if trans_label else []) + extra_cols_list
                header_row_idx, column_map, cleaned_header = find_header_and_columns(
                    df_raw, required, optional, max_scan
                )

                actual_value_col = column_map.get(value_label)
                if header_row_idx is None or not actual_value_col:
                    reason = f"Required column '{value_label}' not found in header scan"
                    st.warning(f"⚠️ Skipping sheet '{sheet_name}' in '{file_name}': {reason}")
                    skipped_sheets_info.append({"File": file_name, "Sheet": sheet_name, "Reason": reason})
                    audit_data_structure[file_name][sheet_name] = {
                        'df_raw': df_raw.copy(), 'header_row': header_row_idx, 'deleted_excel_rows': set(),
                        'col_map': {}, 'cleaned_header': None, 'typical_mismatch': False, 'skipped': True
                    }
                    continue

                # Identify rows to delete (using OriginalExcelRow for identification)
                deleted_excel_rows = set()
                if remove_phrases_lower:
                    data_rows_start_index_in_df = header_row_idx + 1 # Start checking after header row in df_raw
                    for df_idx in range(data_rows_start_index_in_df, len(df_raw)):
                        row_series = df_raw.iloc[df_idx]
                        row_strs_lower = [str(cell).lower() if pd.notna(cell) else '' for cell in row_series]
                        if any(phrase in cell_str for cell_str in row_strs_lower for phrase in remove_phrases_lower):
                           original_excel_row = df_raw.loc[df_idx, 'OriginalExcelRow']
                           deleted_excel_rows.add(original_excel_row)
                           # Store full deleted row data using cleaned header names
                           deleted_rec = df_raw.iloc[df_idx].to_dict()
                           renamed_deleted_rec = {'FileName': file_name, 'SheetName': sheet_name, 'OriginalExcelRow': original_excel_row}
                           if cleaned_header: # Map raw data (0, 1, ...) to cleaned names
                               for h_idx, h_name in enumerate(cleaned_header):
                                   renamed_deleted_rec[h_name] = deleted_rec.get(h_idx, None)
                           else: # Fallback if no cleaned header (e.g., took first row)
                               renamed_deleted_rec.update({k:v for k,v in deleted_rec.items() if k not in ['FileName', 'SheetName', 'OriginalExcelRow']})
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
                    except (ValueError, IndexError): pass # Ignore if column not found in cleaned list

                # Extract data for master report
                data_rows_start_index_in_df = header_row_idx + 1
                actual_trans_col = column_map.get(trans_label)
                extra_col_map = {lbl: column_map.get(lbl) for lbl in extra_cols_list}

                for df_idx in range(data_rows_start_index_in_df, len(df_raw)):
                    original_excel_row = df_raw.loc[df_idx, 'OriginalExcelRow']
                    if original_excel_row not in deleted_excel_rows:
                        row_data = df_raw.iloc[df_idx] # Raw data series (indexed by 0, 1, ...)
                        entry = {
                            "SourceFile": file_name, "SourceSheet": sheet_name,
                            "OriginalExcelRow": original_excel_row # Use Excel row for reference
                        }
                        try:
                            # Map required cols using cleaned_header index lookup
                            val_col_idx_in_header = cleaned_header.index(actual_value_col) if actual_value_col and cleaned_header else -1
                            entry[value_label] = row_data.iloc[val_col_idx_in_header] if val_col_idx_in_header != -1 else None

                            trans_col_idx_in_header = cleaned_header.index(actual_trans_col) if actual_trans_col and cleaned_header else -1
                            entry[trans_label] = row_data.iloc[trans_col_idx_in_header] if trans_col_idx_in_header != -1 else None

                            for label, actual_col in extra_col_map.items():
                                ex_col_idx_in_header = cleaned_header.index(actual_col) if actual_col and cleaned_header else -1
                                entry[label] = row_data.iloc[ex_col_idx_in_header] if ex_col_idx_in_header != -1 else None

                            master_data.append(entry)
                        except (ValueError, IndexError) as lookup_error:
                             err_msg = f"Column Mapping/Lookup Error ({file_name}/{sheet_name}/Row {original_excel_row}): {lookup_error}"
                             logging.error(err_msg)
                             error_logs.append(err_msg)
                             st.warning(f"⚠️ {err_msg}. Skipping row.")

                # Store data needed for audit highlighting
                audit_data_structure[file_name][sheet_name] = {
                    'df_raw': df_raw.copy(), # Contains 'OriginalExcelRow'
                    'header_row_in_df': header_row_idx, # Index within df_raw where header was found
                    'deleted_excel_rows': deleted_excel_rows, # Set of OriginalExcelRow values
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
                     'df_raw': pd.DataFrame(), 'header_row_in_df': None, 'deleted_excel_rows': set(),
                     'col_map': {}, 'cleaned_header': None, 'typical_mismatch': False, 'skipped': True, 'error': str(e)
                }
        # --- End Sheet Loop ---
        progress_bar.progress(files_processed / total_files)
    # --- End File Loop ---

    status_text.success(f"✅ File processing completed in {datetime.datetime.now() - start_time}.")

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
        master_df[numeric_col_name] = pd.to_numeric(master_df[value_label].astype(str).str.replace(',', '', regex=False), errors='coerce')
        if numeric_col_name in master_df.columns and master_df[numeric_col_name].notna().any():
             analysis_df = pd.DataFrame(master_df[numeric_col_name].describe())

    # ============================
    # Write Master Report to Buffer
    # ============================
    master_report_generated = False
    try:
        with pd.ExcelWriter(master_output_buffer, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors": True}}) as writer:
            if not master_df.empty:
                 cols_order = ["SourceFile", "SourceSheet", "OriginalExcelRow"] # Use Excel Row
                 if value_label: cols_order.append(value_label)
                 if trans_label: cols_order.append(trans_label)
                 cols_order.extend(extra_cols_list)
                 if f"{value_label}_numeric" in master_df.columns: cols_order.append(f"{value_label}_numeric")
                 cols_order.extend([col for col in master_df.columns if col not in cols_order])
                 master_df.to_excel(writer, sheet_name='AllData', index=False, columns=cols_order)
            else: pd.DataFrame().to_excel(writer, sheet_name='AllData', index=False)

            if not skipped_df.empty: skipped_df.to_excel(writer, sheet_name='SkippedSheets', index=False)
            if not not_typical_df.empty: not_typical_df.to_excel(writer, sheet_name=value_col_not_letter_sheet_name, index=False)
            if not deleted_df.empty:
                # Use OriginalExcelRow in deleted sheet too
                deleted_cols = ['FileName', 'SheetName', 'OriginalExcelRow'] + [col for col in deleted_df.columns if col not in ['FileName', 'SheetName', 'OriginalExcelRow']]
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
        status_text.info("⚙️ Generating Audit Files ZIP...")
        audit_progress = st.progress(0)
        total_audit_sheets = sum(len(sheets) for sheets in audit_data_structure.values())
        audit_sheets_processed = 0

        try:
            with zipfile.ZipFile(audit_zip_buffer, "w", zipfile.ZIP_DEFLATED) as audit_zip:
                for orig_filename, sheets_data in audit_data_structure.items():
                    audit_excel_buffer = io.BytesIO()
                    audit_file_has_content = False
                    try:
                        # Use ExcelWriter for formatting control
                        with pd.ExcelWriter(audit_excel_buffer, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
                            workbook = writer.book
                            red_fill = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
                            yellow_fill = workbook.add_format({"bg_color": "#FFEB9C"})
                            purple_fill = workbook.add_format({"bg_color": "#E4DFEC"})
                            header_fmt = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True})
                            skipped_fmt = workbook.add_format({'font_color': '#A6A6A6'})

                            for sheet_name, audit_info in sheets_data.items():
                                audit_sheets_processed += 1
                                audit_file_has_content = True # Assume content initially
                                df_audit_raw = audit_info.get('df_raw', pd.DataFrame())
                                if df_audit_raw.empty: continue # Skip empty raw DFs

                                ws = workbook.add_worksheet(sheet_name[:31])
                                writer.sheets[sheet_name[:31]] = ws

                                # --- Apply Highlighting Logic (adapted from Colab) ---
                                header_row_in_df = audit_info.get('header_row_in_df') # Index in df_audit_raw
                                deleted_excel_rows = audit_info.get('deleted_excel_rows', set())
                                col_map = audit_info.get('col_map', {})
                                cleaned_header = audit_info.get('cleaned_header')
                                typical_mismatch = audit_info.get('typical_mismatch', False)
                                is_skipped = audit_info.get('skipped', False)
                                error_msg = audit_info.get('error')

                                if is_skipped or error_msg: # Handle skipped/errored sheets first
                                    reason = f"Sheet Skipped ({audit_info.get('reason', 'Unknown')})" if is_skipped else f"Sheet Errored: {error_msg}"
                                    safe_cell_write(ws, 0, 0, reason, skipped_fmt)
                                    # Optionally write raw data below if available, marked as skipped
                                    raw_cols = df_audit_raw.columns.tolist()
                                    for c_idx, col_name in enumerate(raw_cols):
                                        safe_cell_write(ws, 2, c_idx, col_name, skipped_fmt)
                                        for r_idx in range(len(df_audit_raw)):
                                            safe_cell_write(ws, r_idx + 3, c_idx, df_audit_raw.iloc[r_idx, c_idx], skipped_fmt)
                                    continue # Skip normal highlighting

                                # Write data first (cell by cell for format control)
                                raw_cols = df_audit_raw.columns.tolist()
                                for c_idx, col_name in enumerate(raw_cols):
                                    safe_cell_write(ws, 0, c_idx, col_name) # Write header (raw names: 0, 1, ..)
                                for r_idx in range(len(df_audit_raw)):
                                    for c_idx, col_name in enumerate(raw_cols):
                                        safe_cell_write(ws, r_idx + 1, c_idx, df_audit_raw.iloc[r_idx, c_idx]) # +1 for header

                                # Apply formats row by row, then column by column
                                actual_value_col = col_map.get(value_label)
                                actual_trans_col = col_map.get(trans_label)
                                actual_extra_cols = [col_map.get(lbl) for lbl in extra_cols_list if col_map.get(lbl)]
                                used_cols = [c for c in [actual_value_col, actual_trans_col] + actual_extra_cols if c]

                                value_col_idx_in_cleaned = -1
                                used_col_indices_in_cleaned = {}
                                if cleaned_header:
                                    try:
                                        if actual_value_col: value_col_idx_in_cleaned = cleaned_header.index(actual_value_col)
                                        for col in used_cols: used_col_indices_in_cleaned[col] = cleaned_header.index(col)
                                    except ValueError: pass # Ignore if col not found

                                # Apply formats
                                for r_idx in range(len(df_audit_raw)):
                                    original_excel_row = df_audit_raw.loc[r_idx, 'OriginalExcelRow']
                                    is_deleted = original_excel_row in deleted_excel_rows
                                    is_header = header_row_in_df is not None and r_idx == header_row_in_df

                                    # 1. Apply row-level formats (Red for deleted, Grey for header)
                                    row_format = None
                                    if is_deleted:
                                        row_format = red_fill
                                    elif is_header:
                                        row_format = header_fmt

                                    if row_format:
                                        # Apply format to existing cells in the row
                                        for c_idx in range(len(raw_cols)):
                                             cell_val = df_audit_raw.iloc[r_idx, c_idx]
                                             safe_cell_write(ws, r_idx + 1, c_idx, cell_val, row_format)

                                    # 2. Apply column-level formats (Yellow/Purple) only if NOT deleted and NOT header
                                    if not is_deleted and not is_header and cleaned_header:
                                        for c_idx in range(min(len(raw_cols), len(cleaned_header))): # Iterate based on cleaned header length
                                            current_cleaned_col = cleaned_header[c_idx]
                                            cell_val = df_audit_raw.iloc[r_idx, c_idx]
                                            col_format = None

                                            is_value_col = (value_col_idx_in_cleaned == c_idx)
                                            is_used_col = (current_cleaned_col in used_cols)

                                            if is_value_col and typical_mismatch:
                                                col_format = purple_fill # Purple overrides yellow
                                            elif is_used_col:
                                                col_format = yellow_fill

                                            if col_format:
                                                safe_cell_write(ws, r_idx + 1, c_idx, cell_val, col_format)

                                if total_audit_sheets > 0:
                                     audit_progress.progress(audit_sheets_processed / total_audit_sheets)

                        # --- End sheet loop for this file's audit excel ---
                        audit_excel_filename = f"{orig_filename.rsplit('.', 1)[0]}_audit.xlsx"
                        if audit_file_has_content:
                            audit_zip.writestr(audit_excel_filename, audit_excel_buffer.getvalue())
                            files_audited += 1
                        else:
                             logging.info(f"Skipped adding empty/errored audit file to ZIP: {audit_excel_filename}")

                    except Exception as sheet_audit_error:
                         st.error(f"❌ Error creating audit Excel buffer for '{orig_filename}': {sheet_audit_error}")
                         error_logs.append(f"Audit File Creation Error ({orig_filename}): {sheet_audit_error}")
                         logging.error(f"Error creating audit Excel buffer for '{orig_filename}': {sheet_audit_error}", exc_info=True)
                    finally:
                        audit_excel_buffer.close()
            # --- End Audit Zip generation ---
            audit_zip_generated = (files_audited > 0)
            audit_progress.progress(1.0)
            if audit_zip_generated: st.success(f"✅ Audit ZIP generated with {files_audited} file(s).")
            else: st.warning("⚠️ No valid audit files were generated.")

        except Exception as zip_error:
            st.error(f"❌ Error creating Audit ZIP file: {zip_error}")
            logging.error(f"Error creating Audit ZIP: {zip_error}", exc_info=True)

    # ============================
    # Prepare Download Data (using Session State)
    # ============================
    status_text.text("Preparing download...")
    if audit_mode and master_report_generated and audit_zip_generated:
        # Combine Master and Audit into a final ZIP
        try:
            with zipfile.ZipFile(final_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as final_zip:
                # Add Master Report
                master_output_buffer.seek(0)
                final_zip.writestr(output_filename, master_output_buffer.read())
                # Add Audit ZIP
                audit_zip_buffer.seek(0)
                final_zip.writestr(audit_zip_filename_internal, audit_zip_buffer.read())

            st.session_state.download_data = final_zip_buffer.getvalue()
            st.session_state.download_filename = combined_zip_filename
            st.session_state.download_mime = "application/zip"
            st.session_state.download_ready = True
            logging.info(f"Prepared combined ZIP for download: {combined_zip_filename}")

        except Exception as e:
            st.error(f"❌ Failed to create combined ZIP: {e}")
            logging.error(f"Failed to create combined ZIP: {e}", exc_info=True)
            st.session_state.download_ready = False

    elif master_report_generated:
        # Only download Master Report
        st.session_state.download_data = master_output_buffer.getvalue()
        st.session_state.download_filename = output_filename
        st.session_state.download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        st.session_state.download_ready = True
        logging.info(f"Prepared Master Report for download: {output_filename}")
    else:
        st.error("❌ Processing complete, but no files could be generated for download.")
        st.session_state.download_ready = False

    # Clean up buffers
    master_output_buffer.close()
    audit_zip_buffer.close()
    final_zip_buffer.close()
    status_text.text("Processing complete. Download ready below.")


# --- Display Download Button ---
if st.session_state.download_ready and st.session_state.download_data:
    st.markdown("---")
    st.subheader("⬇️ Download Results")
    st.download_button(
        label="📥 Download Processed File(s)",
        data=st.session_state.download_data,
        file_name=st.session_state.download_filename,
        mime=st.session_state.download_mime,
        key="final_download_button"
    )
    st.caption(f"Filename: `{st.session_state.download_filename}`")
else:
    # Show initial message or message if processing failed to produce downloads
    if not uploaded_files:
         st.info("☝️ Configure settings, upload Excel file(s), and click 'Process Uploaded Files'.")
    # (Error messages handled during processing if downloads fail)


# --- Footer ---
st.markdown("---")
st.caption("MaxCode | Smart Excel Grabber")

