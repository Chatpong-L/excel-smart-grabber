
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
    if not letter: return -1
    for char in letter:
        if 'A' <= char <= 'Z': result = result * 26 + (ord(char) - ord('A') + 1)
        else: logging.warning(f"Invalid char '{char}' in letter '{letter}'"); return -1
    return result - 1

# --- Functions needed specifically for the Colab-style Audit ---
def find_and_rename_col_inplace(df, target_label, max_rows=10):
    """ Replicates the Colab example's in-place renaming logic. """
    if not target_label: return None
    target_lower = target_label.lower()
    # Check current columns
    direct_candidates = [c for c in df.columns if target_lower in str(c).lower()]
    if direct_candidates: return direct_candidates[0]
    # Scan rows
    row_limit = min(max_rows, len(df))
    for row_idx in range(row_limit):
        try:
            row_series = df.iloc[row_idx].astype(str)
            if any(target_lower in cell.lower() for cell in row_series):
                new_cols = row_series.tolist()
                # Clean potential duplicate/NaN columns before assigning
                cleaned_cols = []
                counts = {}
                for idx, h in enumerate(new_cols):
                    col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
                    if not col_name: col_name = f"Unnamed: {idx}"
                    if col_name in counts: counts[col_name] += 1; cleaned_cols.append(f"{col_name}.{counts[col_name]}")
                    else: counts[col_name] = 0; cleaned_cols.append(col_name)

                df.columns = cleaned_cols
                df.drop(index=df.index[row_idx], inplace=True) # Drop the header row by its original index position
                df.reset_index(drop=True, inplace=True)
                # Re-check cleaned columns
                direct_candidates = [c for c in df.columns if target_lower in str(c).lower()]
                return direct_candidates[0] if direct_candidates else None
        except Exception as e:
            logging.warning(f"Error processing row {row_idx} during inplace rename: {e}")
            continue # Skip problematic rows
    return None

def find_and_rename_multiple_inplace(df, labels_list, max_rows=10):
    """ Helper to run inplace rename for multiple labels. """
    found_cols = {}
    for lbl in labels_list:
        # Pass max_rows=0 after first successful rename to avoid re-scanning
        scan_rows = max_rows if not found_cols else 0
        col_name = find_and_rename_col_inplace(df, lbl, scan_rows)
        found_cols[lbl] = col_name
    return found_cols

def row_matches_phrases(row_series, phrases):
    """ Replicates the Colab example's row matching logic. """
    row_strs = row_series.astype(str)
    phrases_lower = [p.lower() for p in phrases if p]
    for ph_lower in phrases_lower:
        # Check if phrase exists in any cell of the row
        if row_strs.str.lower().str.contains(ph_lower, na=False).any():
            return True
    return False
# --- End Audit-specific helper functions ---


# --- Main Data Extraction Helper (Non-inplace) ---
# (Keeping this slightly different approach for master data extraction robustness)
def find_header_and_columns_for_master(df_raw, required_labels, optional_labels, max_scan_rows):
    """ Finds header and maps columns without modifying df_raw. """
    if df_raw.empty: return None, {}, None
    num_rows_to_scan = min(max_scan_rows, len(df_raw))
    all_labels_to_find = [lbl for lbl in (required_labels + optional_labels) if lbl]

    if not all_labels_to_find: # Assume first row is header if no labels provided
         header_row_index = 0 if len(df_raw) > 0 else None
         if header_row_index is None: return None, {}, []
         header_values = df_raw.iloc[header_row_index].tolist()
         cleaned_columns = []; counts = {}
         for idx, h in enumerate(header_values):
             col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
             if not col_name: col_name = f"Unnamed: {idx}"
             if col_name in counts: counts[col_name] += 1; cleaned_columns.append(f"{col_name}.{counts[col_name]}")
             else: counts[col_name] = 0; cleaned_columns.append(col_name)
         return header_row_index, {}, cleaned_columns

    for i in range(num_rows_to_scan):
        try:
            row_values_series = df_raw.iloc[i]
            row_values_lower = [str(v).lower() if pd.notna(v) else '' for v in row_values_series]
        except Exception: continue

        required_labels_lower = [req.lower() for req in required_labels if req]
        found_required_count = sum(1 for req_label_lower in required_labels_lower if any(req_label_lower in cell_lower for cell_lower in row_values_lower))

        if found_required_count == len(required_labels_lower):
            header_row_index = i
            header_values = df_raw.iloc[header_row_index].tolist()
            cleaned_columns = []; counts = {}
            for idx, h in enumerate(header_values):
                col_name = str(h).strip() if pd.notna(h) else f"Unnamed: {idx}"
                if not col_name: col_name = f"Unnamed: {idx}"
                if col_name in counts: counts[col_name] += 1; cleaned_columns.append(f"{col_name}.{counts[col_name]}")
                else: counts[col_name] = 0; cleaned_columns.append(col_name)

            column_mapping = {}
            header_values_lower_for_search = [str(v).lower() if pd.notna(v) else '' for v in header_values]
            for user_label in all_labels_to_find:
                user_label_lower = user_label.lower(); found_match = None
                for idx, header_cell_lower in enumerate(header_values_lower_for_search):
                    if user_label_lower in header_cell_lower: found_match = cleaned_columns[idx]; break
                column_mapping[user_label] = found_match
            return header_row_index, column_mapping, cleaned_columns
    return None, {}, None


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
audit_zip_filename_internal = "Audit_Files.zip"
combined_zip_filename = st.sidebar.text_input("Combined Download Filename (if Audit)", "Processing_Results.zip", key="comb_fname")

extra_cols_list = [col.strip() for col in extra_cols_raw.strip().splitlines() if col.strip()]
remove_phrases_list = [phrase.strip() for phrase in remove_phrases_raw.strip().splitlines() if phrase.strip()]

uploaded_files = st.file_uploader(
    "📂 Upload Excel Files (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="uploader"
)

# Initialize session state
if 'download_data' not in st.session_state: st.session_state.download_data = None
if 'download_filename' not in st.session_state: st.session_state.download_filename = None
if 'download_mime' not in st.session_state: st.session_state.download_mime = None
if 'download_ready' not in st.session_state: st.session_state.download_ready = False

# --- Main Execution Logic ---
if st.button("▶️ Process Uploaded Files", key="run_button") and uploaded_files:
    st.session_state.download_ready = False; st.session_state.download_data = None
    start_time = datetime.datetime.now()
    st.info(f"Processing started at {start_time.strftime('%Y-%m-%d %H:%M:%S')}...")

    master_data, skipped_sheets_info, not_typical_col_info, deleted_rows_data, error_logs = [], [], [], [], []
    audit_data_structure = {} # Stores data needed for Colab-style audit generation

    master_output_buffer = io.BytesIO()
    audit_zip_buffer = io.BytesIO()
    final_zip_buffer = io.BytesIO()

    # --- Input Validation ---
    if not value_label: st.error("❗ 'Main: Column Name to Grab' cannot be empty."); st.stop()
    typical_value_index = col_letter_to_index(typical_letter)
    if typical_value_index == -1 and typical_letter: st.warning(f"Invalid typical column letter '{typical_letter}'. Typical check disabled.")
    value_col_not_letter_sheet_name = f"ValueColNot{typical_letter.upper()}" if typical_value_index != -1 else "ValueColTypicalCheck"

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
            progress_bar.progress(files_processed / total_files)
            continue

        # --- Sheet Processing Loop ---
        for sheet_name in sheet_names:
            logging.info(f"-- Processing sheet: {sheet_name}")
            try:
                # --- 1. Data Extraction for Master Report (using non-inplace method) ---
                df_raw_master = xls.parse(sheet_name, header=None, dtype=object)
                df_raw_master.index.name = 'OriginalExcelRow'
                df_raw_master.reset_index(inplace=True)

                required_master = [value_label]
                optional_master = ([trans_label] if trans_label else []) + extra_cols_list
                header_row_idx_master, column_map_master, cleaned_header_master = find_header_and_columns_for_master(
                    df_raw_master, required_master, optional_master, max_scan
                )
                actual_value_col_master = column_map_master.get(value_label)

                # Skip sheet if essential column not found for master data
                if header_row_idx_master is None or not actual_value_col_master:
                    reason = f"Required column '{value_label}' not found for master data extraction"
                    st.warning(f"⚠️ Skipping sheet '{sheet_name}' in '{file_name}': {reason}")
                    skipped_sheets_info.append({"File": file_name, "Sheet": sheet_name, "Reason": reason})
                    # Still attempt to process for audit below if audit_mode is on
                else:
                    # Extract master data (only if header/value col found)
                    deleted_excel_rows_for_master = set() # Need to calculate this again based on master header row
                    if remove_phrases_list:
                         data_start_idx = header_row_idx_master + 1
                         temp_df_for_deletion_check = xls.parse(sheet_name, header=header_row_idx_master, dtype=object) # Read with correct header
                         temp_df_for_deletion_check.index.name = 'OriginalExcelRowAfterHeader'
                         temp_df_for_deletion_check.reset_index(inplace=True)

                         row_removed_mask_master = temp_df_for_deletion_check.apply(lambda row: row_matches_phrases(row, remove_phrases_list), axis=1)
                         # Map index from temp_df back to original excel row (header_row_idx_master + 1 + index)
                         deleted_indices_in_temp = temp_df_for_deletion_check[row_removed_mask_master].index
                         deleted_excel_rows_for_master = {header_row_idx_master + 1 + idx for idx in deleted_indices_in_temp}


                    data_rows_start_idx_master = header_row_idx_master + 1
                    actual_trans_col_master = column_map_master.get(trans_label)
                    extra_col_map_master = {lbl: column_map_master.get(lbl) for lbl in extra_cols_list}

                    for df_idx in range(data_rows_start_idx_master, len(df_raw_master)):
                        original_excel_row = df_raw_master.loc[df_idx, 'OriginalExcelRow']
                        if original_excel_row not in deleted_excel_rows_for_master: # Use master's deleted set
                            row_data_master = df_raw_master.iloc[df_idx]
                            entry = {"SourceFile": file_name, "SourceSheet": sheet_name, "OriginalExcelRow": original_excel_row}
                            try:
                                val_idx = cleaned_header_master.index(actual_value_col_master) if actual_value_col_master and cleaned_header_master else -1
                                entry[value_label] = row_data_master.iloc[val_idx] if val_idx != -1 else None
                                trans_idx = cleaned_header_master.index(actual_trans_col_master) if actual_trans_col_master and cleaned_header_master else -1
                                entry[trans_label] = row_data_master.iloc[trans_idx] if trans_idx != -1 else None
                                for label, actual_col in extra_col_map_master.items():
                                    ex_idx = cleaned_header_master.index(actual_col) if actual_col and cleaned_header_master else -1
                                    entry[label] = row_data_master.iloc[ex_idx] if ex_idx != -1 else None
                                master_data.append(entry)
                            except (ValueError, IndexError) as e: error_logs.append(f"Master Data Lookup Error ({file_name}/{sheet_name}/Row {original_excel_row}): {e}")


                # --- 2. Data Preparation for Audit (Colab-style) ---
                if audit_mode:
                    try:
                        # Read sheet again, this time for in-place modification simulation
                        df_audit_processed = xls.parse(sheet_name, dtype=object) # Read normally first
                        df_audit_processed.index.name = 'OriginalExcelRow' # Keep track of original excel row index
                        df_audit_processed.reset_index(inplace=True) # Make it a column

                        # Apply inplace renaming logic *to this dataframe*
                        val_col_audit = find_and_rename_col_inplace(df_audit_processed, value_label, max_scan)

                        if not val_col_audit:
                             # If value col not found even for audit, store minimal info
                             audit_data_structure[file_name][sheet_name] = {'audit_df': df_audit_processed.copy(), 'highlight_info': {'error': 'Value column not found for audit'}}
                             continue # Skip rest of audit prep for this sheet

                        # Check typical mismatch based on this processed df
                        audit_found_col_index = -1
                        try: audit_found_col_index = df_audit_processed.columns.get_loc(val_col_audit)
                        except KeyError: pass # Ignore if column disappeared
                        audit_typical_mismatch = (typical_value_index != -1 and audit_found_col_index != typical_value_index)

                        # Find other columns on this potentially modified df
                        trans_col_audit = find_and_rename_col_inplace(df_audit_processed, trans_label, 0) # Scan 0 rows now
                        found_extras_audit = find_and_rename_multiple_inplace(df_audit_processed, extra_cols_list, 0)

                        # Identify and split deleted rows based on *this* processed df
                        row_removed_mask_audit = df_audit_processed.apply(lambda row: row_matches_phrases(row, remove_phrases_list), axis=1)
                        df_deleted_audit = df_audit_processed[row_removed_mask_audit].copy()
                        df_kept_audit = df_audit_processed[~row_removed_mask_audit].copy()

                        # Recombine exactly like the Colab example
                        # Align columns before concat if necessary (though inplace rename should keep them same)
                        common_cols = df_kept_audit.columns.union(df_deleted_audit.columns)
                        df_kept_audit = df_kept_audit.reindex(columns=common_cols)
                        df_deleted_audit = df_deleted_audit.reindex(columns=common_cols)

                        full_df_for_audit = pd.concat([df_kept_audit, df_deleted_audit], ignore_index=False) # Use original index from df_audit_processed
                        full_df_for_audit["__deleted__"] = False
                        full_df_for_audit.loc[df_deleted_audit.index, "__deleted__"] = True
                        full_df_for_audit.sort_index(inplace=True) # Sort by original index from df_audit_processed

                        # Prepare highlight info dictionary for this sheet
                        highlight_info = {
                            "deleted_mask": full_df_for_audit["__deleted__"].tolist(), # Mask corresponding to full_df_for_audit rows
                            "value_col": val_col_audit, # Actual name found in audit df
                            "trans_col": trans_col_audit, # Actual name found in audit df
                            "extra_cols": [c for c in found_extras_audit.values() if c], # Actual names found
                            "typical_mismatch": audit_typical_mismatch
                        }
                        # Store the final DataFrame to be written and the highlight info
                        audit_data_structure[file_name][sheet_name] = {
                            'audit_df': full_df_for_audit, # This df has the right columns and __deleted__ flag
                            'highlight_info': highlight_info
                        }
                         # Also collect deleted rows for the master report's "Deleted" sheet using this audit df's info
                        if not df_deleted_audit.empty:
                             audit_deleted_rows = df_deleted_audit.to_dict('records')
                             for rec in audit_deleted_rows:
                                 rec["FileName"] = file_name
                                 rec["SheetName"] = sheet_name
                                 # Ensure OriginalExcelRow is present if it was in df_audit_processed
                                 if 'OriginalExcelRow' not in rec: rec['OriginalExcelRow'] = 'N/A'
                                 deleted_rows_data.append(rec)

                    except Exception as audit_prep_error:
                         st.error(f"❌ Error preparing audit data for sheet '{sheet_name}' in '{file_name}'. Error: {audit_prep_error}")
                         error_logs.append(f"Audit Data Prep Error ({file_name}/{sheet_name}): {audit_prep_error}")
                         audit_data_structure[file_name][sheet_name] = {'audit_df': pd.DataFrame(), 'highlight_info': {'error': str(audit_prep_error)}}


            except Exception as e:
                st.error(f"❌ Error processing sheet '{sheet_name}' in file '{file_name}'. Skipping sheet. Error: {e}")
                skipped_sheets_info.append({"File": file_name, "Sheet": sheet_name, "Reason": f"Processing error: {e}"})
                error_logs.append(f"Sheet Processing Error ({file_name}/{sheet_name}): {e}")
                if audit_mode: # Ensure audit structure exists even on error
                    audit_data_structure[file_name][sheet_name] = {'audit_df': pd.DataFrame(), 'highlight_info': {'error': str(e)}}

        # --- End Sheet Loop ---
        progress_bar.progress(files_processed / total_files)
    # --- End File Loop ---

    status_text.success(f"✅ File processing completed in {datetime.datetime.now() - start_time}.")

    # ============================
    # Post-Processing (Combine DFs, Analysis)
    # ============================
    master_df = pd.DataFrame(master_data) if master_data else pd.DataFrame()
    skipped_df = pd.DataFrame(skipped_sheets_info) if skipped_sheets_info else pd.DataFrame()
    not_typical_df = pd.DataFrame(not_typical_col_info) if not_typical_col_info else pd.DataFrame()
    # Use deleted rows collected during audit prep if audit_mode, otherwise it remains empty
    deleted_df = pd.DataFrame(deleted_rows_data) if deleted_rows_data else pd.DataFrame()
    errors_df = pd.DataFrame(error_logs, columns=["ErrorLog"]) if error_logs else pd.DataFrame()

    analysis_df = pd.DataFrame()
    if not master_df.empty and value_label and value_label in master_df.columns:
        numeric_col_name = f"{value_label}_numeric"
        master_df[numeric_col_name] = pd.to_numeric(master_df[value_label].astype(str).str.replace(',', '', regex=False), errors='coerce')
        if master_df[numeric_col_name].notna().any(): analysis_df = pd.DataFrame(master_df[numeric_col_name].describe())

    # ============================
    # Write Master Report to Buffer
    # ============================
    master_report_generated = False
    try:
        with pd.ExcelWriter(master_output_buffer, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors": True}}) as writer:
            if not master_df.empty:
                 cols_order = ["SourceFile", "SourceSheet", "OriginalExcelRow"]
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
                # Ensure OriginalExcelRow is a key column
                deleted_cols = ['FileName', 'SheetName', 'OriginalExcelRow'] + [col for col in deleted_df.columns if col not in ['FileName', 'SheetName', 'OriginalExcelRow', '__deleted__']] # Exclude __deleted__
                deleted_df.to_excel(writer, sheet_name='DeletedRows', index=False, columns=deleted_cols)
            if not analysis_df.empty: analysis_df.to_excel(writer, sheet_name='ValueColumnAnalysis')
            if not errors_df.empty: errors_df.to_excel(writer, sheet_name='ProcessingErrors', index=False)
        master_report_generated = True
    except Exception as e:
        st.error(f"❌ Error creating Master Report Excel buffer: {e}")


    # ============================
    # Generate Audit Files + ZIP (if enabled) - USING COLAB LOGIC
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
                        with pd.ExcelWriter(audit_excel_buffer, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
                            workbook = writer.book
                            red_fill = workbook.add_format({"bg_color": "#FFC7CE"})
                            yellow_fill = workbook.add_format({"bg_color": "#FFEB9C"})
                            purple_fill = workbook.add_format({"bg_color": "#E4DFEC"})
                            # No specific header format needed if applying later

                            for sheet_name, audit_info in sheets_data.items():
                                audit_sheets_processed += 1
                                df_sheet = audit_info.get('audit_df')
                                hi = audit_info.get('highlight_info', {})

                                if df_sheet is None or df_sheet.empty:
                                    # Handle sheets that failed audit prep
                                    ws = workbook.add_worksheet(sheet_name[:31])
                                    ws.write(0, 0, f"Error during audit preparation: {hi.get('error', 'Unknown error')}")
                                    audit_file_has_content = True # Include error sheet
                                    continue

                                audit_file_has_content = True
                                # Prepare df for writing (drop the helper deleted flag)
                                df_to_write = df_sheet.drop(columns=['__deleted__'], errors='ignore')
                                # Write data using pandas to_excel first
                                df_to_write.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                                ws = writer.sheets[sheet_name[:31]]

                                # Get highlight info
                                deleted_mask_list = hi.get("deleted_mask", [])
                                val_col = hi.get("value_col")
                                trans_col = hi.get("trans_col")
                                extra_cols = hi.get("extra_cols", [])
                                mismatch = hi.get("typical_mismatch", False)

                                # Replicate Colab Highlighting Logic Exactly
                                # 1) Highlight deleted rows red
                                if deleted_mask_list:
                                     for r_idx, is_del in enumerate(deleted_mask_list):
                                         if is_del:
                                             # Apply format to the entire row (Excel row index r_idx + 1)
                                             # ws.set_row(row_num, height, cell_format)
                                             ws.set_row(r_idx + 1, None, red_fill) # +1 because Excel rows are 1-based in xlsxwriter

                                # 2) Highlight used columns yellow (if not deleted)
                                used_cols = []
                                if val_col and (val_col in df_to_write.columns): used_cols.append(val_col)
                                if trans_col and (trans_col in df_to_write.columns): used_cols.append(trans_col)
                                for ec in extra_cols:
                                    if ec and (ec in df_to_write.columns): used_cols.append(ec)

                                for col_name in used_cols:
                                    try:
                                        col_idx = df_to_write.columns.get_loc(col_name)
                                        for r_idx in range(len(df_to_write)):
                                            # Check mask length; skip if inconsistent
                                            if r_idx < len(deleted_mask_list) and not deleted_mask_list[r_idx]:
                                                # Get existing value and rewrite with format
                                                cell_value = df_to_write.iloc[r_idx, col_idx]
                                                # Use direct write here as per Colab example (assumes safe types in df_to_write)
                                                ws.write(r_idx + 1, col_idx, cell_value, yellow_fill)
                                    except KeyError:
                                        logging.warning(f"Audit highlight: Column '{col_name}' not found in written data for {orig_filename}/{sheet_name}.")
                                    except IndexError:
                                         logging.warning(f"Audit highlight: Index mismatch for deleted_mask check in {orig_filename}/{sheet_name}.")


                                # 3) Highlight mismatched value column purple (if not deleted)
                                if mismatch and val_col and (val_col in df_to_write.columns):
                                     try:
                                         col_idx = df_to_write.columns.get_loc(val_col)
                                         for r_idx in range(len(df_to_write)):
                                             # Check mask length; skip if inconsistent
                                             if r_idx < len(deleted_mask_list) and not deleted_mask_list[r_idx]:
                                                 # Get existing value and rewrite with format
                                                 cell_value = df_to_write.iloc[r_idx, col_idx]
                                                 # Use direct write here as per Colab example
                                                 ws.write(r_idx + 1, col_idx, cell_value, purple_fill) # Purple overrides yellow
                                     except KeyError:
                                         logging.warning(f"Audit highlight (purple): Column '{val_col}' not found in written data for {orig_filename}/{sheet_name}.")
                                     except IndexError:
                                         logging.warning(f"Audit highlight (purple): Index mismatch for deleted_mask check in {orig_filename}/{sheet_name}.")


                                if total_audit_sheets > 0: audit_progress.progress(audit_sheets_processed / total_audit_sheets)

                        # --- End sheet loop for this file's audit excel ---
                        audit_excel_filename = f"{orig_filename.rsplit('.', 1)[0]}_audit.xlsx"
                        if audit_file_has_content:
                            audit_zip.writestr(audit_excel_filename, audit_excel_buffer.getvalue())
                            files_audited += 1
                    except Exception as sheet_audit_error:
                         st.error(f"❌ Error creating audit Excel buffer for '{orig_filename}': {sheet_audit_error}")
                         error_logs.append(f"Audit File Creation Error ({orig_filename}): {sheet_audit_error}")
                    finally:
                        audit_excel_buffer.close()
            # --- End Audit Zip generation ---
            audit_zip_generated = (files_audited > 0)
            audit_progress.progress(1.0)
            if audit_zip_generated: st.success(f"✅ Audit ZIP generated with {files_audited} file(s).")
            else: st.warning("⚠️ No valid audit files were generated.")
        except Exception as zip_error: st.error(f"❌ Error creating Audit ZIP file: {zip_error}")

    # ============================
    # Prepare Download Data
    # ============================
    status_text.text("Preparing download...")
    if audit_mode and master_report_generated and audit_zip_generated:
        try:
            with zipfile.ZipFile(final_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as final_zip:
                master_output_buffer.seek(0); final_zip.writestr(output_filename, master_output_buffer.read())
                audit_zip_buffer.seek(0); final_zip.writestr(audit_zip_filename_internal, audit_zip_buffer.read())
            st.session_state.download_data = final_zip_buffer.getvalue()
            st.session_state.download_filename = combined_zip_filename
            st.session_state.download_mime = "application/zip"
            st.session_state.download_ready = True
        except Exception as e: st.error(f"❌ Failed to create combined ZIP: {e}"); st.session_state.download_ready = False
    elif master_report_generated:
        st.session_state.download_data = master_output_buffer.getvalue()
        st.session_state.download_filename = output_filename
        st.session_state.download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        st.session_state.download_ready = True
    else: st.error("❌ Processing complete, but no files generated for download."); st.session_state.download_ready = False

    master_output_buffer.close(); audit_zip_buffer.close(); final_zip_buffer.close()
    status_text.text("Processing complete. Download ready below.")

# --- Display Download Button ---
if st.session_state.download_ready and st.session_state.download_data:
    st.markdown("---"); st.subheader("⬇️ Download Results")
    st.download_button(
        label="📥 Download Processed File(s)", data=st.session_state.download_data,
        file_name=st.session_state.download_filename, mime=st.session_state.download_mime,
        key="final_download_button"
    )
    st.caption(f"Filename: `{st.session_state.download_filename}`")
elif not uploaded_files: st.info("☝️ Configure settings, upload Excel file(s), and click 'Process Uploaded Files'.")

st.markdown("---"); st.caption("MaxCode | Smart Excel Grabber")
