import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime

st.set_page_config(page_title="Excel Smart Grabber", layout="wide")

st.title("📊 Excel Smart Grabber 3000 (Audit-Ready Version)")
st.markdown("Upload Excel files and extract specified columns from all sheets. Deleted rows and mismatched columns will be flagged and auditable.")

# ==== Sidebar Inputs ====
value_label = st.sidebar.text_input("Value Column Label", "รวมเงิน")
trans_label = st.sidebar.text_input("Transaction Column Label", "รายการ")
typical_letter = st.sidebar.text_input("Expected Column Letter (e.g. M)", "M")
extra_cols = st.sidebar.text_area("Extra Columns (one per line)", "")
remove_phrases = st.sidebar.text_area("Remove Row Phrases (one per line)", "TOTAL")
max_scan = st.sidebar.number_input("Header Scan Limit", 1, 30, 10)
audit_mode = st.sidebar.checkbox("Generate Audit File?", value=True)
output_filename = st.sidebar.text_input("Output Excel Name", "Master_Report.xlsx")

uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

# ===== Helper =====
def col_letter_to_index(letter):
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1

def find_column(df, label, max_rows):
    candidates = [c for c in df.columns if label.lower() in str(c).lower()]
    if candidates:
        return candidates[0]
    for i in range(min(max_rows, len(df))):
        row = df.iloc[i].astype(str)
        if any(label.lower() in cell.lower() for cell in row):
            df.columns = row
            df.drop(index=i, inplace=True)
            df.reset_index(drop=True, inplace=True)
            return find_column(df, label, 0)
    return None

# ===== Main Logic =====
if st.button("▶️ Run Excel Grabber") and uploaded_files:
    audit_zip = io.BytesIO()
    with zipfile.ZipFile(audit_zip, "w") as audit_bundle:
        master = []
        deleted_rows = []
        not_typical = []
        skipped = []
        for up in uploaded_files:
            xls = pd.ExcelFile(up)
            audit_writer = pd.ExcelWriter(f"{up.name}_audit.xlsx", engine="xlsxwriter")
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                df.reset_index(drop=True, inplace=True)
                raw = df.copy()
                val_col = find_column(df, value_label, max_scan)
                if not val_col:
                    skipped.append((up.name, sheet))
                    continue
                trans_col = find_column(df, trans_label, max_scan)
                extras = {col: find_column(df, col, max_scan) for col in extra_cols.strip().splitlines() if col.strip()}
                row_mask = df.apply(lambda r: any(p.lower() in str(r).lower() for p in remove_phrases.strip().splitlines()), axis=1)
                deleted = df[row_mask].copy()
                df = df[~row_mask]
                for i, r in df.iterrows():
                    entry = {
                        "FileName": up.name,
                        "SheetName": sheet,
                        value_label: r.get(val_col),
                        trans_label: r.get(trans_col) if trans_col else None
                    }
                    for label, col in extras.items():
                        entry[label] = r.get(col) if col else None
                    master.append(entry)
                for i, r in deleted.iterrows():
                    rec = r.to_dict()
                    rec["FileName"] = up.name
                    rec["SheetName"] = sheet
                    deleted_rows.append(rec)
                val_idx = df.columns.get_loc(val_col)
                if val_idx != col_letter_to_index(typical_letter):
                    not_typical.append((up.name, sheet))

                # === Audit Highlight ===
                raw["__deleted__"] = raw.index.isin(deleted.index)
                ws = audit_writer.book.add_worksheet(sheet)
                for col_idx, col in enumerate(raw.columns):
                    ws.write(0, col_idx, col)
                    for row_idx in range(len(raw)):
                        cell = raw.iloc[row_idx, col_idx]
                        fmt = None
                        if raw.iloc[row_idx]["__deleted__"]:
                            fmt = audit_writer.book.add_format({"bg_color": "#FFC7CE"})  # red
                        elif col == val_col or col == trans_col or col in extras.values():
                            fmt = audit_writer.book.add_format({"bg_color": "#FFEB9C"})  # yellow
                        if fmt:
                            ws.write(row_idx+1, col_idx, cell, fmt)
                        else:
                            ws.write(row_idx+1, col_idx, cell)
                raw.drop(columns=["__deleted__"], inplace=True)
            audit_writer.close()
            audit_bundle.write(f"{up.name}_audit.xlsx", open(f"{up.name}_audit.xlsx", "rb").read())
            os.remove(f"{up.name}_audit.xlsx")

    # Export Master Report
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", engine_kwargs={"options":{"nan_inf_to_errors": True}}) as writer:
        pd.DataFrame(master).to_excel(writer, sheet_name="AllData", index=False)
        pd.DataFrame(deleted_rows).to_excel(writer, sheet_name="DeletedRows", index=False)
        pd.DataFrame(not_typical, columns=["File", "Sheet"]).to_excel(writer, sheet_name="NotTypical", index=False)
        pd.DataFrame(skipped, columns=["File", "Sheet"]).to_excel(writer, sheet_name="SkippedSheets", index=False)

    st.success("✅ Processing complete!")
    st.download_button("📥 Download Master Excel", out.getvalue(), output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if audit_mode:
        st.download_button("📥 Download Audit ZIP", audit_zip.getvalue(), "audit_files.zip", mime="application/zip")
else:
    st.info("Upload file(s) and click ▶️ Run to start.")