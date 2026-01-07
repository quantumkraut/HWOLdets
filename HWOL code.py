import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

# --- UI ---
st.set_page_config(page_title="HWOL OEMA")

st.title("HWOL OEMA")

# File uploader
labfile = st.file_uploader("Upload OEMA lab Excel File (.xlsx):", type=["xlsx"])

# Column number input (defaults to 1)
column_no = st.number_input(
    "Required sample column number (only change if more than one sample):",
    min_value=1, value=1, step=1
)

# Status box
status_placeholder = st.empty()

# Process button
process = st.button("Click to Generate OEMA HWOL")

# Reactive-like variable
processed_file_path = None

# --- Processing ---
if process and labfile is not None:
    try:
        status_placeholder.text("Processing file... please wait.")

        # Load template Excel
        template_path = "MultiSamplesOEMA.xlsx"
        if not os.path.exists(template_path):
            status_placeholder.text("❌ Error: Template file not found.")
            st.stop()

        # Load lab Excel sheet 3 (equivalent to R)
        test = pd.read_excel(labfile, sheet_name=2)  # 0-indexed, sheet 3 = 2

        # Safe row/column slicing
        row_start = min(13, test.shape[0]-1)  # R index 14 = pandas 13
        col_index = min(5 + column_no, test.shape[1]-1)  # R index 6 + input = pandas 5 + input

        # Extract subset safely
        test = test.iloc[row_start:, [0, col_index]].copy()
        test.columns = ["Determinant", "Result"]

        # --- ROUND DECIMAL PLACES ---
        def round_if_numeric(x):
            try:
                num = float(x)
                # Round if more than 3 decimals
                if '.' in str(x) and len(str(x).split('.')[1]) >= 4:
                    return f"{round(num, 2):.2f}"
                else:
                    return x
            except:
                return x

        test["Result"] = test["Result"].apply(round_if_numeric)

        # --- SAFE GREP FUNCTION ---
    def safe_grep(pattern, series):
        matches = series[series.str.contains(pattern, na=False, regex=False)]
        if len(matches) == 0:
            return ""
        else:
            return matches.iloc[0]


        # --- Build ordered Determinants (same as R test2) ---
        determinants = [
            "", "",
            "Antimony", "Arsenic", "Barium", "Beryllium", "Boron", "Cadmium",
            "Chromium (III)", "Hexavalent", "Copper", "Cobalt", "Lead", "Manganese",
            "Mercury", "Molybdenum", "Nickel", "Selenium", "Thallium", "Tin",
            "Calcium", "Phophorous", "Vanadium", "Sulphur", "Zinc", "Iron",
            "Total TPH C10-C40", "",
            "MTBE", "Benzene", "Toulene", "Ethylbenzene", "Total Xylene",
            "Total Cyanide", "pH",
            "Napthalene", "Acenaphthylene", "Acenaphthene", "Fluorene", "Phenanthrene",
            "Anthracene", "Fluoranthene", "Pyrene", "Benzo (a) anthracene",
            "Chrysene", "Benzo (b) fluoranthene", "Benzo (K) fluoranthene",
            "Benzo (a) pyrene", "Indeno (1,2,3", "Dibenz", "perylene"
        ]
        # Fill up to same length as R if needed
        while len(determinants) < 80:
            determinants.append("")

        # Build DataFrame safely
        test2 = pd.DataFrame({
            "Determinant": [safe_grep(d, test["Determinant"]) if d else "" for d in determinants]
        })

        # Join with actual results
        joinedup = pd.merge(test2, test, on="Determinant", how="left")

        # Extract only the results column
        values = joinedup[["Result"]]

        # --- Write to template ---
        wb = load_workbook(template_path)
        ws = wb.active

        for i, val in enumerate(values["Result"], start=3):  # Excel row 3
            ws.cell(row=i, column=3, value=val)

        # Save to temporary file
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(tmp.name)
        processed_file_path = tmp.name

        status_placeholder.text("✅ Done! Download your autofilled HWOL file below.")

        # Provide download button
        st.download_button(
            "DOWNLOAD HWOL",
            data=open(processed_file_path, "rb"),
            file_name=f"OEMA_HWOL_{pd.Timestamp.today().date()}.xlsx"
        )

    except Exception as e:
        status_placeholder.text(f"❌ Error: {e}")


