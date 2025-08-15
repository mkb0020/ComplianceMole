#This is just a practice exercise to work on manipulating data.  
#I have a csv with sample data from a chemical audit that checked for Concentration, pH, Temperature, Pressure, and Flow Rate for 10 random chemicals.
#And I have another csv file with the acceptable ranges for each measurement for each chemical
#The goal is to:
  #1. Drag and drop the sample data csv file to an exe which will ask the user to input some info.
  #2. Compare the sample data to the acceptable ranges and
  #3. Output an xlsx document that shows what is out of compliance, the min and max for each measurement for each chemical, and scores the company based on a weighted average 
  #4. Reformat the xlsx doc with different size fonts, different cell colors, and different borders
  #5. Save the new doc to a location of the users choice

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from tkinter import Tk, simpledialog, filedialog
from datetime import datetime
import os

# -------------------- Utilities --------------------
def _norm(s: str) -> str:
    """
    Normalize a column name for matching:
    - lowercases
    - replace common unit tokens to consistent forms
    - remove most punctuation/spaces -> underscores
    """
    s = (s or "").strip().lower()
    # normalize some units/spellings
    s = s.replace("°c", "celsius")
    s = s.replace("(c)", "celsius")
    s = s.replace("c°", "celsius")
    s = s.replace(" c ", " celsius ")
    s = s.replace("kpa", "kpa")
    s = s.replace("l/min", "l_min")
    s = s.replace("l per min", "l_min")
    s = s.replace("l per minute", "l_min")
    s = s.replace("ph", "ph")  # keep as ph
    # strip punctuation to underscores
    for ch in " -()/\\[].,":
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    return s.strip("_")


# Canonical report headers you want in the Excel output (pretty names)
CANONICAL_REPORT_HEADERS = [
    "SAMPLE ID",
    "CHEMICAL",
    "CONCENTRATION (ppm)",
    "pH LEVEL",
    "TEMPERATURE (Celsius)",
    "PRESSURE (kPa)",
    "FLOW RATE (L_min)",
    "STATUS",
    "COMMENT",
]

# Mapping from normalized tokens -> pretty canonical header names
CSV_HEADER_ALIASES = {
    # SAMPLE ID
    "sample_id": "SAMPLE ID",
    "sampleid": "SAMPLE ID",
    "id": "SAMPLE ID",
    "sample": "SAMPLE ID",

    # CHEMICAL
    "chemical": "CHEMICAL",
    "compound": "CHEMICAL",
    "analyte": "CHEMICAL",
    "reagent": "CHEMICAL",

    # CONCENTRATION (ppm)
    "concentration_ppm": "CONCENTRATION (ppm)",
    "concentration": "CONCENTRATION (ppm)",
    "conc_ppm": "CONCENTRATION (ppm)",
    "conc": "CONCENTRATION (ppm)",
    "concentration_ppm_": "CONCENTRATION (ppm)",

    # pH LEVEL
    "ph_level": "pH LEVEL",
    "ph": "pH LEVEL",

    # TEMPERATURE (Celsius)
    "temperature_celsius": "TEMPERATURE (Celsius)",
    "temperature_c": "TEMPERATURE (Celsius)",
    "temp_c": "TEMPERATURE (Celsius)",
    "temperature": "TEMPERATURE (Celsius)",
    "temp": "TEMPERATURE (Celsius)",

    # PRESSURE (kPa)
    "pressure_kpa": "PRESSURE (kPa)",
    "pressure": "PRESSURE (kPa)",

    # FLOW RATE (L_min)
    "flow_rate_l_min": "FLOW RATE (L_min)",
    "flowrate_l_min": "FLOW RATE (L_min)",
    "flow_rate": "FLOW RATE (L_min)",
    "flowrate": "FLOW RATE (L_min)",
    "flow": "FLOW RATE (L_min)",
}

# Expected normalized columns for the ranges workbook
RANGE_COL_MAP = {
    "chemical": "chemical",  # index col
    "concentration_ppm_min": "Concentration_ppm_Min",
    "concentration_ppm_max": "Concentration_ppm_Max",
    "ph_level_min": "pH_Level_Min",
    "ph_level_max": "pH_Level_Max",
    "temperature_c_min": "Temperature_C_Min",
    "temperature_c_max": "Temperature_C_Max",
    "pressure_kpa_min": "Pressure_kPa_Min",
    "pressure_kpa_max": "Pressure_kPa_Max",
    "flow_rate_l_min_min": "Flow_Rate_L_min_Min",
    "flow_rate_l_min_max": "Flow_Rate_L_min_Max",
}


# -------------------- Phase 1: File + User Info --------------------
def select_file():
    Tk().withdraw()
    return filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )

def load_csv(file_path):
    df = pd.read_csv(file_path)
    last_row = len(df)
    return df, last_row

def get_user_info():
    root = Tk()
    root.withdraw()
    first = simpledialog.askstring("User Info", "Enter First Name:") or ""
    middle = simpledialog.askstring("User Info", "Enter Middle Name (Optional):") or ""
    last = simpledialog.askstring("User Info", "Enter Last Name:") or ""
    company = simpledialog.askstring("User Info", "Enter Company Name:") or ""

    return {
        "FirstName": first,
        "MiddleName": middle,
        "LastName": last,
        "FirstIntl": first[:1],
        "MidIntl": middle[:1],
        "LastIntl": last[:1],
        "CompanyName": company,
        "DateToday": datetime.today().strftime("%Y%m%d"),
    }

def get_save_path(user_info):
    Tk().withdraw()
    initials = f"{user_info['FirstIntl']}{user_info['MidIntl']}{user_info['LastIntl']}"
    base_name = f"Chemical_Compliance_Report_{user_info['CompanyName']}_{user_info['DateToday']}_{initials}.xlsx"
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=base_name)

    if not save_path:
        return ""

    # If chosen path exists, ask for version
    if os.path.exists(save_path):
        version = simpledialog.askstring("Version", "Enter version number:") or "2"
        base_name = f"Chemical_Compliance_Report_{user_info['CompanyName']}_{user_info['DateToday']}_{initials}_v{version}.xlsx"
        save_path = os.path.join(os.path.dirname(save_path), base_name)

    return save_path


# -------------------- Column Standardization --------------------
def standardize_csv_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename CSV columns to your canonical pretty headers where possible.
    Then ensure STATUS and COMMENT exist, and reorder to your preferred order.
    """
    rename_map = {}
    used_targets = set()

    for col in df.columns:
        key = _norm(col)
        target = CSV_HEADER_ALIASES.get(key)
        if target and target not in used_targets:
            rename_map[col] = target
            used_targets.add(target)

    df = df.rename(columns=rename_map)

    # Ensure required data columns exist (you can create empty if missing, but ideally they should be present)
    for required in CANONICAL_REPORT_HEADERS[:-2]:  # all except STATUS/COMMENT
        if required not in df.columns:
            # Create missing columns as empty (NaN)
            df[required] = pd.NA

    # Add STATUS and COMMENT if missing
    if "STATUS" not in df.columns:
        df["STATUS"] = pd.NA
    if "COMMENT" not in df.columns:
        df["COMMENT"] = pd.NA

    # Reorder columns
    ordered_cols = [c for c in CANONICAL_REPORT_HEADERS if c in df.columns]
    # include any extra columns at the end (if any)
    extra_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + extra_cols]

    # Convert numeric columns to numeric (quietly coercing)
    for num_col in [
        "CONCENTRATION (ppm)",
        "pH LEVEL",
        "TEMPERATURE (Celsius)",
        "PRESSURE (kPa)",
        "FLOW RATE (L_min)",
    ]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    return df


# -------------------- Phase 2: Load Ranges + Check --------------------
def load_ranges():
    ranges_path = r"C:\Users\mkb00\PROJECTS\PythonProjects\ComplianceMole\CompliantRanges.xlsx"
    rng = pd.read_excel(ranges_path)

    # Build a normalized->original map for the ranges sheet
    norm_to_original = {}
    for col in rng.columns:
        norm_to_original[_norm(col)] = col

    # Make sure we have a CHEMICAL column to index by
    chem_col = None
    for cand in ("chemical", "chemicals", "compound", "analyte"):
        if cand in norm_to_original:
            chem_col = norm_to_original[cand]
            break
    if not chem_col:
        raise KeyError("Could not find a 'Chemical' column in CompliantRanges.xlsx")

    # Rename ranges columns to a consistent set (using RANGE_COL_MAP)
    col_renames = {}
    for needed_norm, pretty in RANGE_COL_MAP.items():
        if needed_norm == "chemical":
            continue
        # if present under some alias, capture it
        for norm_name, orig_name in norm_to_original.items():
            if norm_name == needed_norm:
                col_renames[orig_name] = pretty

    # If they already match the pretty names, this is a no-op for those
    rng = rng.rename(columns=col_renames)

    # Now verify we have all required columns (pretty names)
    missing = [v for k, v in RANGE_COL_MAP.items() if k != "chemical" and v not in rng.columns]
    if missing:
        raise KeyError(
            "Missing expected columns in CompliantRanges.xlsx: "
            + ", ".join(missing)
        )

    rng = rng.set_index(chem_col)

    return rng


def check_compliance(df: pd.DataFrame, ranges_df: pd.DataFrame) -> pd.DataFrame:
    # Defensive: if required data columns are missing, we still proceed but mark UNKNOWN
    req_cols = [
        "CHEMICAL",
        "CONCENTRATION (ppm)",
        "pH LEVEL",
        "TEMPERATURE (Celsius)",
        "PRESSURE (kPa)",
        "FLOW RATE (L_min)",
    ]
    for c in req_cols:
        if c not in df.columns:
            # If the CSV truly lacks a required column, mark rows unknown
            df["STATUS"] = "UNKNOWN"
            df["COMMENT"] = "Required columns missing in source CSV."
            return df

    # Iterate rows and evaluate
    for idx, row in df.iterrows():
        chem = row["CHEMICAL"]

        # If chemical is NaN or not in the reference
        if pd.isna(chem) or chem not in ranges_df.index:
            df.at[idx, "STATUS"] = "UNKNOWN CHEMICAL"
            df.at[idx, "COMMENT"] = "No compliance data found."
            continue

        limits = ranges_df.loc[chem]

        issues = []

        # Helpers to read limits safely
        cmin = limits["Concentration_ppm_Min"]
        cmax = limits["Concentration_ppm_Max"]
        phmin = limits["pH_Level_Min"]
        phmax = limits["pH_Level_Max"]
        tmin = limits["Temperature_C_Min"]
        tmax = limits["Temperature_C_Max"]
        pmin = limits["Pressure_kPa_Min"]
        pmax = limits["Pressure_kPa_Max"]
        frmin = limits["Flow_Rate_L_min_Min"]
        frmax = limits["Flow_Rate_L_min_Max"]

        # Pull row values
        conc = row["CONCENTRATION (ppm)"]
        ph = row["pH LEVEL"]
        tempc = row["TEMPERATURE (Celsius)"]
        press = row["PRESSURE (kPa)"]
        flow = row["FLOW RATE (L_min)"]

        # Validate each metric (NaN counts as violation)
        def _out_of_range(val, lo, hi):
            try:
                return not (lo <= val <= hi)
            except TypeError:
                return True  # val is NaN or non-numeric

        if _out_of_range(conc, cmin, cmax):
            issues.append(f"Concentration not within acceptable range: {cmin} - {cmax}.")
        if _out_of_range(ph, phmin, phmax):
            issues.append(f"pH not within acceptable range: {phmin} - {phmax}.")
        if _out_of_range(tempc, tmin, tmax):
            issues.append(f"Temperature not within acceptable range: {tmin} - {tmax}.")
        if _out_of_range(press, pmin, pmax):
            issues.append(f"Pressure not within acceptable range: {pmin} - {pmax}.")
        if _out_of_range(flow, frmin, frmax):
            issues.append(f"Flow Rate not within acceptable range: {frmin} - {frmax}.")

        if issues:
            df.at[idx, "STATUS"] = "NON-COMPLIANT"
            df.at[idx, "COMMENT"] = " ".join(issues)
        else:
            df.at[idx, "STATUS"] = "COMPLIANT"
            df.at[idx, "COMMENT"] = "Within Acceptable Ranges"

    return df

# -------------------- Main --------------------
if __name__ == "__main__":
    csv_path = select_file()
    if not csv_path:
        print("No file selected. Exiting.")
        raise SystemExit

    df, last_row = load_csv(csv_path)

    user_info = get_user_info()
    save_path = get_save_path(user_info)
    if not save_path:
        print("No save location chosen. Exiting.")
        raise SystemExit

    # Standardize CSV headers (fixes your 7 vs 9 mismatch)
    df = standardize_csv_headers(df)

    # Load ranges (tolerant to header naming differences)
    ranges_df = load_ranges()

    # Run compliance logic
    df = check_compliance(df, ranges_df)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Save to Excel (no formatting yet)
df.to_excel(save_path, index=False)

# -------------------- Formatting with openpyxl --------------------
wb = load_workbook(save_path)
ws = wb.active
ws.title = "Sample Data"

# === AUTO COLUMN WIDTH ===
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    ws.column_dimensions[col_letter].width = max_length + 2


# Insert a new row at top for "SAMPLE DATA"
ws.insert_rows(1)
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
cell = ws.cell(row=1, column=1)
cell.value = "SAMPLE DATA"
cell.fill = PatternFill(start_color="5C6586", end_color="5C6586", fill_type="solid")
cell.font = Font(name="Aptos Display", bold=True, color="FFFFFF", size=12)
cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 20
# Bottom double border for row 1
double_side = Side(style="double", color="000000")
for col in range(1, 10):
    ws.cell(row=1, column=col).border = Border(bottom=double_side)

# Format header row (now row 2)
header_fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
for col in range(1, 10):
    c = ws.cell(row=2, column=col)
    c.fill = header_fill
    c.font = Font(name="Aptos Narrow", bold=True, color="000000", size=10.5)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(bottom=double_side)
ws.row_dimensions[2].height = 35

# Format data rows (3 to last)
last_data_row = ws.max_row
for r in range(3, last_data_row + 1):
    for col in range(1, 10):
        c = ws.cell(row=r, column=col)
        c.font = Font(name="Aptos Narrow", color="000000", size=10.5)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    ws.row_dimensions[r].height = 15

# Borders: all thin + thick outline
thin = Side(style="thin", color="000000")
thick = Side(style="thick", color="000000")
for row in ws.iter_rows(min_row=1, max_row=last_data_row, min_col=1, max_col=9):
    for cell in row:
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

# Thick outline border for the whole range
for col in range(1, 10):
    ws.cell(row=1, column=col).border = Border(top=thick, bottom=ws.cell(row=1, column=col).border.bottom,
                                               left=ws.cell(row=1, column=col).border.left,
                                               right=ws.cell(row=1, column=col).border.right)
    ws.cell(row=last_data_row, column=col).border = Border(top=ws.cell(row=last_data_row, column=col).border.top,
                                                          bottom=thick,
                                                          left=ws.cell(row=last_data_row, column=col).border.left,
                                                          right=ws.cell(row=last_data_row, column=col).border.right)

for r in range(1, last_data_row + 1):
    ws.cell(row=r, column=1).border = Border(top=ws.cell(row=r, column=1).border.top,
                                            bottom=ws.cell(row=r, column=1).border.bottom,
                                            left=thick,
                                            right=ws.cell(row=r, column=1).border.right)
    ws.cell(row=r, column=9).border = Border(top=ws.cell(row=r, column=9).border.top,
                                            bottom=ws.cell(row=r, column=9).border.bottom,
                                            left=ws.cell(row=r, column=9).border.left,
                                            right=thick)


from openpyxl.styles import Alignment

# Left-align A1 with 1 indent
ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

# Left-align column I with 1 indent
for row in range(1, last_row + 1):
    ws[f"I{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)



# ---------------------------
# CREATE "Summary" SHEET
# ---------------------------
summary_ws = wb.create_sheet("Summary")

# Styles
header_fill = PatternFill(start_color="5C6586", end_color="5C6586", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
center_bold = Alignment(horizontal="center", vertical="center")

# Title
summary_ws.merge_cells('B2:F2')
summary_ws['B2'] = "COMPLIANCE ANALYSIS"
summary_ws['B2'].fill = header_fill
summary_ws['B2'].font = header_font
summary_ws['B2'].alignment = center_bold

# Summary info
summary_ws['B3'] = "Completed By:"
#summary_ws['D3'] = user_info['first']
summary_ws['B4'] = "Date:"
#summary_ws['D4'] = user_info['DateToday']
summary_ws['B5'] = "Company:"
#summary_ws['D5'] = user_info['CompanyName']

# Overall Score label
summary_ws['B7'] = "Weighted Overall Score:"
summary_ws['D7'] = 0  # will update later

# ---------------------------
# CHEMICALS TABLE
# ---------------------------
chemicals = df['CHEMICAL'].unique()
chemicals.sort()
start_row = 11

summary_ws['B10'] = "Chemical"
summary_ws['C10'] = "Total Samples"
summary_ws['E10'] = "Acceptable Samples"
summary_ws['G10'] = "Non-Compliant Samples"
summary_ws['I10'] = "Compliance %"
summary_ws['K10'] = "Priority"

for col in ['B','C','E','G','I','K']:
    summary_ws[f"{col}10"].fill = header_fill
    summary_ws[f"{col}10"].font = header_font
    summary_ws[f"{col}10"].alignment = center_bold

row = start_row
for chem in chemicals:
    chem_df = df[df['CHEMICAL'] == chem]
    total = len(chem_df)
    acceptable = len(chem_df[chem_df['STATUS'] == 'Acceptable'])
    noncompliant = total - acceptable
    percent = acceptable / total if total > 0 else 0

    summary_ws[f"B{row}"] = chem
    summary_ws[f"C{row}"] = total
    summary_ws[f"E{row}"] = acceptable
    summary_ws[f"G{row}"] = noncompliant
    summary_ws[f"I{row}"] = percent
    summary_ws[f"I{row}"].number_format = '0.00%'

    # Priority
    if percent < 0.45:
        summary_ws[f"K{row}"] = "HIGH"
    elif percent > 0.55:
        summary_ws[f"K{row}"] = "LOW"
    else:
        summary_ws[f"K{row}"] = "MEDIUM"

    row += 1

# Totals row
summary_ws[f"C{row}"] = f"=SUM(C{start_row}:C{row-1})"
summary_ws[f"E{row}"] = f"=SUM(E{start_row}:E{row-1})"
summary_ws[f"G{row}"] = f"=SUM(G{start_row}:G{row-1})"

# Weighted overall score formula in D7
weights_formula_parts = []
for r in range(start_row, row):
    weights_formula_parts.append(f"I{r}*(C{r}/100)")
summary_ws['D7'] = f"={'+'.join(weights_formula_parts)}"
summary_ws['D7'].number_format = '0.00%'

# ---------------------------
# RANGES TABLE
# ---------------------------
ranges_start = row + 3
summary_ws.merge_cells(f"B{ranges_start}:F{ranges_start}")
summary_ws[f"B{ranges_start}"] = "RANGES"
summary_ws[f"B{ranges_start}"].fill = header_fill
summary_ws[f"B{ranges_start}"].font = header_font
summary_ws[f"B{ranges_start}"].alignment = center_bold

ranges_headers = ["Chemical", "Min Concentration", "Max Concentration", "Avg Concentration",
                  "Min pH", "Max pH", "Avg pH",
                  "Min Temp", "Max Temp", "Avg Temp"]

for col_idx, header in enumerate(ranges_headers, start=2):
    cell = summary_ws.cell(row=ranges_start+1, column=col_idx)
    cell.value = header
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_bold

r = ranges_start + 2
for chem in chemicals:
    chem_df = df[df['CHEMICAL'] == chem]
    summary_ws.cell(row=r, column=2, value=chem)
    summary_ws.cell(row=r, column=3, value=chem_df['CONCENTRATION (ppm)'].min())
    summary_ws.cell(row=r, column=4, value=chem_df['CONCENTRATION (ppm)'].max())
    summary_ws.cell(row=r, column=5, value=chem_df['CONCENTRATION (ppm)'].mean())
    summary_ws.cell(row=r, column=6, value=chem_df['pH LEVEL'].min())
    summary_ws.cell(row=r, column=7, value=chem_df['pH LEVEL'].max())
    summary_ws.cell(row=r, column=8, value=chem_df['pH LEVEL'].mean())
    summary_ws.cell(row=r, column=9, value=chem_df['TEMPERATURE (Celsius)'].min())
    summary_ws.cell(row=r, column=10, value=chem_df['TEMPERATURE (Celsius)'].max())
    summary_ws.cell(row=r, column=11, value=chem_df['TEMPERATURE (Celsius)'].mean())
    r += 1



# Save final formatted file
wb.save(save_path)
print(f"Final formatted report saved at: {save_path}")
