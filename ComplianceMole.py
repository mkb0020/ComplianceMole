# Compliance Mole 
# Analyzes a CSV chemical data report to determine if each sample is within compliance and produces a reformatted excel document with the compliance status and reason for non-compliance(if applicable)

import pandas as pd
import openpyxl
from tkinter import Tk, simpledialog, filedialog
from datetime import datetime
import os

# ---------- Phase 1: File Selection & User Input ----------
def select_file():
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    return file_path

def load_csv(file_path):
    df = pd.read_csv(file_path)
    last_row = len(df)
    return df, last_row

def get_user_info():
    root = Tk()
    root.withdraw()
    first = simpledialog.askstring("User Info", "Enter First Name:")
    middle = simpledialog.askstring("User Info", "Enter Middle Name (Optional):") or ""
    last = simpledialog.askstring("User Info", "Enter Last Name:")
    company = simpledialog.askstring("User Info", "Enter Company Name:")

    return {
        "FirstName": first,
        "MiddleName": middle,
        "LastName": last,
        "FirstIntl": first[0],
        "MidIntl": middle[0] if middle else "",
        "LastIntl": last[0],
        "CompanyName": company,
        "DateToday": datetime.today().strftime("%Y%m%d")
    }

def get_save_path(user_info):
    Tk().withdraw()
    base_name = f"Chemical_Compliance_Report_{user_info['CompanyName']}_{user_info['DateToday']}_{user_info['FirstIntl']}{user_info['MidIntl']}{user_info['LastIntl']}.xlsx"
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=base_name)

    if os.path.exists(save_path):
        version = simpledialog.askstring("Version", "Enter version number:")
        base_name = f"Chemical_Compliance_Report_{user_info['CompanyName']}_{user_info['DateToday']}_{user_info['FirstIntl']}{user_info['MidIntl']}{user_info['LastIntl']}_v{version}.xlsx"
        save_path = os.path.join(os.path.dirname(save_path), base_name)

    return save_path

def rename_columns(df):
    df.columns = [
        "SAMPLE ID",
        "CHEMICAL",
        "CONCENTRATION (ppm)",
        "pH LEVEL",
        "TEMPERATURE (Celsius)",
        "PRESSURE (kPa)",
        "FLOW RATE (L_min)",
        "STATUS",
        "COMMENT"
    ]
    return df

# ---------- Phase 2: Compliance Checking ----------
def load_ranges():
    ranges_path = r"C:\Users\mkb00\PROJECTS\PythonProjects\ComplianceMole\CompliantRanges.xlsx"
    ranges_df = pd.read_excel(ranges_path)
    ranges_df.set_index("Chemical", inplace=True)
    return ranges_df

def check_compliance(df, ranges_df):
    for idx, row in df.iterrows():
        chem = row["CHEMICAL"]
        if chem not in ranges_df.index:
            df.at[idx, "STATUS"] = "UNKNOWN CHEMICAL"
            df.at[idx, "COMMENT"] = "No compliance data found."
            continue

        limits = ranges_df.loc[chem]

        # Track violations
        issues = []

        # Concentration
        if not (limits["Concentration_ppm_Min"] <= row["CONCENTRATION (ppm)"] <= limits["Concentration_ppm_Max"]):
            issues.append(f"Concentration not within acceptable range: {limits['Concentration_ppm_Min']} - {limits['Concentration_ppm_Max']}")

        # pH
        if not (limits["pH_Level_Min"] <= row["pH LEVEL"] <= limits["pH_Level_Max"]):
            issues.append(f"pH not within acceptable range: {limits['pH_Level_Min']} - {limits['pH_Level_Max']}")

        # Temperature
        if not (limits["Temperature_C_Min"] <= row["TEMPERATURE (Celsius)"] <= limits["Temperature_C_Max"]):
            issues.append(f"Temperature not within acceptable range: {limits['Temperature_C_Min']} - {limits['Temperature_C_Max']}")

        # Pressure
        if not (limits["Pressure_kPa_Min"] <= row["PRESSURE (kPa)"] <= limits["Pressure_kPa_Max"]):
            issues.append(f"Pressure not within acceptable range: {limits['Pressure_kPa_Min']} - {limits['Pressure_kPa_Max']}")

        # Flow Rate
        if not (limits["Flow_Rate_L_min_Min"] <= row["FLOW RATE (L_min)"] <= limits["Flow_Rate_L_min_Max"]):
            issues.append(f"Flow Rate not within acceptable range: {limits['Flow_Rate_L_min_Min']} - {limits['Flow_Rate_L_min_Max']}")

        # Final status/comment
        if issues:
            df.at[idx, "STATUS"] = "NON-COMPLIANT"
            df.at[idx, "COMMENT"] = " | ".join(issues)
        else:
            df.at[idx, "STATUS"] = "COMPLIANT"
            df.at[idx, "COMMENT"] = "Within Acceptable Ranges"

    return df

# ---------- Main ----------
if __name__ == "__main__":
    csv_path = select_file()
    if not csv_path:
        print("No file selected. Exiting.")
        exit()

    df, last_row = load_csv(csv_path)
    user_info = get_user_info()
    save_path = get_save_path(user_info)

    df = rename_columns(df)
    ranges_df = load_ranges()
    df = check_compliance(df, ranges_df)

    df.to_excel(save_path, index=False)
    print(f"Done! File saved at: {save_path}")

