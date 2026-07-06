from dotenv import load_dotenv
import pandas as pd
import re
import os
from openpyxl import load_workbook

# --- 1. CONFIGURATION & ENVIRONMENT SETUP ---
load_dotenv('File Addresses.env')

# Load the Excel file
input_file = os.getenv("AGGREGATED_MODULE_DATA")

if not os.path.exists(input_file):
    print("\n❌ Input file not found.")
    print(f"   Expected file: {os.path.abspath(input_file)}")
    print("   Please ensure the aggregated module data file exists and try again.\n")
    raise SystemExit(1)

df = pd.read_excel(input_file, engine="openpyxl")

# Function to clean and split names
def process_name(name):
    if pd.isna(name):
        return "", ""

    # Remove titles
    name = re.sub(r'^(Dr|Professor)\s+', '', str(name), flags=re.IGNORECASE)

    parts = name.strip().split()

    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], ""
    else:
        return parts[0], parts[-1]

# Apply name processing
results = df["Contributor Name(s)"].apply(process_name)

# Create DataFrame
output_df = pd.DataFrame(results.tolist(), columns=["First Name", "Surname"])

# Add empty columns
output_df["Email Address"] = ""
output_df["Office Address"] = ""

# --- Remove duplicates ---
output_df = output_df.drop_duplicates(subset=["First Name", "Surname"])

# --- Custom sorting logic ---
def is_valid_surname(s):
    if not isinstance(s, str) or s.strip() == "":
        return False
    return not s.strip().isdigit()

output_df["sort_flag"] = output_df["Surname"].apply(
    lambda x: 0 if is_valid_surname(x) else 1
)

# Sort
output_df = output_df.sort_values(
    by=["sort_flag", "Surname"],
    ascending=[True, True]
)

# Drop helper column
output_df = output_df.drop(columns=["sort_flag"])
# %%


# Write to Excel
output_file = "Module Contacts Initialisation File.xlsx"
output_df.to_excel(output_file, index=False, engine="openpyxl")

# Adjust column widths
workbook = load_workbook(output_file)
sheet = workbook.active

column_widths = {
    "A": 20,
    "B": 25,
    "C": 30,
    "D": 35
}

for col, width in column_widths.items():
    sheet.column_dimensions[col].width = width

workbook.save(output_file)

print(f"Output written to {output_file}")