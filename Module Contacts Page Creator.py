# -*- coding: utf-8 -*-
"""
Created on Mon Nov 3 18:06:56 2025
@author: nnrw
"""
from dotenv import load_dotenv
import openpyxl
import os
from typing import Dict, List, Union
import pandas as pd
from fuzzywuzzy import fuzz, process
from pathlib import Path
import sys
import re
from datetime import datetime # Kept for HTML generation
from jinja2 import Environment, FileSystemLoader

# --- 1. CONFIGURATION & ENVIRONMENT SETUP ---
load_dotenv('File Addresses.env')
AGGREGATED_MODULE_DATA = os.getenv("AGGREGATED_MODULE_DATA")
CONTACT_DETAILS_FILE_NAME = os.getenv("STAFF_CONTACT_DETAILS")
template_dir = os.getenv("TEMPLATE_DIR")
env = Environment(loader=FileSystemLoader(template_dir))
OUTPUT_DIR = os.getenv("MODULE_CONTACTS_INFO_OUTPUT_DIR")

# --- STAFF/MODULE DATA STRUCTURES ---
StaffDetails = Dict[str, str] 
STAFF_DETAILS_BY_ID: Dict[str, StaffDetails] = {} 
ModuleContributorIDs = Dict[str, List[Dict[str, str]]]
MODULE_CONTRIBUTORS_BY_CODE: ModuleContributorIDs = {}
CourseModules = Dict[int, List[str]]
OrderedUniqueCourseIDs = List[int]
AggregatedModuleStaffData = Dict[int, Dict[str, List[Dict[str, str]]]]

def main():
    status_print("🚀 Starting Python script to generate Module Contacts pages...")
    status_print("")
    successful_page_updates = 0
    total_modules_attempted = 0
    failed_courses = []
    total_courses = 0
    try:
        load_staff_details(STAFF_DETAILS_BY_ID)
        module_contributor_ids = load_module_contributor_ids()
        aggregated_staff_data = aggregate_module_staff_data(
                    module_contributor_ids, 
                    STAFF_DETAILS_BY_ID
                )
        courses_to_process = sorted(aggregated_staff_data.keys())
        total_courses = len(courses_to_process)
        try:
            module_codes = sorted(aggregated_staff_data.keys())
            for module_code in module_codes:
                total_modules_attempted += 1
                staff_list = aggregated_staff_data.get(module_code, [])
                if not staff_list:
                    status_print(f"    ⚠️ DATA MISSING: Module {module_code} has NO staff entries. Skipping.")
                    continue
        
                title = f"{module_code} Module Contacts, 26-27"
                html = generate_module_contacts_page_html(module_code, aggregated_staff_data)
                _create_or_update_with_cache(module_code, title, html, OUTPUT_DIR)
         
                # --- Photo pages ---
                photo_page_title = f"{module_code} Module Contact Photos, 26-27"
                photo_page_html = generate_photo_contacts_page_html(module_code, aggregated_staff_data)
                _create_or_update_with_cache(module_code, photo_page_title, photo_page_html, OUTPUT_DIR)

        except Exception as e:
             status_print(f"    -> ❌ FAILED processing modules: {e}")
             failed_courses.append("module_run")

    except KeyboardInterrupt:
        status_print("\n🛑 Halt detected. Finalizing summary...")
    finally:
        status_print("\n--- Script Summary ---")
        status_print(f"Total Unique Courses Identified: {total_courses}")
        status_print(f"Total Module Page updates attempted: {total_modules_attempted}")
        status_print(f"Total Module Pages successfully updated: {successful_page_updates}")
        if failed_courses:
            status_print(f"Note: {len(failed_courses)} courses failed entirely due to API errors.")
        status_print("✅ Done.")

def _generate_staff_table_html(module_code, staff_list):
    rows = []
    for s in staff_list:
        rows.append(f"<tr><td style='padding:10px; border:1px solid #ccc;'>{s.get('name')}</td>"
                    f"<td style='padding:10px; border:1px solid #ccc;'>{s.get('role')}</td>"
                    f"<td style='padding:10px; border:1px solid #ccc;'><a href='mailto:{s.get('email')}'>{s.get('email')}</a></td>"
                    f"<td style='padding:10px; border:1px solid #ccc;'>{s.get('office')}</td></tr>")
    
    return f"""<table style='width:100%; border-collapse:collapse;'>
                <thead><tr style='background:#f2f2f2;'><th style='padding:10px; border:1px solid #ccc;'>Full Name</th>
                <th style='padding:10px; border:1px solid #ccc;'>Role</th>
                <th style='padding:10px; border:1px solid #ccc;'>Email</th>
                <th style='padding:10px; border:1px solid #ccc;'>Office</th></tr></thead>
                <tbody>{''.join(rows)}</tbody></table>"""

def _generate_short_table_html(module_code, staff_list):
    rows = []
    for s in staff_list:
        rows.append(f"<tr><td style='padding:10px; border:1px solid #ccc;'>{s.get('name')}</td>"
                    f"<td style='padding:10px; border:1px solid #ccc;'>{s.get('role')}</td>"
                    f"<td style='padding:10px; border:1px solid #ccc;'><a href='mailto:{s.get('email')}'>{s.get('email')}</a></td>"
                    )
    
    return f"""<table style='width:100%; border-collapse:collapse;'>
                <thead><tr style='background:#f2f2f2;'><th style='padding:10px; border:1px solid #ccc;'>Full Name</th>
                <th style='padding:10px; border:1px solid #ccc;'>Role</th>
                <th style='padding:10px; border:1px solid #ccc;'>Email</th>
                </tr></thead>
                <tbody>{''.join(rows)}</tbody></table>"""

def _generate_photo_table_html(module_code, staff_list):

    chunks = [staff_list[i:i + 3] for i in range(0, len(staff_list), 3)]
    all_tables_html = []
    for chunk in chunks:
        name_cells = []
        photo_placeholder_cells = []
        email_cells = []
        office_cells = []
        
        cell_width = "33.33%"

        for s in chunk:
            name = s.get('name', '')
            email = s.get('email', '')
            office = s.get('office', '')

            name_cells.append(
                f"<td style='width:{cell_width}; padding:10px; border:1px solid #ccc; "
                f"text-align:center; font-weight:bold; background:#f9f9f9;'>{name}</td>"
            )
            photo_placeholder_cells.append(
                f"<td style='width:{cell_width}; padding:40px 10px; border:1px solid #ccc; "
                f"text-align:center; color:#999; font-size:1.1em; font-weight:bold;'>"
                f"[INSERT PHOTO HERE]</td>"
            )
            email_cells.append(
                f"<td style='width:{cell_width}; padding:10px; border:1px solid #ccc; "
                f"text-align:center;'><a href='mailto:{email}'>{email}</a></td>"
            )
            office_cells.append(
                f"<td style='width:{cell_width}; padding:10px; border:1px solid #ccc; "
                f"text-align:center;'>{office}</td>"
            )

        while len(name_cells) < 3:
            empty_cell = f"<td style='width:{cell_width}; border:none;'></td>"
            name_cells.append(empty_cell)
            photo_placeholder_cells.append(empty_cell)
            email_cells.append(empty_cell)
            office_cells.append(empty_cell)

        table_block = f"""
        <table style='width:100%; border-collapse:collapse; table-layout: fixed; margin-bottom: 20px;'>
            <tbody>
                <tr>{''.join(name_cells)}</tr>
                <tr>{''.join(photo_placeholder_cells)}</tr>
                <tr>{''.join(email_cells)}</tr>
                <tr>{''.join(office_cells)}</tr>
            </tbody>
        </table>
        """
        all_tables_html.append(table_block)

    return "\n".join(all_tables_html)

def _generate_html_footer(module_code: Union[int, str]) -> str:
    """Creates a standard timestamped footer for all generated pages."""

    # Generate the current time at the exact moment this function is called
    timestamp = datetime.now().strftime('%d %B %Y at %H:%M:%S')
    admin_contact = os.getenv("ADMIN_CONTACT")
    return f"""
    <hr style="margin-top: 20px;">
    <p style="margin-top: 20px; font-size: small; color: #6b7280;">This content was generated for module code: <strong>{module_code}</strong> on {timestamp}. If you notice any errors or can help us improve the accuracy of the information provided, please contact {admin_contact} to let us know.</p>
    """
def generate_module_contacts_page_html(module_code: str, module_staff_data: dict) -> str:
    """
    module_staff_data format:
    {
        "MOD123": [staff_dict, staff_dict, ...],
        "MOD456": [...]
    }
    """
    staff_list = module_staff_data.get(module_code, [])
    current_time_str = datetime.now().strftime('%Y-%m-%d at %H:%M:%S')
    footer_html = _generate_html_footer(module_code)

    if not staff_list:
        return f"""
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ccc; border-radius: 8px; background-color: #fff3cd;">
            <h3 style="color: #856404;">⚠️ Staff Data Missing</h3>
            <p>No staff contributors were found or linked for module <strong>{module_code}</strong>.</p>
            <p style="font-size: small; color: #6b7280;">Generated on {current_time_str}.</p>
        </div>
        """

    try:
        main_template = env.get_template('module_contacts_template.html')
        return main_template.render(
            module_code=module_code,
            content=_generate_staff_table_html(module_code, staff_list),
            # content=_generate_short_table_html(module_code, staff_list)
            footer=footer_html
        )

    except Exception as e:
        return f"<p style='color:red;'>Error loading main template: {e}</p>"


def generate_photo_contacts_page_html(module_code: str, module_staff_data: dict) -> str:
    staff_list = module_staff_data.get(module_code, [])
    
    current_time_str = datetime.now().strftime('%Y-%m-%d at %H:%M:%S')
    footer_html = _generate_html_footer(module_code)
    
    if not staff_list:
        return f"""
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ccc; border-radius: 8px; background-color: #fff3cd;">
            <h3 style="color: #856404;">⚠️ Staff Data Missing</h3>
            <p>No staff contributors were found or linked for module <strong>{module_code}</strong>.</p>
            <p style="font-size: small; color: #6b7280;">Generated on {current_time_str}.</p>
        </div>
        """
    try:
        main_template = env.get_template('alternate_module_contacts_template.html')
        return main_template.render(
            module_code=module_code,
            content = _generate_photo_table_html(module_code, staff_list),
            footer=footer_html
        )
    except Exception as e:
        return f"<p style='color:red;'>Error loading main template: {e}</p>"

    return

def _safe_filename(title: str) -> str:
    """Convert page title to a safe filename."""
    # Replace invalid characters and spaces
    filename = re.sub(r"[^\w\s-]", "", title).strip().lower()
    filename = re.sub(r"[-\s]+", "_", filename)
    return f"{filename}.html"

def _create_or_update_with_cache(course, title, body_html, OUTPUT_DIR):
    """Checks the local page_cache dict instead of calling Canvas API every time."""
#    existing_slug = page_cache.get(title)
    
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    file_path = Path(OUTPUT_DIR) / _safe_filename(title)

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(body_html)

    print(f"    -> 💾 HTML file saved: {file_path}")

def load_staff_details(staff_details_map: Dict[str, dict]):
    try:
        df = pd.read_excel(CONTACT_DETAILS_FILE_NAME, header=None)
    except FileNotFoundError:
        status_print(f"Error: '{CONTACT_DETAILS_FILE_NAME}' not found. Staff details not loaded.")
        return
    except Exception as e:
        status_print(f"Error loading staff details file: {e}")
        return

    if df.shape[1] < 4:
        status_print(f"Error: '{CONTACT_DETAILS_FILE_NAME}' does not contain the expected 4 columns (A-D).")
        return

    for _, row in df.iterrows():
        try:
            # Extract raw values
            val_first = row.iloc[0]
            val_surname = row.iloc[1]
            val_email = row.iloc[2]
            val_office = row.iloc[3]

            # Process strings and handle NaNs for each field
            first_name = str(val_first).strip() if pd.notna(val_first) else ""
            surname = str(val_surname).strip() if pd.notna(val_surname) else ""
            email = str(val_email).strip() if pd.notna(val_email) else ""
            office = str(val_office).strip() if pd.notna(val_office) else ""
            
            full_name = f"{first_name} {surname}".strip()

            if full_name:
                staff_details_map[full_name] = {
                    "name": full_name,
                    "email": email,
                    "office": office,
                }
            else:
                status_print("WARNING: Skipping row due to missing staff name.")
        except Exception as e:
            status_print(f"WARNING: An unexpected error occurred processing a row: {e}")

  
    status_print(f"Staff details successfully loaded from {CONTACT_DETAILS_FILE_NAME}. {len(staff_details_map)} entries found.")
    
def load_module_contributor_ids() -> ModuleContributorIDs:
    try:
        df = pd.read_excel(AGGREGATED_MODULE_DATA, sheet_name="Module Contributors")
    except Exception as e:
        status_print(f"Error loading module contributors: {e}")
        return {}

    contributors = {}
    last_module_code = None
    for _, row in df.iterrows():
        current_module_code_raw = row.get('Module Code')
        if pd.notna(current_module_code_raw) and str(current_module_code_raw).strip():
            current_module_code = str(current_module_code_raw).strip()
            last_module_code = current_module_code
        else:
            current_module_code = last_module_code

        if current_module_code:
            if current_module_code not in contributors:
                contributors[current_module_code] = []
            contributors[current_module_code].append({
                'staff_id': row['Contributor Name(s)'],
                'role': row['Description'],
                'contribution': row['Contribution %(s)']
            })
    return contributors

def _normalize_name_for_matching(name: str) -> str:
    if not isinstance(name, str): return ""
    titles = ["dr.", "dr", "professor", "prof.", "prof"]
    words = name.strip().lower().split()
    return " ".join([w for w in words if w not in titles]).strip()

def aggregate_module_staff_data(module_contributor_ids, staff_details_by_id):
    status_print("Starting data aggregation...")
    final_data = {}
    available_staff_names = list(staff_details_by_id.keys())
    module_contributors = module_contributor_ids.items()
    for module_code, module_info in module_contributors:
        staff_list = []
        contributor_ids = module_info
        final_data[module_code] = []
        for staff_id in contributor_ids:
            name_raw = staff_id.get('staff_id')
            if name_raw and available_staff_names:
                norm_name = _normalize_name_for_matching(name_raw)
                match = process.extractOne(norm_name, available_staff_names, score_cutoff=70, scorer=fuzz.token_sort_ratio)
                if match:
                    details = staff_details_by_id.get(match[0], {})
                    staff_list.append({
                        'name': details.get('name', 'N/A'),
                        'email': details.get('email', 'N/A'),
                        'office': details.get('office', 'N/A'),
                        'role': staff_id.get('role'),
                        'contribution': staff_id.get('contribution')
                    })
            if staff_list:
                final_data[module_code] = staff_list
 
    return final_data

def status_print(message: str, file=sys.stdout):
    """Prints a clean message to console WITHOUT a timestamp."""
    print(message, file=file)

def safe_read_cell(sheet: openpyxl.worksheet.worksheet.Worksheet, 
                   row_idx: int, 
                   col_idx: int) -> Union[str, int, float, None]:
    """
    Safely reads a cell value from an openpyxl sheet, returning a cleaned value 
    (stripped string) or None if empty/unreadable.
    Returns the value in its native type (int/float/string) if possible.
    """
    try:
        cell = sheet.cell(row=row_idx, column=col_idx)
        value = cell.value

        if value is None:
            return None
        
        if isinstance(value, str):
            stripped = value.strip()
            return stripped if stripped else None
        
        return value
        
    except Exception:
        return None

if __name__ == "__main__":
    main()