# -*- coding: utf-8 -*-
"""
'Canvas Baseline' Assessment Information Installer Script
Created on Thu Dec 11 17:15:52 2025
@author: nnrw

This script automates the creation and updating of assessment information pages in Canvas.
It maps Excel data to Canvas Course IDs, generates HTML content via Jinja2 templates,
and performs safety checks to ensure no student data is overwritten.
"""
from dotenv import load_dotenv
import re
import os
import sys
import openpyxl
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Union, Tuple
from jinja2 import Environment, FileSystemLoader

# --- 1. CONFIGURATION & ENVIRONMENT SETUP ---
load_dotenv('File Addresses.env')
template_dir = os.getenv("TEMPLATE_DIR")
env = Environment(loader=FileSystemLoader(template_dir))
OUTPUT_DIR = os.getenv("ASSESSMENT_INFO_OUTPUT_DIR")
AGGREGATED_MODULE_DATA = os.getenv("AGGREGATED_MODULE_DATA")

COL_CCINDEX = {'MODULE_CODE': 1, 'MODULE_TITLE': 2, 'COURSE_ID': 4}
COL_AMD = {
    'MODULE_TITLE': 2, 'EXAM_ASSESSMENT': 7, 'EXAM_DURATION': 8, 
    'EXAM_SEM': 9, 'EXAM_WEIGHTING': 10, 'EXAM_COMMENT': 11,
    'COURSEWORK_ASSESSMENT': 12, 'COURSEWORK_SEM': 13, 
    'COURSEWORK_WEIGHTING': 15, 'COURSEWORK_COMMENT': 16,
    'PF_DESCRIPTION': 17, 'PF_DETAIL': 18
}

def main():
    stats = {"courses_processed": 0, "modules_processed": 0, "pages_created": 0, "errors": 0}
    try:
        aggregated_data = process_aggregated_data(AGGREGATED_MODULE_DATA)
    except Exception as e:
        print(f"Error reading Aggregated Data file: {e}")
        return

    try:
        # --- 3. CORE PROCESSING LOOP (PER MODULE) ---
        for module_code, module_config in aggregated_data.items():
            print("-" * 50)
            try:
                stats["modules_processed"] += 1
                print(module_code)
    
                # Determine which specific instruction sub-pages are required
                module_slugs = {
                    'written': None,
                    'digital': None,
                    'coursework': None,
                    'take_home': None
                }
    
                # Create the Main Overview Page for this module
                main_title = f"{module_code} Assessment Information 26-27"
                content_main = generate_assessment_overview(
                    module_code,          # replaces course_id
                    module_config,
                    module_slugs['written'],
                    module_slugs['digital'],
                    module_slugs['coursework'],
                    module_slugs['take_home'],
                    env
                )
    
                assessment_page_obj = _create_or_update_page(main_title, content_main, OUTPUT_DIR)
                
                if assessment_page_obj is None:
                    stats["pages_created"] += 1
                else:
                    stats["errors"] += 1
                    
            except Exception as e:
                print(f"CRITICAL ERROR in module {module_code}: {e}")
                stats["errors"] += 1
    
    finally:
        print("\n" + "=" * 30)
        print("         STATISTICS")
        print("=" * 30)
        print(f"Modules Processed:  {stats['modules_processed']}")
        print(f"Pages Created or Updated:     {stats['pages_created']}")
        print(f"Errors:            {stats['errors']}")
        print("=" * 30)
        print("\n✅ Done.")

def generate_assessment_overview(
    module_code: Union[int, str], 
    module_config: Dict[str, Any],
    written_exam_instructions_page_slug: str, 
    digital_exam_instructions_page_slug: str, 
    coursework_instructions_page_slug: str, 
    take_home_exam_instructions_page_slug: str,
    env: Any 
) -> str:
    """
    Generates the HTML content by rendering sub-components and then 
    compiling them into a final master template.
    """
    
    def render_blue_button(description: str, slug: str, button_text: str) -> str:
        template_vars = {
            "slug": slug,
            "description": description,
            "button_text": button_text
        }
        try:
            template = env.get_template('blue_button_template.html')
            return template.render(template_vars)
        except Exception as e:
            return f"<p style='color:red;'>Error loading button template: {e}</p>"

    module_code = module_config.get('module_code', 'UNKNOWN')
    exam_assessments = module_config.get("exam_assessment", [])
    coursework_assessments = module_config.get("coursework_assessment", [])
    pf_component_assessments = module_config.get("PF component", [])
    
    has_exam = bool(exam_assessments)
    has_coursework = bool(coursework_assessments)
    has_pf_component = bool(pf_component_assessments)
    
    written_exams = [item for item in exam_assessments if "Written Examination" in item.get('assessment', '')]
    has_24hr_take_home_exam = any(item.get('duration') == "1440" for item in written_exams)
    has_standard_written_exam = any(item.get('duration') != "1440" for item in written_exams)
    has_digital_exam = any("Digital Examination" in item.get('assessment', '') for item in exam_assessments)

    exam_blocks_html = ""
    if has_exam:
        exam_table_rows_html = _generate_exam_table_rows(exam_assessments)
        exam_blocks_html += f"""
            <h3 style="color: #1f2937; padding-bottom: 10px; margin-top: 30px;">{module_code} Exams:</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;" border="1" cellspacing="0" cellpadding="5">
                <thead>
                    <tr style="background-color: #f8f8f8;">
                        <th style="width: 20%; text-align: left; padding: 10px;"><strong>Assessment</strong></th>
                        <th style="width: 10%; text-align: center; padding: 10px;"><strong>Semester</strong></th>
                        <th style="width: 10%; text-align: center; padding: 10px;"><strong>Weight (%)</strong></th>
                        <th style="width: 10%; text-align: center; padding: 10px;"><strong>Duration (Mins)</strong></th>
                        <th style="width: 50%; text-align: left; padding: 10px;"><strong>Detail</strong></th>
                    </tr>
                </thead>
                <tbody>{exam_table_rows_html}</tbody>
            </table>
        """
        
        if has_24hr_take_home_exam:
            exam_blocks_html += render_blue_button(
                description="Further information about <strong>24-hour Take Home Exams</strong> is available from the link below:",
                slug=os.getenv("TAKE_HOME_EXAM_PAGE_SLUG"),
                button_text="Further Information about Take Home (24 hr) Exams"
            )
        
        if has_standard_written_exam:
            exam_blocks_html += render_blue_button(
                description="Further information about invigilated <strong>Written</strong> exam rules and resources is available from the link below:",
                slug=os.getenv("WRITTEN_EXAMS_PAGE_SLUG"),
                button_text="Further Information about Written Exams"
            )

        if has_digital_exam:
            exam_blocks_html += render_blue_button(
                description="The rules and arrangements that apply to <strong>Digital</strong> examinations are further described by the page linked below:",
                slug=os.getenv("DIGITAL_EXAM_PAGE_SLUG"),
                button_text="Further Information about Digital Exams"
            )

    assessment_blocks_html = ""
    if has_coursework:
        coursework_table_rows_html = _generate_coursework_table_rows(coursework_assessments)
        assessment_blocks_html += f"""
            <h3 style="color: #1f2937; padding-bottom: 10px; margin-top: 30px;">{module_code} Coursework:</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;" border="1" cellspacing="0" cellpadding="5">
                <thead>
                    <tr style="background-color: #f8f8f8;">
                        <th style="width: 20%; text-align: left; padding: 10px;"><strong>Assessment</strong></th>
                        <th style="width: 10%; text-align: center; padding: 10px;"><strong>Semester</strong></th>
                        <th style="width: 10%; text-align: center; padding: 10px;"><strong>Weight (%)</strong></th>
                        <th style="width: 50%; text-align: left; padding: 10px;"><strong>Detail</strong></th>
                    </tr>
                </thead>
                <tbody>{coursework_table_rows_html}</tbody>
            </table>
        """

    if has_pf_component:
        pf_table_rows_html = _generate_pf_table_rows(pf_component_assessments)
        assessment_blocks_html += f"""
            <h3 style="color: #1f2937; padding-bottom: 10px; margin-top: 30px;">{module_code} Pass/Fail Component:</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;" border="1" cellspacing="0" cellpadding="5">
                <thead>
                    <tr style="background-color: #f8f8f8;">
                        <th style="width: 20%; text-align: left; padding: 10px;"><strong>Assessment</strong></th>
                        <th style="width: 50%; text-align: left; padding: 10px;"><strong>Detail</strong></th>
                    </tr>
                </thead>
                <tbody>{pf_table_rows_html}</tbody>
            </table>
        """

    if has_coursework or has_pf_component:
        assessment_blocks_html += render_blue_button(
            description="Follow the link below for further information about <strong>Coursework</strong> assessment rules and policies:",
            slug=os.getenv("COURSEWORK_PAGE_SLUG"),
            button_text="Further Information about Coursework"
        )
        
    try:
        main_template = env.get_template('assessment_overview_main.html')
        return main_template.render(
            module_code=module_code,
            final_html_body=exam_blocks_html + assessment_blocks_html,
            footer=_generate_html_footer(module_code)
        )
    except Exception as e:
        return f"<p style='color:red;'>Error loading main template: {e}</p>"

def _generate_html_footer(module_code: Union[int, str]) -> str:
    """Creates a standard timestamped footer for all generated pages."""
    timestamp = datetime.now().strftime('%d %B %Y at %H:%M:%S')
    admin_contact = os.getenv("ADMIN_CONTACT")
    return f"""
    <hr style="margin-top: 20px;">
    <p style="margin-top: 20px; font-size: small; color: #6b7280;">This content was generated for module code: <strong>{module_code}</strong> on {timestamp}. If you notice any errors or can help us improve the accuracy of the information provided, please contact {admin_contact} to let us know.</p>
    """

def _generate_exam_table_rows(rows: List[Dict[str, Any]]) -> str:
    """Constructs HTML <tr> elements for the Exam overview table."""
    ROW_TEMPLATE = """
<tr>
    <td style="width: 20%; text-align: left; padding: 10px;">{assessment}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{semester}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{weight}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{duration}</td>
    <td style="width: 50%; text-align: left; padding: 10px;">{exam comment}</td>
</tr>
"""
    return "\n".join(ROW_TEMPLATE.format(**row) for row in rows)

def _generate_coursework_table_rows(rows: List[Dict[str, Any]]) -> str:
    """Constructs HTML <tr> elements for the Coursework overview table."""
    ROW_TEMPLATE = """
<tr>
    <td style="width: 25%; text-align: left; padding: 10px;">{assessment}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{semester}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{weight}</td>
    <td style="width: 30%; text-align: left; padding: 10px;">{coursework comment}</td>
</tr>
"""
    return "\n".join(ROW_TEMPLATE.format(**row) for row in rows)

def _generate_pf_table_rows(rows: List[Dict[str, Any]]) -> str:
    """Constructs HTML <tr> elements for Pass/Fail component table."""
    ROW_TEMPLATE = """
<tr>
    <td style="width: 40%; text-align: left; padding: 10px;">{PF detail}</td>
    <td style="width: 10%; text-align: center; padding: 10px;">{semester}</td>
</tr>
"""
    return "\n".join(ROW_TEMPLATE.format(**row) for row in rows)

def _safe_filename(title: str) -> str:
    """Convert page title to a safe filename."""
    # Replace invalid characters and spaces
    filename = re.sub(r"[^\w\s-]", "", title).strip().lower()
    filename = re.sub(r"[-\s]+", "_", filename)
    return f"{filename}.html"


def _create_or_update_page(title: str, body_html: str, OUTPUT_DIR) -> None:
    """
    Upsert logic for Canvas Pages.
    Also exports the HTML content to a local .html file.
    """
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    file_path = Path(OUTPUT_DIR) / _safe_filename(title)

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(body_html)

    print(f"    -> 💾 HTML file saved: {file_path}")
    
def process_aggregated_data(data_filepath: str) -> Tuple[Dict[int, List[Dict[str, Any]]]]:
    """
    Extracts assessment data from 'Aggregated_Module_Data.xlsx'.
    
    """
    print(f"\nProcessing aggregated data from '{data_filepath}'...")

    final_data_structure: Dict[int, List[Dict[str, Any]]] = {}
    modules_found_in_sheet = set()
    
    if not os.path.exists(data_filepath):
        print(f"Error: Aggregated Data file not found at '{data_filepath}'.")
        return final_data_structure

    try:
        workbook = openpyxl.load_workbook(data_filepath, data_only=True)
        sheet = workbook["Module Assessments"] 
    except Exception as e:
        print(f"Error loading Aggregated Data: {e}")
        return final_data_structure

    rows_processed = 0
    current_module_code: Union[str, None] = None
    
    for row_idx in range(2, sheet.max_row + 1):
        cell_module_code = safe_read_cell(sheet, row_idx, COL_AMD['MODULE_TITLE'])
        
        # If the code cell is empty, we assume the row belongs to the last seen module code
        if cell_module_code:
            current_module_code = str(cell_module_code).strip().upper()
        
        if current_module_code:
            normalized_module_code = current_module_code
            
            modules_found_in_sheet.add(normalized_module_code)
            
            exam_assessment = safe_read_cell(sheet, row_idx, COL_AMD['EXAM_ASSESSMENT'])
            coursework_assessment = safe_read_cell(sheet, row_idx, COL_AMD['COURSEWORK_ASSESSMENT'])
            pf_component = safe_read_cell(sheet, row_idx, COL_AMD['PF_DESCRIPTION'])

            # Only process rows that have at least one type of assessment
            if not (exam_assessment or coursework_assessment or pf_component):
                continue

            rows_processed += 1


            if normalized_module_code not in final_data_structure:
                final_data_structure[normalized_module_code] = {
                    "module_code": normalized_module_code,
                    "exam_assessment": [],
                    "coursework_assessment": [],
                    "PF component": []
                    }

            # Check if we already started an entry for this module under this Course ID
            module_entry = final_data_structure[normalized_module_code]

            def get_val(col_key):
                return safe_read_cell(sheet, row_idx, COL_AMD[col_key]) or "N/A"
            
            # Map specific assessment details into sub-dictionaries
            if exam_assessment:
                exam_row = {
                    "assessment": exam_assessment,
                    "weight": get_val('EXAM_WEIGHTING'),
                    "semester": get_val('EXAM_SEM'),
                    "duration": get_val('EXAM_DURATION'),
                    "exam comment": get_val('EXAM_COMMENT'),
                }
                module_entry["exam_assessment"].append(exam_row)

            if coursework_assessment:
                coursework_row = {
                    "assessment": coursework_assessment,
                    "weight": get_val('COURSEWORK_WEIGHTING'),
                    "semester": get_val('COURSEWORK_SEM'),
                    "coursework comment": get_val('COURSEWORK_COMMENT'),
                    "release date": "N/A", 
                    "submission date": "N/A",
                }
                module_entry["coursework_assessment"].append(coursework_row)

            if pf_component:
                pf_component_row = {
                    "assessment": pf_component,
                    "PF detail": get_val('PF_DETAIL'),
                    "weight": "N/A",
                    "semester": "N/A",
                    "release date": "N/A", 
                    "submission date": "N/A", 
                }
                module_entry["PF component"].append(pf_component_row)
    
    return final_data_structure

def safe_read_cell(sheet: openpyxl.worksheet.worksheet.Worksheet, 
                   row_idx: int, 
                   col_idx: int) -> Union[str, int, float, None]:
    """
    Safely extracts value from an Excel cell.
    
    Returns stripped string if the value is text, the native number if numeric,
    or None if the cell is empty or an error occurs.
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

def _check_module_assessment_types(module_config: Dict[str, Any]) -> Dict[str, bool]:
    """
    Analyzes a single module's assessment data to determine the presence of 
    Written, Digital, and Coursework/PF components.
    
    Returns a dictionary of boolean flags indicating which page types are needed.
    """
    has_written = False
    has_digital = False
    
    exam_assessments = module_config.get("exam_assessment", [])
    
    for item in exam_assessments:
        assessment_detail = item.get('assessment', '').lower()
        
        if "written examination" in assessment_detail:
            has_written = True
        if "digital" in assessment_detail or "online exam" in assessment_detail:
            has_digital = True
        if has_written and has_digital:
            break 

    has_coursework = bool(module_config.get("coursework_assessment", []))
    has_pf_component = bool(module_config.get("PF component", []))
    
    needs_coursework_page = has_coursework or has_pf_component
        
    return {
        'written': has_written,
        'digital': has_digital,
        'coursework': needs_coursework_page
    }

def get_all_exam_durations(
    course_id: Union[int, str], 
    assessment_data: Dict[Union[int, str], List[Dict[str, Any]]]
) -> List[Dict[str, str]]:
    """
    Scans the assessment dictionary to find all exam durations for a specific course.
    Useful for modules that have multiple exams with varying lengths.
    """
    exam_duration_list: List[Dict[str, str]] = []
    
    try:
        lookup_key = int(course_id)
    except (ValueError, TypeError):
        return exam_duration_list

    module_data_list = assessment_data.get(lookup_key, [])
    
    for module_data in module_data_list:
        if isinstance(module_data, dict):
            exam_assessments = module_data.get("exam_assessment", [])
            for assessment_dict in exam_assessments:
                assessment_name = assessment_dict.get("assessment")
                duration_value = assessment_dict.get("duration")
                
                if assessment_name and duration_value:
                    exam_duration_list.append({
                        "assessment": str(assessment_name),
                        "duration": str(duration_value)
                    })
                    
    return exam_duration_list

if __name__ == "__main__":
    try:  
        main()
        pass
    except KeyboardInterrupt:
        print("\n\n👋 Program execution stopped by user (Ctrl+C). Exiting gracefully.")
        sys.exit(0)