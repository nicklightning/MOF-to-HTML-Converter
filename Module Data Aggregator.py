from dotenv import load_dotenv
import re
import os
import glob
import openpyxl
from openpyxl.styles import Alignment, Font

# --- 1. CONFIGURATION & ENVIRONMENT SETUP ---
load_dotenv('File Addresses.env')
INPUT_FOLDER = os.getenv("XML_MOF_FILES_DIR")
OUTPUT_FILE = os.getenv("AGGREGATED_MODULE_DATA")

# --- Utility Functions ---
def custom_sort_key(filename):
    """
    Extracts the 4th, 5th, and 6th characters from the filename (indices 3, 4, 5)
    and converts them to an integer for numerical sorting (e.g., 'MOF123.xml' -> 123).
    """
    try:
        # Get characters at index 3, 4, and 6
        numeric_part = filename[3:6]
        return int(numeric_part)
    except (IndexError, ValueError):
        # Handle files that don't match the expected pattern
        return float('inf') # Places non-matching files at the end

def save_output_file(output_wb, output_file):
    """Handles saving the output workbook with error messages."""
    try:
        output_wb.save(output_file)
        print(f"\n[SAVE] Successfully saved aggregated data to: '{output_file}'")
    except Exception as e:
        print(f"\n[ERROR] Could not save output file. Please ensure '{output_file}' is not open in another program.")
        print(f"Details: {e}")

def extract_module_data_from_xml(filepath: str) -> dict:
    """
    Reads an XML file, performs character substitution for line breaks,
    and extracts all required data.
    """
    try:
        # Attempt to open and read the existing file
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"ERROR: File not found at '{filepath}'.")
        return {}
    except IOError as e:
        print(f"ERROR: Error reading file: {e}")
        return {}

    # --- CHARACTER SUBSTITUTION & CLEANUP ---
    # 1. Replace &#xD; (Carriage Return) with \n
    content = content.replace('&#xD;', '\n')
    
    # 2. Replace &#xA; (Line Feed) with \n
    content = content.replace('&#xA;', '\n')

    # 3. Trim multiple consecutive newlines down to a single newline (\n)
    # This ensures no blank lines are present, making list items appear tightly packed.
    content = re.sub(r'\n{2,}', '\n', content)


    # --- REGEX PATTERNS ---
    # 1. Primary Data Regex (A-F)
    code_pattern = re.compile(r'\bCode="([^"]*)"', re.DOTALL)
    title_pattern = re.compile(r'ShortTitle="([^"]*)"', re.DOTALL)
    s1_credit_pattern = re.compile(r'\bSemester1CreditValue="([^"]*)"', re.DOTALL)
    s2_credit_pattern = re.compile(r'\bSemester2CreditValue="([^"]*)"', re.DOTALL)
    s3_credit_pattern = re.compile(r'\bSemester3CreditValue="([^"]*)"', re.DOTALL)

    # 2. Specific Exam Comment Regex (Column K)
    exam_comment_pattern = re.compile(
        r'<Details\d+\s+[^>]*Comment2="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )

    # 3. Description 5 Assessment Regex (Column Q)
    description5_pattern = re.compile(
        r'<Details\d+\s+[^>]*Description5="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )

    # 4. Specific Assessment Comment Regex (Column R)
    specific_comment_pattern = re.compile(
        r'<Details\d+\s+[^>]*Comment8="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )

    # 5. Contributor Data Regex (G, H, I)
    contributor_pattern = re.compile(
        r'<Details\s+ContributorName="\s*([^"]*)"\s+Description="([^"]*)"[^>]*ContributionPercentage="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )

    # 6. Primary Assessment Data Regex (G, H, I, J)
    assessment_pattern = re.compile(
        r'<Details\d+\s+Description1="([^"]*)"\s+Length1="([^"]*)"\s+Semester1="([^"]*)"[^>]*Percentage="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )

    # 7. Secondary Assessment Data Regex (L, M, N, O, P)
    assessment2_pattern = re.compile(
        r'<Details\d+\s+Description2="([^"]*)"\s+Semester2="([^"]*)"\s+WhenSet2="([^"]*)"[^>]*Percentage2="([^"]*)"\s+Comment4="([^"]*)"',
        re.DOTALL | re.IGNORECASE
    )
    
    # 8. Syllabus Outline Regex
    syllabus_pattern = re.compile(r'\bOutlineOfSyllabus="([^"]*)"', re.DOTALL)
    
    # --- FIND ALL MATCHES ---
    module_codes = code_pattern.findall(content)
    short_titles = title_pattern.findall(content)
    s1_credits = s1_credit_pattern.findall(content)
    s2_credits = s2_credit_pattern.findall(content)
    s3_credits = s3_credit_pattern.findall(content)
    exam_comments = exam_comment_pattern.findall(content)
    description5_assessments = description5_pattern.findall(content)
    specific_comments = specific_comment_pattern.findall(content)
    contributors = contributor_pattern.findall(content)
    assessments = assessment_pattern.findall(content)
    assessments2 = assessment2_pattern.findall(content)
    syllabus_outlines = syllabus_pattern.findall(content)

    # Process Contributors
    contributor_names = [c[0].strip() for c in contributors]
    contributor_descriptions = [c[1].strip() for c in contributors]
    contributor_percentages = [c[2].strip() for c in contributors]

    # Process Primary Assessments
    assessment_descriptions = [a[0].strip() for a in assessments]
    assessment_lengths = [a[1].strip() for a in assessments]
    assessment_semesters = [a[2].strip() for a in assessments]
    assessment_percentages = [a[3].strip() for a in assessments]

    # Process Secondary Assessments
    assessment2_descriptions = [a[0].strip() for a in assessments2]
    assessment2_semesters = [a[1].strip() for a in assessments2]
    assessment2_when_sets = [a[2].strip() for a in assessments2]
    assessment2_percentages = [a[3].strip() for a in assessments2]
    assessment2_comments = [a[4].strip() for a in assessments2]

    results = {}

    # Extract the first valid match for single fields, using "" as fallback
    results["module_code"] = module_codes[0] if module_codes else ""
    results["short_title"] = short_titles[0] if short_titles else ""
    results["s1_credits"] = s1_credits[0] if s1_credits else ""
    results["s2_credits"] = s2_credits[0] if s2_credits else ""
    results["s3_credits"] = s3_credits[0] if s3_credits else ""
    results["exam_comment"] = exam_comments[0] if exam_comments else ""
    results["description5_assessment"] = description5_assessments[0] if description5_assessments else ""
    results["practical_skills_comment"] = specific_comments[0] if specific_comments else ""
    results["syllabus_outline"] = syllabus_outlines[0] if syllabus_outlines else ""

    # Assign the lists
    results["contributor_names"] = contributor_names
    results["contributor_descriptions"] = contributor_descriptions
    results["contributor_percentages"] = contributor_percentages
    results["assessment_descriptions"] = assessment_descriptions
    results["assessment_lengths"] = assessment_lengths
    results["assessment_semesters"] = assessment_semesters
    results["assessment_percentages"] = assessment_percentages
    results["assessment2_descriptions"] = assessment2_descriptions
    results["assessment2_semesters"] = assessment2_semesters
    results["assessment2_when_sets"] = assessment2_when_sets
    results["assessment2_percentages"] = assessment2_percentages
    results["assessment2_comments"] = assessment2_comments

    return results

def aggregate_module_data():
    """
    Main function to find XML files, extract data, and compile the aggregate spreadsheet.
    """
    if not os.path.exists(INPUT_FOLDER):
        print(f"Error: The folder '{INPUT_FOLDER}' was not found.")
        print("Please create the folder and place your XML MOF files inside.")
        return

    # 1. Find all XML files and apply custom sorting
    search_path = os.path.join(INPUT_FOLDER, '*.xml')
    all_files = glob.glob(search_path)

    # Sort files numerically based on the 4th, 5th, and 6th digits of the filename
    sorted_files = sorted(all_files, key=custom_sort_key)

    if not sorted_files:
        print(f"No XML files (*.xml) found in the '{INPUT_FOLDER}' folder.")
        return

    # 2. Setup the Output Workbook
    output_wb = openpyxl.Workbook()

    # Alignment for headers (center)
    header_alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    
    # Alignment for Syllabus Outline: top-left (Col G)
    syllabus_cell_alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')
    
    # NEW: Alignment for Syllabus Outline's primary data (Cols A-F): align top
    primary_data_top_alignment = Alignment(vertical='top')
    
    bold_font = Font(bold=True)
    
    # Remove the default sheet created by openpyxl
    if 'Sheet' in output_wb.sheetnames:
        output_wb.remove(output_wb['Sheet'])


    # --- Worksheet 1: Module Contributors ---
    output_ws_contrib = output_wb.create_sheet(title="Module Contributors", index=0)

    # Define headers for Contributors sheet
    headers_contrib = [
        'Filename', 'Module Code', 'Short Title', 'S1 Credits', 'S2 Credits', 'S3 Credits', # A-F
        'Contributor Name(s)', 'Description', 'Contribution %(s)' # G-I
    ]
    output_ws_contrib.append(headers_contrib)

    # Apply wrapped text alignment to the header row (Row 1)
    for cell in output_ws_contrib[1]:
        cell.alignment = header_alignment
        cell.font = bold_font

    # Set sensible column widths
    column_widths_contrib = {
        'A': 30, 'B': 15, 'C': 40, 'D': 10, 'E': 10, 'F': 10, 'G': 35,
        'H': 40, 'I': 15, 
    }

    for col_letter, width in column_widths_contrib.items():
        output_ws_contrib.column_dimensions[col_letter].width = width
        output_ws_contrib.row_dimensions[1].height = 40

    # --- Worksheet 2: Module Assessments ---
    output_ws_assessments = output_wb.create_sheet(title="Module Assessments", index=1)

    # Define headers for Assessments sheet
    headers_assessments = [
        'Filename', 'Module Code', 'Short Title', 'S1 Credits', 'S2 Credits', 'S3 Credits', # A-F
        'Exams', 'Exam Duration', 'Semester', 'Weighting', # G-J
        'Exam Comment', # K
        'Other Assessment', 'Other Semester', 'When Set', 'Other Weighting', 'Assessment Comment',# L-P
        'Description 5 Assessment', # Q
        'Skills Assessment (P/F)' # R
    ]
    output_ws_assessments.append(headers_assessments)

    # Apply wrapped text alignment to the header row (Row 1)
    for cell in output_ws_assessments[1]:
        cell.alignment = header_alignment
        cell.font = bold_font

    # Set sensible column widths for Assessments sheet
    column_widths_assessments = {
        'A': 30, 'B': 15, 'C': 40, 'D': 10, 'E': 10, 'F': 10, 'G': 25,
        'H': 18, 'I': 12, 'J': 15,
        'K': 60, 
        'L': 25, 'M': 15, 'N': 15, 'O': 15, 'P': 60,
        'Q': 35, 
        'R': 40,
    }
    for col_letter, width in column_widths_assessments.items():
        output_ws_assessments.column_dimensions[col_letter].width = width
        output_ws_assessments.row_dimensions[1].height = 40
        
    # --- Worksheet 3: Syllabus Outline ---
    output_ws_syllabus = output_wb.create_sheet(title="Syllabus Outline", index=2)
    
    # Define headers for Syllabus sheet
    headers_syllabus = [
        'Filename', 'Module Code', 'Short Title', 'S1 Credits', 'S2 Credits', 'S3 Credits', # A-F
        'Syllabus Outline' # G
    ]
    output_ws_syllabus.append(headers_syllabus)
    
    # Apply wrapped text alignment to the header row (Row 1)
    for cell in output_ws_syllabus[1]:
        cell.alignment = header_alignment
        cell.font = bold_font
        
    # Set sensible column widths for Syllabus sheet
    column_widths_syllabus = {
        'A': 30, 'B': 15, 'C': 40, 'D': 10, 'E': 10, 'F': 10, 
        'G': 120 
    }
    for col_letter, width in column_widths_syllabus.items():
        output_ws_syllabus.column_dimensions[col_letter].width = width
        output_ws_syllabus.row_dimensions[1].height = 40

    print(f"Found {len(sorted_files)} XML files. Starting data extraction...")

    interrupted = False

    try:    
    # 3. Process each file
        for file_path in sorted_files:
            file_name = os.path.basename(file_path)
            print(f"Processing: {file_name}")
    
            # Extract data
            data = extract_module_data_from_xml(file_path)
    
            if not data:
                continue
    
            # Data for iteration
            names = data.get("contributor_names", [])
            descriptions_contrib = data.get("contributor_descriptions", [])
            percentages = data.get("contributor_percentages", [])
            num_contributors = len(names)
    
            descriptions = data.get("assessment_descriptions", [])
            lengths = data.get("assessment_lengths", [])
            semesters = data.get("assessment_semesters", [])
            percentages_assess = data.get("assessment_percentages", [])
            num_assessments = len(descriptions)
    
            descriptions2 = data.get("assessment2_descriptions", [])
            semesters2 = data.get("assessment2_semesters", [])
            when_sets2 = data.get("assessment2_when_sets", [])
            percentages_assess2 = data.get("assessment2_percentages", [])
            comments4 = data.get("assessment2_comments", [])
            num_assessments2 = len(descriptions2)
    
            # Single Data Fields
            exam_comment = data.get("exam_comment", "")
            description5_assessment = data.get("description5_assessment", "")
            practical_comment = data.get("practical_skills_comment", "")
            syllabus_outline = data.get("syllabus_outline", "")
    
            # Data elements for the primary row (Columns A-F).
            primary_data_cols = [
                file_name,
                data.get("module_code", ""),
                data.get("short_title", ""),
                data.get("s1_credits", ""),
                data.get("s2_credits", ""),
                data.get("s3_credits", ""),
            ]
    
            # --- Populate Syllabus Outline sheet ---
            syllabus_row = primary_data_cols + [syllabus_outline]
            output_ws_syllabus.append(syllabus_row)
            
            current_row = output_ws_syllabus.max_row
    
            # Apply TOP alignment to all primary data cells (A-F)
            for col_idx in range(6): # Columns A (0) through F (5)
                cell = output_ws_syllabus[current_row][col_idx]
                cell.alignment = primary_data_top_alignment
    
            # Apply TOP/WRAP alignment to the syllabus cell (G)
            syllabus_cell = output_ws_syllabus[current_row][6] # Column G is index 6
            syllabus_cell.alignment = syllabus_cell_alignment 
    
    
            # --- Populate Module Assessments sheet ---
            max_assessments = max(num_assessments, num_assessments2)
            rows_to_write_assessments = max(1, max_assessments)

            for i in range(rows_to_write_assessments):
                assessment_row = []
    
                # Columns A-F: Primary Data 
                if i == 0:
                    assessment_row.extend(primary_data_cols)
                else:
                    assessment_row.extend([""] * 6)
    
                # Columns G, H, I, J: Primary Assessment Data
                if i < num_assessments:
                    assessment_row.append(descriptions[i])
                    assessment_row.append(lengths[i])
                    assessment_row.append(semesters[i])
                    assessment_row.append(percentages_assess[i])
                else:
                    assessment_row.extend([""] * 4)
    
                # Column K: Exam Comment
                if i == 0:
                    assessment_row.append(exam_comment)
                else:
                    assessment_row.append("")
    
                # Columns L, M, N, O, P: Secondary Assessment Data
                if i < num_assessments2:
                    assessment_row.append(descriptions2[i])
                    assessment_row.append(semesters2[i])
                    assessment_row.append(when_sets2[i])
                    assessment_row.append(percentages_assess2[i])
                    assessment_row.append(comments4[i])
                else:
                    assessment_row.extend([""] * 5)
    
                # Column Q: Description 5 Assessment 
                if i == 0:
                    assessment_row.append(description5_assessment)
                else:
                    assessment_row.append("")
    
                # Column R: Practical Skills Assessment Comment 
                if i == 0:
                    assessment_row.append(practical_comment)
                else:
                    assessment_row.append("")
    
                output_ws_assessments.append(assessment_row)
    
    
            # --- Populate Module Contributors sheet ---
            rows_to_write_contrib = max(1, num_contributors)
    
            for i in range(rows_to_write_contrib):
                row = []
    
                # Columns A-F: Primary Data
                if i == 0:
                    row.extend(primary_data_cols)
                else:
                    row.extend([""] * 6)
    
                # Columns G, H, I: Contributor Data
                if i < num_contributors:
                    row.append(names[i])
                    row.append(descriptions_contrib[i])
                    row.append(percentages[i])
                else:
                    row.append("")
                    row.append("")
                    row.append("")
    
                output_ws_contrib.append(row)
    except KeyboardInterrupt:
        print("\n\n👋 Program execution stopped by user (Ctrl+C). Incomplete file saved.")
        interrupted = True

    print("\n--- Aggregation complete! ---")

    # 4. Save (ALWAYS runs, but safely handled)
    try:
        if interrupted:
            output_file = OUTPUT_FILE.replace(".xlsx", "_PARTIAL.xlsx")
            print("💾 Data aggregation interrupted by user, saving partial output...")
        else:
            output_file = OUTPUT_FILE
            print("💾 Saving final output...")
    
        output_wb.save(output_file)
        print(f"✅ File saved: {output_file}")
    
    except KeyboardInterrupt:
        print("\n⚠️ Program interrupted while saving output (.xlsx) file. File may be incomplete or corrupted.")
    except Exception as e:
        print(f"❌ Failed to save file: {e}")

if __name__ == '__main__':
    aggregate_module_data()
