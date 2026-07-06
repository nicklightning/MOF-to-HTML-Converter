# -*- coding: utf-8 -*-
"""
Renames .xlsx files using module code found in row 7 (columns W–Z)
"""
from dotenv import load_dotenv
import os
import time
from openpyxl import load_workbook
import sys

load_dotenv('File Addresses.env')
INPUT_FOLDER = os.getenv("XLSX_MOF_FILES_DIR")

def extract_module_code_from_excel(filepath):
    wb = load_workbook(filepath, read_only=True, data_only=True)
    try:
        ws = wb.active

        for col in range(20, 30):  
            value = ws.cell(row=7, column=col).value
            if value and isinstance(value, str) and value.strip():
                return value.strip()

        raise ValueError("No module code found in cells W7–Z7.")

    finally:
        wb.close()  # ✅ CRITICAL FIX
        time.sleep(0.1)  # ✅ CRITICAL: allow OS to release file lock

def rename_excel_files(INPUT_FOLDER):

    if INPUT_FOLDER is None:
        INPUT_FOLDER = os.path.dirname(os.path.abspath(__file__))

    print(f"--- Running in directory: {INPUT_FOLDER} ---")

    if not os.path.isdir(INPUT_FOLDER):
        print(f"Error: Directory '{INPUT_FOLDER}' not found.")
        return

    rename_tasks = []

    # ✅ PHASE 1: Read everything (no renaming yet)
    for filename in os.listdir(INPUT_FOLDER):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            old_filepath = os.path.join(INPUT_FOLDER, filename)

            print(f"\nReading: {filename}")

            try:
                module_code = extract_module_code_from_excel(old_filepath)

                sanitized_code = "".join(
                    c for c in module_code if c.isalnum() or c in ('-', '_')
                )

                if not sanitized_code:
                    raise ValueError("Empty module code after sanitisation")

                new_filename = f"{sanitized_code}.xlsx"
                new_filepath = os.path.join(INPUT_FOLDER, new_filename)

                rename_tasks.append((old_filepath, new_filepath, filename, new_filename))

            except Exception as e:
                print(f"❌ ERROR reading {filename}: {e}")
                continue

    # ✅ IMPORTANT: let OS + OneDrive release locks
    print("\n⏳ Waiting for file locks to clear...")
    time.sleep(2)

    # ✅ PHASE 2: Perform renames
    renamed_count = 0

    for old_filepath, new_filepath, filename, new_filename in rename_tasks:

        # Skip if already correct
        if old_filepath == new_filepath:
            continue

        # Skip duplicates
        if os.path.exists(new_filepath):
            print(f"⚠️ Skipping (exists): {new_filename}")
            continue

        try:
            os.rename(old_filepath, new_filepath)
            print(f"✅ Renamed '{filename}' → '{new_filename}'")
            renamed_count += 1

        except PermissionError:
            print(f"⏭️ Skipped locked file: {filename}")


    # ✅ Add warning if nothing was renamed
    if renamed_count == 0:
        print("⚠️ No files were renamed — all files found are already named after academic modules and/or no other .xlsx files were found.")

    print(f"\n--- Complete: {renamed_count} file(s) renamed ---")

if __name__ == "__main__":
    time.sleep(0.5)
    try:  
        rename_excel_files(INPUT_FOLDER)
        pass
    except KeyboardInterrupt:
        print("\n\n👋 Program execution stopped by user (Ctrl+C). Exiting gracefully.")
        sys.exit(0)
