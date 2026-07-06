# -*- coding: utf-8 -*-
"""
Created on Mon Nov 24 15:50:20 2025

@author: nnrw
"""
from dotenv import load_dotenv
import sys
import os
import re
import time

load_dotenv('File Addresses.env')
INPUT_FOLDER = os.getenv("XML_MOF_FILES_DIR")

def rename_xml_files(INPUT_FOLDER):
    """
    Scans a directory for .xml files, extracts a module code (e.g., NES1000) 
    using a regular expression, and renames the file to {module_code}.xml.

    Args:
        INPUT_FOLDER (str): The path to the folder containing the XML files.
    """
    print(f"--- Starting XML Renamer in directory: {os.path.abspath(INPUT_FOLDER)} ---")

    # 1. Define the regular expression pattern
    # This pattern looks for the literal string 'Code="' followed by one or more 
    # characters that are NOT a double quote ('"'), and captures those characters (the module code).
    # This is a very robust way to extract the content inside the quotes.
    module_code_pattern = re.compile(r'Code="([^"]+)"')
    
    # Counter for successful renames
    renamed_count = 0
    
    # 2. Check if the target directory exists
    if not os.path.isdir(INPUT_FOLDER):
        print(f"Error: Directory '{INPUT_FOLDER}' not found. Please create it and place your .xml files inside.")
        return

    # 3. Iterate over all items in the target directory
    for filename in os.listdir(INPUT_FOLDER):
        if filename.endswith(".xml"):
            old_filepath = os.path.join(INPUT_FOLDER, filename)
            
            print(f"\nProcessing: {filename}")
            
            try:
                # 4. Read the file content
                with open(old_filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 5. Search for the pattern
                match = module_code_pattern.search(content)
                
                if match:
                    # 6. Extract the module code from the captured group
                    module_code = match.group(1)
                    
                    # Basic validation and sanitization of the extracted code
                    if not module_code:
                        print(f"Skipping: Found an empty code string in {filename}. Content inside Code=\"\" is empty.")
                        continue
                        
                    # Remove any characters that might cause file system issues (optional but safer)
                    sanitized_code = "".join(c for c in module_code if c.isalnum() or c in ('-', '_'))
                    
                    new_filename = f"{sanitized_code}.xml"
                    new_filepath = os.path.join(INPUT_FOLDER, new_filename)
                    
                    # 7. Check if the new name is different before renaming
                    if old_filepath == new_filepath:
                        print(f"Skipping: File is already named correctly as '{new_filename}'.")
                        continue

                    # 8. Perform the rename operation
                    os.rename(old_filepath, new_filepath)
                    print(f"Success: Renamed '{filename}' to '{new_filename}'")
                    renamed_count += 1
                    
                else:
                    print(f"Warning: Module code (Code=\"...\") not found in {filename}.")
            
            except FileNotFoundError:
                # This should ideally not happen since we're iterating os.listdir()
                print(f"Error: File not found during processing: {old_filepath}")
            except Exception as e:
                print(f"An unexpected error occurred while processing {filename}: {e}")

    print(f"\n--- Renaming Complete. {renamed_count} file(s) were successfully renamed. ---")

# Execute the function
if __name__ == "__main__":
    # Wait a moment before running to ensure any necessary setup (like creating the MOFs folder)
    # can be done if this script is run immediately after creation.
    time.sleep(0.5) 
    try:  
        rename_xml_files(INPUT_FOLDER)
        pass
    except KeyboardInterrupt:
        print("\n\n👋 Program execution stopped by user (Ctrl+C). Exiting gracefully.")
        sys.exit(0)