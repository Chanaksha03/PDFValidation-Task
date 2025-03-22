import os
import pandas as pd
import subprocess
from PyPDF2 import PdfReader

def validate_pdf(pdf_path, file_name, cma_code, factsheet_code):
    # Full path to the PDF file
    if os.access(pdf_path, os.R_OK):
        print(f"PDF {file_name} is readable")
    else:
        print(f"Permission denied or file not found for {file_name}")
        return {"Error": f"PDF '{file_name}' not readable."}

    try:
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)

        # Initialize results dictionary
        results = {
            "keywords_found": set(),
            "CMA Form Code Found": False,
            "Factsheet Form Code Found": False,
            "Total Pages": total_pages
        }

        # Extract text and search for the form codes
        for page_num in range(total_pages):
            page = reader.pages[page_num]
            text = page.extract_text()  # Extract text from each page

            if text:
                # Check for CMAFormCode and FactsheetFormCode in the text
                if cma_code in text:
                    results["CMA Form Code Found"] = True
                if factsheet_code in text:
                    results["Factsheet Form Code Found"] = True

        return results

    except Exception as e:
        return {"Error": str(e)}

def fetch_and_update_pdf_results(excel_file, pdf_column_name, cma_column_name, factsheet_column_name, result_column_name, language_column_name, expected_language, search_directory, starting_row, ending_row):
    try:
        # Load the full Excel sheet
        df = pd.read_excel(excel_file)
        
        # Strip spaces from column names to handle any extra spaces
        df.columns = df.columns.str.strip()
        
        # Ensure the necessary columns exist
        required_columns = [pdf_column_name, cma_column_name, factsheet_column_name, result_column_name, language_column_name]
        for col in required_columns:
            if col not in df.columns:
                print(f"Column '{col}' not found in the Excel sheet.")
                return

        # Build a map of filenames to their paths
        file_map = {}
        for root, dirs, files in os.walk(search_directory):
            for file in files:
                file_map[file] = os.path.join(root, file)

        # Iterate through the rows and PDFs in the specified range
        for index in range(starting_row, ending_row):
            row = df.iloc[index]
            pdf_name = str(row[pdf_column_name]).strip()
            cma_code = str(row[cma_column_name]).strip()
            factsheet_code = str(row[factsheet_column_name]).strip()
            lang_value = str(row[language_column_name]).strip()

            if pd.isna(pdf_name) or pdf_name not in file_map:
                print(f"Skipping row {index + 1} - PDF '{pdf_name}' not found or empty.")
                continue

            # Check if the language is correct
            if lang_value != expected_language:
                df.at[index, result_column_name] = "fail"
                print(f"Row {index + 1} failed language validation. Expected: {expected_language}, Found: {lang_value}")
                continue  # Skip further processing if language is incorrect

            pdf_path = file_map[pdf_name]
            print(f"Validating PDF for row {index + 1}: {pdf_path}")

            # Validate the PDF
            validation_result = validate_pdf(pdf_path, pdf_name, cma_code, factsheet_code)

            # Determine pass or fail based on validation (CMA and Factsheet Form codes must be found)
            is_pass = validation_result["CMA Form Code Found"] and validation_result["Factsheet Form Code Found"]
            df.at[index, result_column_name] = "pass" if is_pass else "fail"

            # Open the PDF (this will open in the default viewer)
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(pdf_path)
                else:  # macOS/Linux
                    subprocess.run(['open', pdf_path])
            except Exception as e:
                print(f"Failed to open {pdf_path}: {e}")

        # After processing the specified rows, save the entire DataFrame (not just a part of it) back to Excel
        df.to_excel(excel_file, index=False)
        print("Excel file updated successfully with pass/fail results.")

    except Exception as e:
        print(f"An error occurred: {e}")


# Example usage
excel_file = r'C:\Users\Chanaksha Pawar\OneDrive\Desktop\PR-Validation-JAN2025\Compare_Sheet.xlsx'
pdf_column_name = 'CombinedFormCode_Pdf_version'
cma_column_name = 'CMAFormCode'
factsheet_column_name = 'FactsheetFormCode'
result_column_name = 'Result'  # Column where results will be updated (pass/fail)
language_column_name = 'LANG'  # Column containing the language
expected_language = 'EN' or 'SP'  # The language to validate
search_directory = r'C:\Users\Chanaksha Pawar\OneDrive\Desktop\PR-Validation-JAN2025\Requirements\oil'

# Define the starting and ending rows (index-based)
starting_row = 0  # Starting from row 1 (Excel row 2)
ending_row = 4    # Ending at row 4 (Excel row 5)

fetch_and_update_pdf_results(excel_file, pdf_column_name, cma_column_name, factsheet_column_name, result_column_name, language_column_name, expected_language, search_directory, starting_row, ending_row)
