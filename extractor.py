import pdfplumber
import pandas as pd
import re
import os

def process_pdf(pdf_path):
    """Extracts data from the first page of a PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        if text is None:
            raise ValueError("No text found on the first page.")
        
        
        # Using search to find the first occurrence of each pattern in the extracted text
        date_search = re.search(r"M.D.Y\s+/\s+M.J.A:\s+([0-9/]+)", text)
        model_search = re.search(r"Model\s+/\s+Modèle:\s+(\S+)", text)
        submitter_search = re.search(r"Submitter\s+/\s+Réquérant:\s+([^\n]+?)\s+Page", text)
        report_num_search = re.search(r"Report No[./]*\s+No\.\s+Rapport:\s*(\S+)", text)

        # Extracting data if the pattern was found, otherwise None
        date = date_search.group(1) if date_search else None
        model = model_search.group(1) if model_search else None
        submitter = submitter_search.group(1).strip() if submitter_search else None
        report_num = report_num_search.group(1) if report_num_search else None
        
        return {
            "File": pdf_path,
            "Date": date,
            "Submitter": submitter,
            "Model": model,
            "Report No.": report_num
        }

def update_excel(new_data, excel_path):
    """Updates or creates an Excel file with new data."""
    # Load existing data or initialize new DataFrame
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
    else:
        df = pd.DataFrame(columns=["File", "Date", "Submitter", "Model", "Report No."])
    
    # Convert new_data dictionary to DataFrame
    new_data_df = pd.DataFrame([new_data])

    # Append new data using concat
    df = pd.concat([df, new_data_df], ignore_index=True)
    df.to_excel(excel_path, index=False)

if __name__ == "__main__":
    # Define the path to the output Excel file for testing
    excel_path = '/Users/matthew/PDFtoExcel/output.xlsx'

    # List of paths to the PDF files for testing
    pdf_paths = [
        '/Users/matthew/PDFtoExcel/fileTest1.pdf',
        '/Users/matthew/PDFtoExcel/fileTest2.pdf',
        '/Users/matthew/PDFtoExcel/fileTest3.pdf'
    ]

    # Process each PDF and update the Excel file
    for pdf_path in pdf_paths:
        new_data = process_pdf(pdf_path)
        update_excel(new_data, excel_path)

    print("Data extraction and update completed.")
