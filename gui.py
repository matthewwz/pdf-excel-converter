import tkinter as tk
import pandas as pd
from tkinter import messagebox, filedialog
from extractor import process_pdf, update_excel

# Global variable to hold the selected Excel file path
excel_file_path = None

def select_excel_file():
    global excel_file_path
    if excel_file_path is not None:
        messagebox.showwarning("Warning", "An Excel file has already been selected or created!")
        return

    temp_path = filedialog.askopenfilename(
        title='Select Excel file',
        filetypes=[('Excel files', '*.xlsx')],
        defaultextension='.xlsx'
    )
    # Check if a file was actually selected
    if temp_path:  # This will be True if the string is non-empty
        excel_file_path = temp_path
        messagebox.showinfo("Success", f"Excel file selected: {excel_file_path}")
    else:
        messagebox.showinfo("Cancelled", "No file was selected.")

def create_excel_file():
    global excel_file_path
    if excel_file_path is not None:
        messagebox.showwarning("Warning", "An Excel file has already been selected or created!")
        return

    temp_path = filedialog.asksaveasfilename(
        title='Create New Excel File',
        filetypes=[('Excel files', '*.xlsx')],
        defaultextension='.xlsx'
    )
    # Check if a file path was actually provided
    if temp_path:  # This checks if the string is non-empty
        excel_file_path = temp_path
        df = pd.DataFrame(columns=["File", "Date", "Model", "Submitter"])
        df.to_excel(excel_file_path, index=False)
        messagebox.showinfo("Success", f"New Excel file created: {excel_file_path}")
    else:
        messagebox.showinfo("Cancelled", "File creation was cancelled.")


def disable_excel_buttons():
    select_excel_button.config(state=tk.DISABLED)
    create_excel_button.config(state=tk.DISABLED)

def select_pdf_files():
    global excel_file_path
    if excel_file_path is None:
        messagebox.showwarning("Warning", "Please select or create an Excel file first.")
        return

    file_paths = filedialog.askopenfilenames(title='Select PDF files', filetypes=[('PDF files', '*.pdf')])
    for file_path in file_paths:
        if file_path.endswith('.pdf'):
            try:
                new_data = process_pdf(file_path)
                update_excel(new_data, excel_file_path)
                messagebox.showinfo("Success", f"Processed {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {file_path}\nError: {str(e)}")

# Set up the GUI
root = tk.Tk()
root.title("PDF to Excel Processor")
root.geometry("500x250")

select_excel_button = tk.Button(root, text="Select Excel File", command=select_excel_file)
select_excel_button.pack()

create_excel_button = tk.Button(root, text="Create New Excel File", command=create_excel_file)
create_excel_button.pack()

select_pdf_button = tk.Button(root, text="Select PDF Files", command=select_pdf_files)
select_pdf_button.pack()

label = tk.Label(root, text="1. Select or Create Excel File\n2. Select PDF files to process", pady=20, padx=20)
label.pack(expand=True, fill=tk.BOTH)

# Run the GUI
root.mainloop()
