import psutil
import os
from io import BytesIO
import openpyxl 
from openpyxl import load_workbook
import pandas as pd
import formulas
import messagebox   
#Check if the Excel File Exists and Close if Open
def close_open_excel():
    # Iterate through all running processes
    for proc in psutil.process_iter(['name']):
        try:
            # Check if process name contains 'EXCEL'
            if "EXCEL.EXE" in proc.info['name'].upper():
                proc.terminate()
                print("Excel process found and closed.")
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    print("Excel was not running.")
    

# Usage
#close_open_excel()

# Check Excel exists

def check_excel_exists(file_path):
    print("Checking if Excel file exists at:", file_path)
    print ("Hi123")
   # file_path = os.path.expandvars(os.path.expanduser(str(file_path)).strip().strip('"\''))

    abs_path = os.path.abspath(file_path)
    print("Resolved absolute path:", abs_path)

    parent = os.path.dirname(abs_path) or "."
    try:
        print("Parent directory listing (for debugging):", os.listdir(parent))
    except Exception as e:
        print(f"Could not list parent directory {parent}: {e}")

    if os.path.exists(abs_path):
        if os.path.isfile(abs_path):
            print(f"Success: File found at {abs_path}")
            return True
        else:
            print(f"Path exists but is not a file: {abs_path}")
            return False
    else:
        print("Error: File does not exist.")
        return False

# Usage
#path = r"C:\Users\velur\Desktop\Selenium-Python\Emp2Term.xlsx"
#check_excel_exists(path)

#Download Excel File

def workbook_to_bytes(workbook):
    """
    Convert an existing openpyxl workbook to bytes.   
    Args:
        workbook: openpyxl.Workbook object       
    Returns:
        bytes: Excel file as bytes
    """
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()

def download_excel(file_path):
    print("Excel issue")
    wb = openpyxl.load_workbook(file_path)
    excel_bytes = workbook_to_bytes(wb)
 
