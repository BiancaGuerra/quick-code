from zipfile import ZipFile
import pandas as pd
import os
import win32com.client

def vba_extract(xlsm_file):
    """Extract the VBA macro from a given XLSM file. 

    Args:
        xlsm file (str): The path to the XLSM file from which to extract the VBA macro.

    Returns:
        str: The name of the file containing the extracted VBA macro (vbaProject.bin).
    """

    vba_filename = 'vbaProject.bin'     # Default name for the extracted VBA macro file (can be changed)

    # Open the XLSM file as a zip archive.
    xlsm_zip = ZipFile(xlsm_file, 'r')

    # Read the xl/vbaProject.bin file from the zip archive.
    vba_data = xlsm_zip.read('xl/' + vba_filename)

    # Write the extracted VBA data to a local file.
    vba_file = open(vba_filename, "wb")
    vba_file.write(vba_data)
    vba_file.close()
    
    return vba_filename

def add_macro(file_path, file_name, xlsm_name, vba_filename):
    """Add a VBA macro to an existing XLSX file and save it as an XLSM file.

    Args:
        file_path (str): The directory path where the XLSX file is located.
        file_name (str): The name of the existing XLSX file (must include .xlsx extension).
        xlsm_name (str): The name of the new file to save with the VBA macro (must include .xlsm extension).
        vba_filename (str): The path to the VBA macro file to be added (vbaProject.bin).
    """

    # Create a Pandas ExcelWriter object to read the existing XLSX file.
    writer = pd.ExcelWriter(os.path.join(file_path, file_name), engine='xlsxwriter')
    
    # Specify the filename for the new XLSM file.
    writer.book.filename = os.path.join(file_path, xlsm_name)
    
    # Add the VBA project to the workbook.
    writer.book.add_vba_project(vba_filename)
    
    writer.close()  # Save and close the workbook


def run_macro(xlsm_name, macro_name):
    """Run a specified VBA macro from a given XLSM file.

    Args:
        xlsm_name (str): The name of the XLSM file containing the VBA macro (must include .xlsm extension).
        macro_name (str): The name of the VBA macro to be executed.
    """

    # Start an instance of the Excel application.
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False  # Keep Excel hidden during execution

    # Open the specified XLSM file.
    workbook = excel.Workbooks.Open(xlsm_name)

    # Execute the specified VBA macro.
    excel.Run(macro_name)

    # Save changes to the workbook and close it.
    workbook.Save()
    workbook.Close()

    # Quit the Excel application.
    excel.Quit()