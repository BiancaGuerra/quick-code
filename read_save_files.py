import os
import pandas as pd

def read_all_sheets(file_path: str, file_name: str) -> pd.DataFrame:
    """Read an Excel file with multiple sheets and concatenate all sheets into a single DataFrame.

    Args:
        file_path (str): The directory path where the Excel file is located.
        file_name (str): The name of the Excel file (including its extension).

    Returns:
        pd.DataFrame: A DataFrame containing all sheets concatenated into a single table.
    """
    dataframes = []  # List to store DataFrames from each sheet
    all_sheets = pd.read_excel(os.path.join(file_path, file_name), sheet_name=None, header=6)  # Read the file with multiple sheets
    
    # Iterate through all sheets
    for sheet, temp_df in all_sheets.items():
        dataframes.append(temp_df)  # Append each sheet's DataFrame to the list

    if not dataframes:
        raise ValueError("No sheets found in the Excel file.")

    df = pd.concat(dataframes, ignore_index=True)  # Concatenate all DataFrames into one

    return df

def read_save_files(csv_file: str, excel_file: str, new_file: str) -> None:
    """Read a CSV file and an Excel file, merge them, and save the results to a new Excel file.

    Args:
        csv_file (str): Path to the CSV file, including its name and extension.
        excel_file (str): Path to the Excel file, including its name and extension.
        new_file (str): Path where the new Excel file will be saved, including its name and extension.
    """
    # Read the files
    df_csv = pd.read_csv(csv_file, sep=';', encoding='latin-1')
    df_excel = pd.read_excel(excel_file, engine='openpyxl')

    # Fill in with the name of the columns
    key_column_excel = ''       # Specify the key column in the Excel file
    key_column_csv = ''     # Specify the key column in the CSV file
    column_year = ''        # Specify the column that contains year information

    # Merge the two dataframes
    df_excel = df_excel.merge(df_csv, how='left', left_on=key_column_excel, right_on=key_column_csv)

    # Create new dataframes by filtering the year
    df1 = df_excel[df_excel[column_year] == 2021]
    df2 = df_excel[df_excel[column_year] == 2022]
    df3 = df_excel[df_excel[column_year] == 2023]
    df4 = df_excel[df_excel[column_year] == 2024]

    # Write a new file with 4 different sheets
    with pd.ExcelWriter(new_file, engine='xlsxwriter', mode='w') as writer:
        df1.to_excel(writer, sheet_name='2021', index=False)
        df2.to_excel(writer, sheet_name='2022', index=False)
        df3.to_excel(writer, sheet_name='2023', index=False)
        df4.to_excel(writer, sheet_name='2024', index=False)
