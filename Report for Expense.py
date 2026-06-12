import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import glob

root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()
excel_files = glob.glob(f'{file_path}*.xlsx')

if file_path:
    # Get a list of all Excel files in the folder
    excel_files = glob.glob(os.path.join(file_path, '*.xlsx'))
    dfs = []
    for file_path in excel_files:
        source_workbook = pd.ExcelFile(file_path)
        for sheet_name in source_workbook.sheet_names:
            df = source_workbook.parse(sheet_name, header=None)

            # Remove the first row from the first sheet
            if sheet_name == source_workbook.sheet_names[0]:
                df = df.iloc[2:]

            dfs.append(df)
            dfs = pd.concat(dfs, ignore_index=True)

    print(dfs)

    source_workbook2 = pd.ExcelFile(r"C:\Users\813703\OneDrive - Chulabhorn Royal Academy\Shared Documents\..Drug Master คลังยา\Drug master.xlsx")

    sheet_name = "Drug master"
    dfmaster = source_workbook2.parse(sheet_name)

    dfs = dfs.iloc[:, [0, 2, 3, 9, 8]]

    dfs.columns = ['Month', 'Material', 'Material Name', "Unit", 'Quantity']

    dfs['Material'] = dfs['Material'].astype(int)

    dfs = pd.merge(dfs, dfmaster, on="Material", how="left")

    dfs[['Quantity']] = dfs[['Quantity']].abs()

    dfs['Month'] = pd.to_datetime(dfs['Month'], format='%Y-%m-%d')

    dfs['Month'] = dfs['Month'].dt.to_period('M')

    dfs = dfs.pivot_table(index=['Material', 'Material description', "หน่วยย่อย"], columns="Month", values='Quantity', aggfunc='sum').reset_index()

    source_directory, source_filename = os.path.split(file_path)

    # Create the new file name
    new_filename = os.path.splitext(source_filename)[0] + '_Rate.xlsx'

    # Construct the full save path
    save_path = os.path.join(source_directory, new_filename)

    with pd.ExcelWriter(save_path) as writer:
        dfs.to_excel(writer, sheet_name='Rate แยกเดือน', index=False)
