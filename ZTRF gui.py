import time
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
from openpyxl import load_workbook
# Create a tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()  # Hide the root window

# Open a file dialog to select the Excel file
file_path = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel Files", "*.xlsx")]
)

# Check if a fie was selected
if file_path:
    # Use Pandas to read the Excel file into a DataFrame
    sloc = input("ใส่รหัสคลังของท่าน: ")
    df = pd.read_excel(file_path)
    filterdf = df[df['Material Document'].between(7100000000, 7200000000, inclusive='both')].reset_index(drop=True)
    checkerdf = df[["Goods Receipt/Issue Slip", "Reference", "Material Document"]]
    checkerdf = checkerdf[checkerdf["Reference"] > 0]
    checkerdf = checkerdf.rename(columns={"Reference": "Material Document", "Material Document": "Reference"})
    print(checkerdf)
    df = df.groupby(["Goods Receipt/Issue Slip", "Material", "Material description", "Batch"])["Quantity"].sum().reset_index()
    df = df[df['Quantity'] < 0].reset_index(drop=True)
    df = pd.merge(df, filterdf, how='inner').drop(columns=['Reference', 'Movement type', 'Plant'])
    df = pd.merge(df, checkerdf, how='left').drop_duplicates()
    df = df.rename(columns={"Goods Receipt/Issue Slip": "Reservation", "Material Document": "Mat Doc", "Storage location": "คลังจ่าย", "Receiving stor. loc.": "คลังรับ"})
    df = df.rename(columns={"Base Unit of Measure": "หน่วย", "Quantity": "จำนวน"})
    df = df.reset_index(drop=True)
    df['จำนวน'] = df['จำนวน'].abs()
    df = df[["Mat Doc", "Posting Date", "Reservation", 'Material', 'Material description', 'Batch', "จำนวน", 'คลังจ่าย', 'คลังรับ']]
    df['Posting Date'] = df['Posting Date'].dt.strftime('%d.%m.%Y')

    if sloc == 'All':
        time.sleep(0.1)
    else:
        df = df[df["คลังรับ"] == int(sloc)]

    source_directory, source_filename = os.path.split(file_path)
    new_filename = sloc + "_ZTRF.xlsx"
    save_path = os.path.join(source_directory, new_filename)
    df.to_excel(save_path, index=False)

    os.system(f'start {save_path}')

else:
    print("No file selected")

# Close the tkinter root window (optional)
root.destroy()
