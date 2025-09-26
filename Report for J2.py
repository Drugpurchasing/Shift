import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import glob

root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()
excel_files = glob.glob(f'{file_path}*.xls')

if file_path:
    # Get a list of all Excel files in the folder
    excel_files = glob.glob(os.path.join(file_path, '*.xls'))
    dfs = []
    for file_path in excel_files:
        source_workbook = pd.ExcelFile(file_path)
        for sheet_name in source_workbook.sheet_names:
            df = source_workbook.parse(sheet_name, header=None)

            # Remove the first row from the first sheet
            if sheet_name == source_workbook.sheet_names[0]:
                df = df.iloc[2:]

            dfs.append(df)

    # Concatenate DataFrames vertically

    try:

        stacked_df = pd.concat(dfs, ignore_index=True)

    except:

        stacked_df = dfs

    print(stacked_df)
    # Ask for a destination path to save the filtered DataFrame
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])

    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')

    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN",
        "HN", "ชื่อ", "อายุ", "สิทธิ์", "แพทย์", "Clinic",
        "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]

    # Set the new column labels
    stacked_df.columns = new_column_labels

    valid_store_values = [2403, 2401, 2408, 2409, 2417, 2402]

    merged_df = stacked_df

    source_directory, source_filename = os.path.split(file_path)

    # Create the new file name
    new_filename = 'J2.xlsx'

    # Construct the full save path
    save_path = os.path.join(source_directory, new_filename)

    valid_store_values = [1400000010, 1400000020, 1400000021, 1400000025, 1400000029, 1400000030, 1400000040, 1400000044, 1400000052,
                          1400000053,1400000055,1400000098,1400000099,1400000148,1400000187,1400000201,1400000220,1400000221,1400000228,
                          1400000247,1400000264,1400000068,1400000069,1400000093,1400000106,1400000113,1400000115,1400000116,1400000118,
                          1400000124,1400000126,1400000130,1400000165,1400000166,1400000167,1400000168,1400000169,1400000170,1400000171,
                          1400000172,1400000194,1400000284,1400000288,1400000294,1400000295,1400000331,1400000335,1400000344,1400000345,
                          1400000265
                          ]

    # valid_store_values = [1200001871]

    merged_df = merged_df[merged_df["Material"].isin(valid_store_values)]

    merged_df = merged_df.loc[:, ['วันที่จ่ายยา', 'VN / AN', 'HN', 'ชื่อ', 'สิทธิ์', "แพทย์", 'Material', 'รายการยา', 'จำนวน']]
    # optional filter

    with pd.ExcelWriter(save_path) as writer:
        merged_df.to_excel(writer, sheet_name='Raw', index=False)
