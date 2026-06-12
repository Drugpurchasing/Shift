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
    source_workbook2 = pd.ExcelFile(r"C:\Users\bungs\OneDrive - Chulabhorn Royal Academy\Shared Documents\..Drug Master คลังยา\Drug master.xlsx")

    sheet_name = "Drug master"
    dfmaster = source_workbook2.parse(sheet_name)

    # Concatenate DataFrames vertically
    stacked_df = pd.concat(dfs, ignore_index=True)

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

    merged_df = pd.merge(stacked_df, dfmaster, on="Material", how="left")

    merged_df.loc[~merged_df["Store"].isin(valid_store_values), "Store"] = "อื่นๆ"

    merged_df["ราคาทุนรวม"] = merged_df["จำนวน"] * merged_df["ต้นทุน"]

    merged_df['วันที่จ่ายยา'] = pd.to_datetime(merged_df['วันที่จ่ายยา'], format='%Y-%m-%d')

    merged_df['Month'] = merged_df['วันที่จ่ายยา'].dt.to_period('M')

    merged_df = merged_df[merged_df['Store'] != "อื่นๆ"]

    grouped_sum_df = merged_df.groupby([merged_df.columns[12], merged_df.columns[13], merged_df.columns[18], merged_df.columns[29]])[[merged_df.columns[14], merged_df.columns[39], merged_df.columns[17]]].sum().reset_index()
    grouped_sum_df.columns = ['Material', 'Material description', 'Store', 'Unit', "Quantity", 'Sum of Cost price', "Sum of Sale price"]
    grouped_sum_df = grouped_sum_df[grouped_sum_df['Store'] != "อื่นๆ"]

    merged_df["หน่วย"] = pd.to_numeric(merged_df["หน่วย"].str.replace(r'.*/ ', '', regex=True), errors='coerce').fillna(1).astype(int)

    merged_df["จำนวน"] = merged_df["จำนวน"] * merged_df["หน่วย"]

    grouped_sumRate_df = merged_df.pivot_table(index=['Material', 'Material description', 'หน่วยย่อย'], columns='Month', values='จำนวน', aggfunc='sum').reset_index()

    grouped_sumRateSplit_df = merged_df.pivot_table(index=['Material', "Store", 'Material description', 'หน่วยย่อย'], columns='Month', values='จำนวน', aggfunc='sum').reset_index()

    source_directory, source_filename = os.path.split(file_path)

    # Create the new file name
    new_filename = 'Drugstore_Rate.xlsx'

    # Construct the full save path
    save_path = os.path.join(source_directory, new_filename)

    # Group by the desired column and sum the other columns
    grouped_Valuesum_df = merged_df.groupby(merged_df.columns[18])[[merged_df.columns[39], merged_df.columns[17]]].sum().reset_index()

    grouped_Valuesum_df.columns = ['Sloc', 'Sum of Cost price', 'Sum of sale price']
    grouped_Valuesum_df = grouped_Valuesum_df[grouped_Valuesum_df['Sloc'] != 'อื่นๆ']

    # valid_store_values = [1200001398, 1200001399, 1200001404, 1200001405, 1200001406, 1200001407, 1200001408]

    # merged_df = merged_df[merged_df["Material"].isin(valid_store_values)]
    # optional filter

    with pd.ExcelWriter(save_path) as writer:
        grouped_sumRate_df.to_excel(writer, sheet_name='Rate แยกเดือน', index=False)
        grouped_sumRateSplit_df.to_excel(writer, sheet_name='Rate แยกเดือน แยกห้อง', index=False)
        merged_df.to_excel(writer, sheet_name='Raw', index=False)
