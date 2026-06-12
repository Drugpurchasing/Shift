import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# Create a GUI window
root = tk.Tk()
root.withdraw()  # Hide the main window

file_path2 = filedialog.askopenfilename(title="ยอดคงคลังสิ้นเดือน", filetypes=[('Excel Files', '*.xlsx')])
file_path = filedialog.askopenfilename(title="Rate", filetypes=[('Excel Files', '*.xls')])


if file_path:
    # Load each sheet into a list of DataFrames
    source_workbook = pd.ExcelFile(file_path)
    source_workbook2 = pd.ExcelFile(r"C:\Users\813703\OneDrive - Chulabhorn Royal Academy\Shared Documents\..Drug Master คลังยา\Drug master.xlsx")
    source_workbook3 = pd.ExcelFile(file_path2)

    sheet_name = "Drug master"
    dfmaster = source_workbook2.parse(sheet_name)

    sheet_name2 = "Sheet1"
    remain = source_workbook3.parse(sheet_name2)

    remain = remain.groupby('Storage location')['Stock Value on Period End'].sum().reset_index()

    remain = remain.rename(columns={'Storage location': 'Store'})

    dfs = []

    for sheet_name in source_workbook.sheet_names:
        df = source_workbook.parse(sheet_name, header=None)

        # Remove the first row from the first sheet
        if sheet_name == source_workbook.sheet_names[0]:
            df = df.iloc[2:]

        dfs.append(df)

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

    merged_df['วันที่จ่ายยา'] = pd.to_datetime(merged_df['วันที่จ่ายยา'], format='%Y-%m-%d')

    merged_df['Month'] = merged_df['วันที่จ่ายยา'].dt.to_period('M')

    merged_df = merged_df[merged_df['Store'] != "อื่นๆ"]

    merged_df["หน่วย"] = pd.to_numeric(merged_df["หน่วย"].str.replace(r'.*/ ', '', regex=True), errors='coerce').fillna(
        1).astype(int)

    merged_df["จำนวน"] = merged_df["จำนวน"] * merged_df["หน่วย"]

    merged_df["ราคาทุนรวม"] = merged_df["จำนวน"] * merged_df["ต้นทุน"] / merged_df["หน่วย"]

    grouped_sumRate_df = merged_df.pivot_table(index=['Material', "Store", 'Material description', 'หน่วยย่อย'], values=['จำนวน', "ราคาทุนรวม", "ราคารวม"], aggfunc='sum').reset_index()

    grouped_Valuesum_df = merged_df.groupby(merged_df.columns[18])[[merged_df.columns[40], merged_df.columns[17]]].sum().reset_index()
    grouped_Valuesum_df.columns = ['Store', 'Sum of Cost price', 'Sum of sale price']
    grouped_Valuesum_df = grouped_Valuesum_df[grouped_Valuesum_df['Store'] != 'อื่นๆ']

    remainFinal = pd.merge(remain, grouped_Valuesum_df, on='Store', how='left')

    remainFinal["วันสำรองคงคลัง"] = remainFinal["Stock Value on Period End"] / remainFinal["Sum of Cost price"] * 30

    source_directory, source_filename = os.path.split(file_path)

    # Create the new file name
    new_filename = os.path.splitext(source_filename)[0] + '_result.xlsx'
    new_filename2 = os.path.splitext(source_filename)[0] + '_Rate.xlsx'

    # Construct the full save path
    save_path = os.path.join(source_directory, new_filename)
    save_path2 = os.path.join(source_directory, new_filename2)

    # Save the filtered DataFrame
    with pd.ExcelWriter(save_path) as writer:
        remainFinal.to_excel(writer, sheet_name='ยอดขาย-คงคลัง-สำรองคงคลัง', index=False)
        grouped_sumRate_df.to_excel(writer, sheet_name='ยอดขาย', index=False)
        merged_df.to_excel(writer, sheet_name='Raw', index=False)
