import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

# Create a tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()  # Hide the root window

# Open a file dialog to select a folder
folder_path = filedialog.askdirectory(
    title="Select Folder Containing Excel Files"
)

file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
source_workbook = pd.ExcelFile(file_path)
sheet_name = 'Sheet1'

source_workbook2 = pd.ExcelFile(
    r"C:\Users\813703\OneDrive - Chulabhorn Royal Academy\Shared Documents\..Drug Master คลังยา\Drug master.xlsx")
dfmaster = source_workbook2.parse("Drug master")
# Get a list of Excel files in the selected folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xls')]

# Initialize an empty DataFrame to store the combined data
stacked_df = pd.DataFrame()

for excel_file in excel_files:
    # Construct the full path to the Excel file
    file_path = os.path.join(folder_path, excel_file)

    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    df['โรงพยาบาลจุฬาภรณ์'] = pd.to_datetime(df['โรงพยาบาลจุฬาภรณ์'], errors='coerce')
    df = df.dropna(subset=['โรงพยาบาลจุฬาภรณ์'])
    df = df.sort_values(by='โรงพยาบาลจุฬาภรณ์').reset_index()
    df.columns = range(10)
    value_to_expand = df.at[0, 1]
    value_to_expand = value_to_expand.replace("รวม", "")
    df[1] = value_to_expand

    df = df[df[4].apply(lambda x: isinstance(x, str))]
    df = df.drop(0, axis=1)
    df[4] = df[4].astype(int)

    negative_values = df[6] < 0
    df.insert(6, '6.5', '')
    df.loc[negative_values, "6.5"] = df.loc[negative_values, 6]
    df.loc[df[6] < 0, 6] = 0
    df['6.5'] = df['6.5'].replace('', 0)

    unit = df.iat[0, 7]
    sum_col6 = df[6].sum()
    sum_col7 = df['6.5'].sum()
    new_row = pd.DataFrame({1: value_to_expand, 5: "รวมทั้งสิ้น", 6: [sum_col6], '6.5': [sum_col7], 7: [unit], 9: ""})
    df = pd.concat([df, new_row], ignore_index=True)
    print(df)

    df.columns = ['ชื่อยาเสพติดให้โทษประเภท 2', 'วัน เดือน ปี', 'AN/VN', 'HN', 'ชื่อ', 'จ่าย', 'รับ', 'หน่วย', 'ราคา', 'ที่อยู่']
    df = df[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'ชื่อ', 'รับ', 'จ่าย', 'หน่วย', 'ที่อยู่']]
    df['จ่ายไป'] = df['ชื่อ'] + " " + df['ที่อยู่']
    df = df[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'จ่ายไป', 'หน่วย', 'รับ', 'หน่วย', 'จ่าย', 'หน่วย']]

    df['วัน เดือน ปี'].fillna('', inplace=True)

    def convert_date(date_str):
        if not pd.isna(date_str):

            date_obj = pd.to_datetime(date_str)

            month_mapping = {
                1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม',
                4: 'เมษายน', 5: 'พฤษภาคม', 6: 'มิถุนายน',
                7: 'กรกฎาคม', 8: 'สิงหาคม', 9: 'กันยายน',
                10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
            }

            # Format the date in Thai format
            thai_day = date_obj.strftime('%d')
            thai_month = month_mapping.get(date_obj.month, date_obj.month)
            thai_year = str(date_obj.year + 543)
            return f"{thai_day} {thai_month} {thai_year}"
        else:
            return ''


    df.insert(3, 'รับจาก อย', '0')
    df.insert(2, 'รหัส', '')

    df['วัน เดือน ปี'] = df['วัน เดือน ปี'].apply(convert_date)
    stacked_df = pd.concat([stacked_df, df], axis=0, ignore_index=True)

dfT = pd.read_excel(source_workbook, sheet_name)

dfmaster = dfmaster[["Material", "TradeName"]]

dfT = pd.merge(dfT, dfmaster, how="left")

dfT = dfT[["Posting Date", "TradeName", "Batch", 'Receiving stor. loc.', "Quantity"]]
dfT.columns = ['วัน เดือน ปี', "ชื่อยาเสพติดให้โทษประเภท 2", 'รหัส', 'จ่ายไป', 'รับจาก อย']

def convert_date(date_str):
    if not pd.isna(date_str):

        date_obj = pd.to_datetime(date_str)

        month_mapping = {
            1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม',
            4: 'เมษายน', 5: 'พฤษภาคม', 6: 'มิถุนายน',
            7: 'กรกฎาคม', 8: 'สิงหาคม', 9: 'กันยายน',
            10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
        }

        # Format the date in Thai format
        thai_day = date_obj.strftime('%d')
        thai_month = month_mapping.get(date_obj.month, date_obj.month)
        thai_year = str(date_obj.year + 543)
        return f"{thai_day} {thai_month} {thai_year}"
    else:
        return ''


dfT['วัน เดือน ปี'] = dfT['วัน เดือน ปี'].apply(convert_date)

dfT.insert(5, 'หน่วย', '')
dfT.insert(6, 'รับ', '')
dfT.insert(7, 'จ่าย', '')

dfT = dfT[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'รหัส', 'จ่ายไป', 'รับจาก อย', 'หน่วย', 'รับ', 'หน่วย', 'จ่าย', 'หน่วย']]

combined_file_path = os.path.join(folder_path, "รายงานการรับเข้าและจ่าย.xlsx",)
Total_df = stacked_df[stacked_df['จ่ายไป'] == "รวมทั้งสิ้น "]

with pd.ExcelWriter(combined_file_path) as writer:
    stacked_df.to_excel(writer, index=False, sheet_name='รายงานแยก')
    Total_df.to_excel(writer, index=False, sheet_name='รายงานรวม')
    dfT.to_excel(writer, index=False, sheet_name='รายงานรับเข้า')
