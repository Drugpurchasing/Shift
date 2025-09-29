import streamlit as st
import pandas as pd
import io
import numpy as np
from PyPDF2 import PdfMerger
from openpyxl.styles import Alignment, Font  # เพิ่ม Library สำหรับฟังก์ชันใหม่


# ==============================================================================
# FUNCTION 1-5 (โค้ดฟังก์ชันเดิม 5 ฟังก์ชัน ไม่มีการเปลี่ยนแปลง)
# ==============================================================================
def process_j2_report(uploaded_files):
    # (โค้ดฟังก์ชันนี้เหมือนเดิมจากครั้งก่อน)
    dfs = []
    for file_obj in uploaded_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0: df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}"); return None
    if not dfs: st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด"); return None
    stacked_df = pd.concat(dfs, ignore_index=True)
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    new_column_labels = ["ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", "สิทธิ์",
                         "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย", "ราคาขายR", "ราคารวม",
                         "Store"]
    if len(stacked_df.columns) != len(new_column_labels): st.error(
        f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}"); return None
    stacked_df.columns = new_column_labels
    valid_material_values = [1400000010, 1400000020, 1400000021, 1400000025, 1400000029, 1400000030, 1400000040,
                             1400000044, 1400000052, 1400000053, 1400000055, 1400000098, 1400000099, 1400000148,
                             1400000187, 1400000201, 1400000220, 1400000221, 1400000228, 1400000247, 1400000264,
                             1400000068, 1400000069, 1400000093, 1400000106, 1400000113, 1400000115, 1400000116,
                             1400000118, 1400000124, 1400000126, 1400000130, 1400000165, 1400000166, 1400000167,
                             1400000168, 1400000169, 1400000170, 1400000171, 1400000172, 1400000194, 1400000284,
                             1400000288, 1400000294, 1400000295, 1400000331, 1400000335, 1400000344, 1400000345,
                             1400000265]
    merged_df = stacked_df[stacked_df["Material"].isin(valid_material_values)].copy()
    final_cols = ['วันที่จ่ายยา', 'VN / AN', 'HN', 'ชื่อ', 'สิทธิ์', "แพทย์", 'Material', 'รายการยา', 'จำนวน']
    merged_df = merged_df[final_cols]
    return merged_df


def process_drug_rate_analysis(data_files, master_file):
    dfs = [];
    for file_obj in data_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0: df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}"); return None, {}
    if not dfs: st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด"); return None, {}
    stacked_df = pd.concat(dfs, ignore_index=True)
    try:
        dfmaster = pd.read_excel(master_file, sheet_name="Drug master")
    except Exception as e:
        st.error(f"ไม่สามารถอ่านชีท 'Drug master' จากไฟล์ Master ได้: {e}"); return None, {}
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    new_column_labels = ["ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", "สิทธิ์",
                         "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย", "ราคาขายR", "ราคารวม",
                         "Store"]
    if len(stacked_df.columns) != len(new_column_labels): st.error(
        f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}"); return None, {}
    stacked_df.columns = new_column_labels;
    merged_df = pd.merge(stacked_df, dfmaster, on="Material", how="left");
    merged_df['Store'] = merged_df['Store'].astype('object')
    valid_store_values = [2403, 2401, 2408, 2409, 2417, 2402];
    merged_df.loc[~merged_df["Store"].isin(valid_store_values), "Store"] = "อื่นๆ"
    merged_df["ราคาทุนรวม"] = pd.to_numeric(merged_df["จำนวน"], errors='coerce') * pd.to_numeric(merged_df["ต้นทุน"],
                                                                                                 errors='coerce')
    merged_df['วันที่จ่ายยา'] = pd.to_datetime(merged_df['วันที่จ่ายยา'], errors='coerce');
    merged_df['Month'] = merged_df['วันที่จ่ายยา'].dt.to_period('M')
    merged_df = merged_df[merged_df['Store'] != "อื่นๆ"]
    direct_map = {
        '(ตรวจที่รพ.จุฬาภรณ์) โครงการคัดกรองมะเร็งปากมดลูก ณ รพ.จุฬาภรณ์และคณะแพทย์ศาสตร์วชิรพยาบาล': 'จ่ายเอง',
        '[TopUp] สวัสดิการเจ้าหน้าที่ราชวิทยาลัยจุฬาภรณ์': 'สวัสดิการเจ้าหน้าที่',
        'องค์การปกครองส่วนท้องถิ่นบำนาญ(เบิกจ่ายตรง)': 'ข้าราชการ'}
    merged_df["สิทธิ์"] = merged_df["สิทธิ์"].map(direct_map).fillna(merged_df["สิทธิ์"])
    opd_merged_df = merged_df[
        merged_df['เลขที่เอกสาร'].isna() | (merged_df['เลขที่เอกสาร'].astype(str).str.strip().isin(['', '0']))]
    opd_2409 = opd_merged_df[opd_merged_df['Store'].notna() & (opd_merged_df['Store'] == 2409)];
    opd_not_2409 = opd_merged_df[opd_merged_df['Store'].notna() & (opd_merged_df['Store'] != 2409)]
    ipd_merged_df = merged_df[
        merged_df['Clinic'].isna() | (merged_df['Clinic'].astype(str).str.strip().isin(['', '0']))]
    ipd_2409 = ipd_merged_df[ipd_merged_df['Store'].notna() & (ipd_merged_df['Store'] == 2409)];
    ipd_not_2409 = ipd_merged_df[ipd_merged_df['Store'].notna() & (ipd_merged_df['Store'] != 2409)]

    def count_unique_by_month(df, subset_cols):
        return df.drop_duplicates(subset=subset_cols).groupby('Month').size().reset_index(name='Unique_Count')

    uniqueOPD = count_unique_by_month(opd_not_2409, ['VN / AN', 'HN', 'Clinic', 'Month']);
    uniqueOPD2409 = count_unique_by_month(opd_2409, ['VN / AN', 'HN', 'Clinic', 'Month'])
    uniqueIPD = count_unique_by_month(ipd_not_2409, ['เลขที่เอกสาร', 'HN', 'Ward', 'Month']);
    uniqueIPD2409 = count_unique_by_month(ipd_2409, ['เลขที่เอกสาร', 'HN', 'Ward', 'Month'])
    merged_df["หน่วย"] = pd.to_numeric(merged_df["หน่วย"].astype(str).str.replace(r'.*/ ', '', regex=True),
                                       errors='coerce').fillna(1).astype(int)
    merged_df["จำนวน"] = merged_df["จำนวน"] * merged_df["หน่วย"];
    merged_df['HN'] = merged_df['HN'].astype(str).str.replace('.0', '', regex=False)
    grouped_countHN_df = merged_df.pivot_table(index=['Material', 'Material description'], columns='Month', values='HN',
                                               aggfunc=pd.Series.nunique).reset_index()
    grouped_sumRate_df = merged_df.pivot_table(index=['Material', 'Material description', 'หน่วยย่อย'], columns='Month',
                                               values='จำนวน', aggfunc='sum').reset_index()
    grouped_sumRateSplit_df = merged_df.pivot_table(index=['Material', "Store", 'Material description', 'หน่วยย่อย'],
                                                    columns='Month', values='จำนวน', aggfunc='sum').reset_index()
    output_dfs = {"Rate แยกเดือน": grouped_sumRate_df, "Rate (M-Sloc)": grouped_sumRateSplit_df,
                  "จำนวนเคสต่อเดือน": grouped_countHN_df, "Raw": merged_df,
                  "Summary_Data": {'จำนวนใบยา OPD': uniqueOPD, 'จำนวนใบยา OPD 2409': uniqueOPD2409,
                                   'จำนวนใบยา IPD': uniqueIPD, 'จำนวนใบยา IPD 2409': uniqueIPD2409, }}
    return merged_df, output_dfs


def process_epi_usage(uploaded_files):
    dfs = [];
    for file_obj in uploaded_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0: df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}"); return None
    if not dfs: st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด"); return None
    stacked_df = pd.concat(dfs, ignore_index=True)
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    new_column_labels = ["ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", "สิทธิ์",
                         "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย", "ราคาขายR", "ราคารวม",
                         "Store"]
    if len(stacked_df.columns) != len(new_column_labels): st.error(
        f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}"); return None
    stacked_df.columns = new_column_labels
    valid_epi_materials = [1400000084, 1400000083, 1400000087, 1400000086, 1400000088, 1400000081, 1400000082,
                           1400000090, 1400000085, 1400000089]
    epi_df = stacked_df[stacked_df["Material"].isin(valid_epi_materials)].copy()
    summary_df = epi_df.groupby(['Material', 'รายการยา'])['จำนวน'].sum().reset_index()
    summary_df.rename(columns={'จำนวน': 'จำนวนรวม'}, inplace=True)
    return summary_df


def process_narcotics_report(xls_files, receipt_report_file, master_file):
    def convert_date_to_thai(date_str):
        if not pd.isna(date_str):
            try:
                date_obj = pd.to_datetime(date_str);
                month_mapping = {1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน', 5: 'พฤษภาคม', 6: 'มิถุนายน',
                                 7: 'กรกฎาคม', 8: 'สิงหาคม', 9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'}
                return f"{date_obj.strftime('%d')} {month_mapping.get(date_obj.month, date_obj.month)} {str(date_obj.year + 543)}"
            except (ValueError, TypeError):
                return ''
        return ''

    stacked_df_list = []
    for file_obj in xls_files:
        try:
            df = pd.read_excel(file_obj)
            df['โรงพยาบาลจุฬาภรณ์'] = pd.to_datetime(df['โรงพยาบาลจุฬาภรณ์'], errors='coerce')
            df = df.dropna(subset=['โรงพยาบาลจุฬาภรณ์']).sort_values(by='โรงพยาบาลจุฬาภรณ์').reset_index(drop=True)
            df.columns = range(df.shape[1]);
            value_to_expand = str(df.at[0, 1]).replace("รวม", "").strip();
            df[1] = value_to_expand
            df = df[df[4].apply(lambda x: isinstance(x, str) and x.strip() != '')]
            df[4] = pd.to_numeric(df[4], errors='coerce').dropna().astype(int);
            df = df.drop(0, axis=1)
            negative_values = df[6] < 0;
            df.insert(6, '6.5', 0);
            df.loc[negative_values, '6.5'] = df.loc[negative_values, 6];
            df.loc[df[6] < 0, 6] = 0
            new_row = pd.DataFrame(
                {1: [value_to_expand], 5: ["รวมทั้งสิ้น"], 6: [df[6].sum()], '6.5': [df['6.5'].sum()],
                 7: [df.iat[0, 7]], 9: [""]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.columns = ['ชื่อยาเสพติดให้โทษประเภท 2', 'วัน เดือน ปี', 'AN/VN', 'HN', 'ชื่อ', 'จ่าย', 'รับ', 'หน่วย',
                          'ราคา', 'ที่อยู่']
            df = df[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'ชื่อ', 'รับ', 'จ่าย', 'หน่วย', 'ที่อยู่']]
            df['จ่ายไป'] = df['ชื่อ'].astype(str) + " " + df['ที่อยู่'].astype(str)
            df = df[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'จ่ายไป', 'หน่วย', 'รับ', 'หน่วย', 'จ่าย', 'หน่วย']]
            df['วัน เดือน ปี'] = df['วัน เดือน ปี'].apply(convert_date_to_thai);
            df.insert(3, 'รับจาก อย', '');
            df.insert(2, 'รหัส', '')
            df.columns = ['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'รหัส', 'จ่ายไป', 'หน่วย1', 'รับ', 'หน่วย2',
                          'จ่าย', 'หน่วย3', 'รับจาก อย']
            df = df[
                ['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'รหัส', 'รับจาก อย', 'จ่ายไป', 'หน่วย1', 'รับ', 'หน่วย2',
                 'จ่าย', 'หน่วย3']]
            stacked_df_list.append(df)
        except Exception as e:
            st.warning(f"ไม่สามารถประมวลผลไฟล์ {file_obj.name}: {e}"); continue
    if not stacked_df_list: st.error("ไม่สามารถประมวลผลไฟล์ข้อมูลการจ่ายยาได้เลย"); return None
    stacked_df = pd.concat(stacked_df_list, axis=0, ignore_index=True)
    try:
        dfT = pd.read_excel(receipt_report_file, sheet_name='Sheet1')
        dfmaster = pd.read_excel(master_file, sheet_name="Drug master")[["Material", "TradeName"]]
        dfT = pd.merge(dfT, dfmaster, how="left")[
            ["Posting Date", "TradeName", "Batch", 'Receiving stor. loc.', "Quantity"]]
        dfT.columns = ['วัน เดือน ปี', "ชื่อยาเสพติดให้โทษประเภท 2", 'รหัส', 'จ่ายไป', 'รับจาก อย']
        dfT['วัน เดือน ปี'] = dfT['วัน เดือน ปี'].apply(convert_date_to_thai);
        dfT.insert(5, 'หน่วย', '');
        dfT.insert(6, 'รับ', '');
        dfT.insert(7, 'จ่าย', '')
        dfT = dfT[['วัน เดือน ปี', 'ชื่อยาเสพติดให้โทษประเภท 2', 'รหัส', 'จ่ายไป', 'รับจาก อย', 'หน่วย', 'รับ', 'หน่วย',
                   'จ่าย', 'หน่วย']]
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์รายงานรับเข้า: {e}"); return None
    total_df = stacked_df[stacked_df['จ่ายไป'].str.strip() == "รวมทั้งสิ้น"].copy()
    return {'รายงานแยก': stacked_df, 'รายงานรวม': total_df, 'รายงานรับเข้า': dfT}


def process_kpi_report(rate_file, inventory_file, master_file):
    try:
        remain = pd.read_excel(inventory_file, sheet_name="Sheet1");
        remain = remain.groupby('Storage location')['Stock Value on Period End'].sum().reset_index()
        remain = remain.rename(columns={'Storage location': 'Store'})
        source_workbook = pd.ExcelFile(rate_file);
        dfs = [source_workbook.parse(sheet_name, header=None) for sheet_name in source_workbook.sheet_names]
        dfs[0] = dfs[0].iloc[2:];
        stacked_df = pd.concat(dfs, ignore_index=True)
        dfmaster = pd.read_excel(master_file, sheet_name="Drug master")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์เริ่มต้น: {e}"); return None
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce');
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    new_column_labels = ["ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", "สิทธิ์",
                         "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย", "ราคาขายR", "ราคารวม",
                         "Store"]
    stacked_df.columns = new_column_labels;
    merged_df = pd.merge(stacked_df, dfmaster, on="Material", how="left");
    merged_df['Store'] = merged_df['Store'].astype('object')
    valid_store_values = [2403, 2401, 2408, 2409, 2417, 2402];
    merged_df.loc[~merged_df["Store"].isin(valid_store_values), "Store"] = "อื่นๆ";
    merged_df = merged_df[merged_df['Store'] != "อื่นๆ"]
    merged_df["หน่วย"] = pd.to_numeric(merged_df["หน่วย"].astype(str).str.replace(r'.*/ ', '', regex=True),
                                       errors='coerce').fillna(1).astype(int);
    merged_df["จำนวน"] = merged_df["จำนวน"] * merged_df["หน่วย"]
    if "ต้นทุน" not in merged_df.columns: st.error("ไม่พบคอลัมน์ 'ต้นทุน' ในข้อมูล"); return None
    merged_df["ราคาทุนรวม"] = merged_df["จำนวน"] * merged_df["ต้นทุน"]
    grouped_sumRate_df = merged_df.pivot_table(index=['Material', "Store", 'Material description', 'หน่วยย่อย'],
                                               values=['จำนวน', "ราคาทุนรวม", "ราคารวม"], aggfunc='sum').reset_index()
    grouped_Valuesum_df = merged_df.groupby('Store')[['ราคาทุนรวม', 'ราคารวม']].sum().reset_index()
    grouped_Valuesum_df.columns = ['Store', 'Sum of Cost price', 'Sum of sale price'];
    remainFinal = pd.merge(remain, grouped_Valuesum_df, on='Store', how='left')
    with np.errstate(divide='ignore', invalid='ignore'):
        remainFinal["วันสำรองคงคลัง"] = (remainFinal["Stock Value on Period End"] / remainFinal[
            "Sum of Cost price"]) * 30
    remainFinal.replace([np.inf, -np.inf], 0, inplace=True);
    remainFinal["วันสำรองคงคลัง"].fillna(0, inplace=True)
    return {'ยอดขาย-คงคลัง-สำรองคงคลัง': remainFinal, 'ยอดขาย': grouped_sumRate_df, 'Raw': merged_df}


# ==============================================================================
# FUNCTION 6: วิเคราะห์ ABC (ABC Analysis) - NEW FUNCTION
# ==============================================================================
def process_abc_analysis(inventory_files, master_file):
    """
    Performs ABC analysis and generates a fully formatted Excel report in memory.
    """
    # --- Part 1: Data Loading and Consolidation ---
    try:
        all_dfs = [pd.read_excel(fp) for fp in inventory_files]
        consolidated_df = pd.concat(all_dfs, ignore_index=True)
        master_df = pd.read_excel(master_file, sheet_name='Drug master', usecols=['Material', 'Drug group'])
        master_df['Material'] = master_df['Material'].astype(str)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์: {e}")
        return None

    # --- Part 2: Main Analysis Logic (perform_abc_analysis) ---
    df = consolidated_df
    df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
    df.dropna(subset=['Posting Date'], inplace=True)
    df['MonthYear'] = df['Posting Date'].dt.to_period('M')
    df['Amt.in Loc.Cur.'] = pd.to_numeric(df['Amt.in Loc.Cur.'], errors='coerce').fillna(0)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    df['Material'] = df['Material'].astype(str)

    monthly_data = df.groupby(['Material', 'Material description', 'Storage location', 'MonthYear']).agg(
        MonthlyNetConsumption=('Amt.in Loc.Cur.', 'sum'),
        MonthlyNetQuantity=('Quantity', 'sum')
    ).reset_index()

    monthly_qty_pivot = monthly_data.pivot_table(
        index=['Material', 'Material description', 'Storage location'],
        columns='MonthYear', values='MonthlyNetQuantity', fill_value=0
    )
    monthly_qty_pivot.columns = [f"Qty_{str(col)}" for col in monthly_qty_pivot.columns]
    monthly_qty_pivot = monthly_qty_pivot.abs()

    final_agg = monthly_data.groupby(['Material', 'Material description', 'Storage location']).agg(
        AvgMonthlyNetQuantity=('MonthlyNetQuantity', 'mean'),
        TotalNetConsumption=('MonthlyNetConsumption', 'sum')
    ).reset_index()

    final_agg['AvgMonthlyNetQuantity'] = final_agg['AvgMonthlyNetQuantity'].abs()
    final_agg = pd.merge(final_agg, monthly_qty_pivot, on=['Material', 'Material description', 'Storage location'],
                         how='left')
    final_agg['NetConsumptionValue'] = final_agg['TotalNetConsumption'].abs()
    abc_data_no_class = final_agg[final_agg['NetConsumptionValue'] > 0].copy()

    if abc_data_no_class.empty:
        st.warning("ไม่พบข้อมูลการใช้งาน (consumption data) ที่มีมูลค่ามากกว่า 0")
        return None

    abc_data_no_class = pd.merge(abc_data_no_class, master_df, on='Material', how='left')
    abc_data_no_class['Drug group'].fillna('N/A', inplace=True)

    all_locations_classified = []
    for location in abc_data_no_class['Storage location'].unique():
        loc_df = abc_data_no_class[abc_data_no_class['Storage location'] == location].copy()
        total_value_loc = loc_df['NetConsumptionValue'].sum()
        loc_df = loc_df.sort_values(by='NetConsumptionValue', ascending=False).reset_index(drop=True)
        loc_df['PercentageValue'] = loc_df['NetConsumptionValue'] / total_value_loc if total_value_loc > 0 else 0
        loc_df['CumulativePercentage'] = loc_df['PercentageValue'].cumsum()

        def assign_abc_class(cum_perc):
            if cum_perc <= 0.80:
                return 'A'
            elif cum_perc <= 0.95:
                return 'B'
            else:
                return 'C'

        loc_df['ABC_Class'] = loc_df['CumulativePercentage'].apply(assign_abc_class)
        all_locations_classified.append(loc_df)

    final_results = pd.concat(all_locations_classified)

    # --- Part 3: Excel Report Generation (save_results_to_excel & apply_formats) ---
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        # Helper function for formatting
        def apply_formats_and_hide_cols(writer, sheet_name, df):
            worksheet = writer.sheets[sheet_name]
            center_align = Alignment(horizontal='center', vertical='center')
            col_map = {'AvgMonthlyNetQuantity': '#,##0', 'NetConsumptionValue': '#,##0.00', 'PercentageValue': '0.00%',
                       'CumulativePercentage': '0.00%'}
            for col in df.columns:
                if isinstance(col, str) and col.startswith('Qty_'): col_map[col] = '#,##0'
            col_letters = {col_name: chr(65 + i) for i, col_name in enumerate(df.columns)}
            for col_name, num_format in col_map.items():
                if col_name in col_letters:
                    col_letter = col_letters[col_name]
                    for row in range(2, worksheet.max_row + 1): worksheet[
                        f'{col_letter}{row}'].number_format = num_format
            for row in range(2, worksheet.max_row + 1): worksheet[
                f'{col_letters["ABC_Class"]}{row}'].alignment = center_align
            for col_name in df.columns:
                if isinstance(col_name, str) and col_name.startswith('Qty_'): worksheet.column_dimensions[
                    col_letters[col_name]].hidden = True
            for col in worksheet.columns:
                if not worksheet.column_dimensions[col[0].column_letter].hidden:
                    max_length = max(len(str(cell.value)) for cell in col if cell.value)
                    worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

        # 1. Executive Summary Sheet
        worksheet = writer.book.create_sheet("Executive Summary", 0)
        writer.sheets['Executive Summary'] = worksheet
        current_row = 1
        summary_abc_count = final_results.groupby(['Storage location', 'ABC_Class']).size().unstack(fill_value=0)
        for c in ['A', 'B', 'C']:
            if c not in summary_abc_count: summary_abc_count[c] = 0
        summary_abc_count = summary_abc_count[['A', 'B', 'C']]
        summary_abc_count['Total'] = summary_abc_count.sum(axis=1)
        summary_abc_count.loc['Total'] = summary_abc_count.sum()
        worksheet.cell(row=current_row, column=1, value='สรุปจำนวนรายการ A, B, C ในแต่ละคลัง').font = Font(bold=True)
        current_row += 1
        summary_abc_count.to_excel(writer, sheet_name='Executive Summary', startrow=current_row, startcol=0)
        current_row += summary_abc_count.shape[0] + 3

        worksheet.cell(row=current_row, column=1,
                       value='กลุ่มยา (Drug Group) ที่มีมูลค่าการใช้งานสูงสุด 3 อันดับแรก (แยกตามคลัง)').font = Font(
            bold=True)
        current_row += 1
        top_groups = final_results.groupby('Storage location').apply(
            lambda x: x.groupby('Drug group')['NetConsumptionValue'].sum().nlargest(3)).reset_index()
        top_groups['NetConsumptionValue'] = top_groups['NetConsumptionValue'].map('{:,.2f}'.format)
        top_groups.to_excel(writer, sheet_name='Executive Summary', startrow=current_row, startcol=0, index=False)
        current_row += top_groups.shape[0] + 3

        worksheet.cell(row=current_row, column=1,
                       value='รายการยาที่มีมูลค่าการใช้งานสูงสุด 5 อันดับแรก (แยกตามคลัง)').font = Font(bold=True)
        current_row += 1
        top_items = final_results.groupby('Storage location').apply(
            lambda x: x.groupby(['Material', 'Material description'])['NetConsumptionValue'].sum().nlargest(
                5)).reset_index()
        top_items['NetConsumptionValue'] = top_items['NetConsumptionValue'].map('{:,.2f}'.format)
        top_items.to_excel(writer, sheet_name='Executive Summary', startrow=current_row, startcol=0, index=False)

        # 2. Individual Location Sheets
        for location in final_results['Storage location'].unique():
            sheet_df = final_results[final_results['Storage location'] == location].copy()
            sheet_name = f'SLoc_{location}'
            monthly_cols = sorted([col for col in sheet_df.columns if isinstance(col, str) and col.startswith('Qty_')])
            output_columns = ['Material', 'Material description', 'Storage location'] + monthly_cols + [
                'AvgMonthlyNetQuantity', 'NetConsumptionValue', 'PercentageValue', 'CumulativePercentage', 'ABC_Class',
                'Drug group']
            sheet_df = sheet_df[output_columns]
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_formats_and_hide_cols(writer, sheet_name, sheet_df)

    return output_buffer.getvalue()


# ==============================================================================
# STREAMLIT USER INTERFACE (UI)
# ==============================================================================
st.set_page_config(layout="wide")

st.sidebar.title("⚙️ เลือกเมนู")
app_mode = st.sidebar.selectbox(
    "โปรดเลือกฟังก์ชันที่ต้องการ:",
    ["หน้าหลัก", "1. รายงานยา จ2", "2. รายงานขายยาประจำเดือน", "3. รายงานยา EPI", "4. รายงานยาเสพติดและวัตถุออกฤทธิ์",
     "5. รายงาน KPI", "6. รวมไฟล์ PDF", "7. วิเคราะห์ ABC"]
)

if app_mode == "หน้าหลัก":
    st.title("ยินดีต้อนรับสู่แอปพลิเคชันประมวลผลข้อมูล")
    st.markdown("กรุณาเลือกฟังก์ชันจากเมนูด้านซ้ายเพื่อเริ่มต้น")
    st.markdown("- **1. รายงานยา จ2**: รวมไฟล์และกรองยาตามรายการที่กำหนด (J2)")
    st.markdown("- **2. รายงานขายยาประจำเดือน**: วิเคราะห์ข้อมูลยาโดยละเอียดพร้อมไฟล์ Master")
    st.markdown("- **3. รายงานยา EPI**: สรุปยอดการใช้ยาตามรายการ EPI")
    st.markdown("- **4. รายงานยาเสพติดและวัตถุออกฤทธิ์**: สร้างรายงานยาเสพติดฯ")
    st.markdown("- **5. รายงาน KPI**: คำนวณวันสำรองคงคลังจากข้อมูลยอดขายและยอดคงคลัง")
    st.markdown("- **6. รวมไฟล์ PDF**: รวมไฟล์ PDF หลายไฟล์ให้เป็นไฟล์เดียว")
    st.markdown("- **7. วิเคราะห์ ABC**: จัดกลุ่มยาตามมูลค่าการใช้งานและสร้างรายงานสรุป")

elif app_mode == "1. รายงานยา จ2":
    st.title("Tool 1: รายงานยา จ2");
    st.info("ฟังก์ชันนี้จะรวมไฟล์ข้อมูลดิบ (.xls) จากนั้นกรองรายการยาตามที่กำหนดสำหรับรายงาน จ2")
    uploaded_files_j2 = st.file_uploader("อัปโหลดไฟล์ข้อมูลดิบของคุณ (*.xls)", type="xls", accept_multiple_files=True,
                                         key="j2_uploader")
    if st.button("ประมวลผลรายงาน จ2", key="j2_button"):
        if uploaded_files_j2:
            with st.spinner("กำลังประมวลผล..."):
                final_df = process_j2_report(uploaded_files_j2)
            if final_df is not None:
                st.success("✅ ประมวลผลสำเร็จ!");
                st.dataframe(final_df)
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: final_df.to_excel(writer,
                                                                                                     sheet_name='Raw',
                                                                                                     index=False)
                st.download_button(label="📥 ดาวน์โหลดไฟล์ J2.xlsx", data=output_buffer.getvalue(), file_name="J2.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ข้อมูล")

elif app_mode == "2. รายงานขายยาประจำเดือน":
    st.title("Tool 2: รายงานขายยาประจำเดือน");
    st.info("ฟังก์ชันนี้ต้องการทั้งไฟล์ข้อมูลดิบ (.xls) และไฟล์ Drug Master (.xlsx)")
    col1, col2 = st.columns(2)
    with col1:
        uploaded_files_raw = st.file_uploader("1. อัปโหลดไฟล์ข้อมูลดิบ (*.xls)", type="xls", accept_multiple_files=True,
                                              key="raw_uploader")
    with col2:
        master_file = st.file_uploader("2. อัปโหลดไฟล์ Drug Master (*.xlsx)", type=["xlsx"], key="master_uploader")
    if st.button("🚀 เริ่มการวิเคราะห์", key="analysis_button"):
        if uploaded_files_raw and master_file:
            with st.spinner("กำลังวิเคราะห์ข้อมูล..."):
                raw_df, output_dfs = process_drug_rate_analysis(uploaded_files_raw, master_file)
            if raw_df is not None:
                st.success("✅ วิเคราะห์สำเร็จ!");
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    for sheet_name, df_to_save in output_dfs.items():
                        if sheet_name != "Summary_Data": df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
                    startrow = 0
                    for label, df_summary in output_dfs["Summary_Data"].items():
                        summary_pivot = df_summary.set_index('Month').T;
                        summary_pivot.index = [label]
                        summary_pivot.to_excel(writer, sheet_name='Summary', startrow=startrow);
                        startrow += summary_pivot.shape[0] + 2
                st.download_button(label="📥 ดาวน์โหลดรายงานวิเคราะห์", data=output_buffer.getvalue(),
                                   file_name="Drugstore_Rate.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.subheader("📊 ดูตัวอย่างผลลัพธ์");
                tab1, tab2, tab3 = st.tabs(["Rate by Month", "Cases per Month", "Raw Merged Data"])
                with tab1:
                    st.dataframe(output_dfs["Rate แยกเดือน"]); 
                with tab2: st.dataframe(output_dfs["จำนวนเคสต่อเดือน"]);
                with tab3:
                    st.dataframe(raw_df)
        else:
            st.warning("⚠️ กรุณาอัปโหลดทั้งไฟล์ข้อมูลดิบและไฟล์ Drug Master")

elif app_mode == "3. รายงานยา EPI":
    st.title("Tool 3: รายงานยา EPI");
    st.info("ฟังก์ชันนี้จะรวมไฟล์ข้อมูลดิบ (.xls) จากนั้นกรองเฉพาะรายการยา EPI และสรุปยอดการใช้งานทั้งหมด")
    uploaded_files_epi = st.file_uploader("อัปโหลดไฟล์ข้อมูลดิบของคุณ (*.xls)", type="xls", accept_multiple_files=True,
                                          key="epi_uploader")
    if st.button("ประมวลผลรายงาน EPI", key="epi_button"):
        if uploaded_files_epi:
            with st.spinner("กำลังประมวลผล..."):
                final_df = process_epi_usage(uploaded_files_epi)
            if final_df is not None:
                st.success("✅ ประมวลผลสำเร็จ!");
                st.subheader("ตารางสรุปยอดใช้ยา EPI");
                st.dataframe(final_df)
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: final_df.to_excel(writer,
                                                                                                     sheet_name='Raw',
                                                                                                     index=False)
                st.download_button(label="📥 ดาวน์โหลดไฟล์ EPI usage.xlsx", data=output_buffer.getvalue(),
                                   file_name="EPI usage.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ข้อมูล")

elif app_mode == "4. รายงานยาเสพติดและวัตถุออกฤทธิ์":
    st.title("Tool 4: รายงานยาเสพติดและวัตถุออกฤทธิ์");
    st.info("ฟังก์ชันนี้ต้องการไฟล์ 3 ชุด: ข้อมูลการจ่ายยา, รายงานรับเข้า, และ Drug Master")
    col1, col2, col3 = st.columns(3)
    with col1:
        xls_files = st.file_uploader("1. อัปโหลดไฟล์ข้อมูลการจ่ายยา (*.xls)", type="xls", accept_multiple_files=True,
                                     key="narcotics_xls_uploader")
    with col2:
        receipt_file = st.file_uploader("2. อัปโหลดไฟล์รายงานรับเข้า (*.xlsx)", type="xlsx",
                                        key="narcotics_receipt_uploader")
    with col3:
        master_file_narcotics = st.file_uploader("3. อัปโหลดไฟล์ Drug Master (*.xlsx)", type="xlsx",
                                                 key="narcotics_master_uploader")
    if st.button("🚀 ประมวลผลรายงานยาเสพติด", key="narcotics_button"):
        if xls_files and receipt_file and master_file_narcotics:
            with st.spinner("กำลังประมวลผล..."):
                output_data = process_narcotics_report(xls_files, receipt_file, master_file_narcotics)
            if output_data:
                st.success("✅ ประมวลผลสำเร็จ!")
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    for sheet_name, df in output_data.items(): df.to_excel(writer, sheet_name=sheet_name, index=False)
                st.download_button(label="📥 ดาวน์โหลดรายงานยาเสพติด.xlsx", data=output_buffer.getvalue(),
                                   file_name="รายงานการรับเข้าและจ่าย.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.subheader("📊 ดูตัวอย่างผลลัพธ์");
                tab1, tab2, tab3 = st.tabs(output_data.keys())
                with tab1:
                    st.dataframe(output_data['รายงานแยก']);
                with tab2: st.dataframe(output_data['รายงานรวม']);
                with tab3:
                    st.dataframe(output_data['รายงานรับเข้า'])
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ส่วน")

elif app_mode == "5. รายงาน KPI":
    st.title("Tool 5: รายงาน KPI");
    st.info("ฟังก์ชันนี้ต้องการไฟล์ 3 ชุด: Rate, ยอดคงคลัง, และ Drug Master เพื่อคำนวณวันสำรองคงคลัง")
    col1, col2, col3 = st.columns(3)
    with col1:
        rate_file = st.file_uploader("1. อัปโหลดไฟล์ Rate (*.xls)", type="xls", key="kpi_rate_uploader")
    with col2:
        inventory_file = st.file_uploader("2. อัปโหลดไฟล์ยอดคงคลังสิ้นเดือน (*.xlsx)", type="xlsx",
                                          key="kpi_inventory_uploader")
    with col3:
        master_file_kpi = st.file_uploader("3. อัปโหลดไฟล์ Drug Master (*.xlsx)", type="xlsx",
                                           key="kpi_master_uploader")
    if st.button("🚀 ประมวลผล KPI", key="kpi_button"):
        if rate_file and inventory_file and master_file_kpi:
            with st.spinner("กำลังคำนวณ KPI..."):
                output_data = process_kpi_report(rate_file, inventory_file, master_file_kpi)
            if output_data:
                st.success("✅ ประมวลผลสำเร็จ!")
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    for sheet_name, df in output_data.items(): df.to_excel(writer, sheet_name=sheet_name, index=False)
                st.download_button(label="📥 ดาวน์โหลดรายงาน KPI.xlsx", data=output_buffer.getvalue(),
                                   file_name="KPI_Report.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.subheader("📊 ดูตัวอย่างผลลัพธ์");
                tab1, tab2, tab3 = st.tabs(output_data.keys())
                with tab1:
                    st.dataframe(output_data['ยอดขาย-คงคลัง-สำรองคงคลัง']);
                with tab2: st.dataframe(output_data['ยอดขาย']);
                with tab3:
                    st.dataframe(output_data['Raw'])
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ส่วน")

elif app_mode == "6. รวมไฟล์ PDF":
    st.title("Tool 6: รวมไฟล์ PDF");
    st.info("อัปโหลดไฟล์ PDF หลายไฟล์เพื่อรวมเป็นไฟล์เดียว")
    uploaded_pdfs = st.file_uploader("1. เลือกไฟล์ PDF ที่ต้องการรวม", type="pdf", accept_multiple_files=True,
                                     key="pdf_uploader")
    output_filename = st.text_input("2. ตั้งชื่อไฟล์ผลลัพธ์ (ไม่ต้องใส่ .pdf)", "merged_output", key="pdf_output_name")
    if st.button("รวมไฟล์ PDF", key="pdf_merge_button"):
        if uploaded_pdfs and output_filename:
            with st.spinner("กำลังรวมไฟล์ PDF..."):
                merger = PdfMerger();
                for pdf_file in uploaded_pdfs: merger.append(pdf_file)
                pdf_buffer = io.BytesIO();
                merger.write(pdf_buffer);
                merger.close()
                final_filename = f"{output_filename.strip()}.pdf"
                st.success(f"✅ รวมไฟล์ PDF สำเร็จ!")
                st.download_button(label=f"📥 ดาวน์โหลดไฟล์ {final_filename}", data=pdf_buffer.getvalue(),
                                   file_name=final_filename, mime="application/pdf")
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ PDF และตั้งชื่อไฟล์ผลลัพธ์")

elif app_mode == "7. วิเคราะห์ ABC":
    st.title("Tool 7: วิเคราะห์ ABC (ABC Analysis)")
    st.info(
        "ฟังก์ชันนี้ต้องการไฟล์ข้อมูลการใช้งาน (Inventory Data) และไฟล์ Drug Master เพื่อจัดกลุ่มยาตามมูลค่าการใช้งาน (A, B, C)")

    col1, col2 = st.columns(2)
    with col1:
        inventory_files = st.file_uploader(
            "1. อัปโหลดไฟล์ข้อมูลการใช้งาน (หลายไฟล์ได้)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="abc_inventory_uploader"
        )
    with col2:
        master_file_abc = st.file_uploader(
            "2. อัปโหลดไฟล์ Drug Master (*.xlsx)",
            type="xlsx",
            key="abc_master_uploader"
        )

    if st.button("🚀 เริ่มการวิเคราะห์ ABC", key="abc_button"):
        if inventory_files and master_file_abc:
            with st.spinner("กำลังทำการวิเคราะห์ ABC... กระบวนการนี้อาจใช้เวลาสักครู่"):
                report_bytes = process_abc_analysis(inventory_files, master_file_abc)

            if report_bytes:
                st.success("✅ การวิเคราะห์ ABC เสร็จสมบูรณ์และสร้างรายงานสำเร็จ!")
                st.download_button(
                    label="📥 ดาวน์โหลดรายงาน ABC Analysis",
                    data=report_bytes,
                    file_name="Consolidated_ABC_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ส่วน")