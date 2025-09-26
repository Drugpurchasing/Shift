import streamlit as st
import pandas as pd
import io

# ==============================================================================
# FUNCTION 1: รายงานยา จ2 (J2 Report)
# ==============================================================================
def process_j2_report(uploaded_files):
    """
    Function based on the very first script.
    Merges files, cleans data, and filters by a specific list of materials.
    """
    # 1. Load and combine all uploaded raw data files (*.xls)
    dfs = []
    for file_obj in uploaded_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0:
                    df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}")
            return None
    
    if not dfs:
        st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด")
        return None
        
    stacked_df = pd.concat(dfs, ignore_index=True)

    # 2. Basic data cleaning
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')

    # 3. Rename columns
    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", 
        "สิทธิ์", "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]
    if len(stacked_df.columns) != len(new_column_labels):
        st.error(f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}")
        return None
    stacked_df.columns = new_column_labels

    # 4. Filter for specific J2 materials
    valid_material_values = [
        1400000010, 1400000020, 1400000021, 1400000025, 1400000029, 
        1400000030, 1400000040, 1400000044, 1400000052, 1400000053, 
        1400000055, 1400000098, 1400000099, 1400000148, 1400000187, 
        1400000201, 1400000220, 1400000221, 1400000228, 1400000247, 
        1400000264, 1400000068, 1400000069, 1400000093, 1400000106, 
        1400000113, 1400000115, 1400000116, 1400000118, 1400000124, 
        1400000126, 1400000130, 1400000165, 1400000166, 1400000167, 
        1400000168, 1400000169, 1400000170, 1400000171, 1400000172, 
        1400000194, 1400000284, 1400000288, 1400000294, 1400000295, 
        1400000331, 1400000335, 1400000344, 1400000345, 1400000265
    ]
    merged_df = stacked_df[stacked_df["Material"].isin(valid_material_values)].copy()

    # 5. Select and reorder the final columns
    final_cols = ['วันที่จ่ายยา', 'VN / AN', 'HN', 'ชื่อ', 'สิทธิ์', "แพทย์", 'Material', 'รายการยา', 'จำนวน']
    merged_df = merged_df[final_cols]

    return merged_df

# ==============================================================================
# FUNCTION 2: รายงานขายยาประจำเดือน (Drug Rate Analysis)
# ==============================================================================
def process_drug_rate_analysis(data_files, master_file):
    """
    Function based on the second script with Drug Master file.
    Merges data, performs extensive analysis, and generates multiple pivot tables.
    """
    # 1. Load and combine all uploaded raw data files (*.xls)
    dfs = []
    for file_obj in data_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0:
                    df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}")
            return None, {}

    if not dfs:
        st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด")
        return None, {}
    
    stacked_df = pd.concat(dfs, ignore_index=True)

    # 2. Load the Drug Master file
    try:
        dfmaster = pd.read_excel(master_file, sheet_name="Drug master")
    except Exception as e:
        st.error(f"ไม่สามารถอ่านชีท 'Drug master' จากไฟล์ Master ได้: {e}")
        return None, {}

    # 3. Data Cleaning and Preprocessing
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    
    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", 
        "สิทธิ์", "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]
    if len(stacked_df.columns) != len(new_column_labels):
        st.error(f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}")
        return None, {}
    stacked_df.columns = new_column_labels

    # 4. Merging and Transformations
    merged_df = pd.merge(stacked_df, dfmaster, on="Material", how="left")
    
    # FIX for FutureWarning: Change dtype to 'object' to allow mixed types
    merged_df['Store'] = merged_df['Store'].astype('object')
    
    valid_store_values = [2403, 2401, 2408, 2409, 2417, 2402]
    merged_df.loc[~merged_df["Store"].isin(valid_store_values), "Store"] = "อื่นๆ"
    
    merged_df["ราคาทุนรวม"] = pd.to_numeric(merged_df["จำนวน"], errors='coerce') * pd.to_numeric(merged_df["ต้นทุน"], errors='coerce')
    merged_df['วันที่จ่ายยา'] = pd.to_datetime(merged_df['วันที่จ่ายยา'], errors='coerce')
    merged_df['Month'] = merged_df['วันที่จ่ายยา'].dt.to_period('M')
    merged_df = merged_df[merged_df['Store'] != "อื่นๆ"]

    # --- Mapping 'สิทธิ์' (example, full dictionary should be here) ---
    direct_map = { '(ตรวจที่รพ.จุฬาภรณ์) โครงการคัดกรองมะเร็งปากมดลูก ณ รพ.จุฬาภรณ์และคณะแพทย์ศาสตร์วชิรพยาบาล': 'จ่ายเอง', '[TopUp] สวัสดิการเจ้าหน้าที่ราชวิทยาลัยจุฬาภรณ์': 'สวัสดิการเจ้าหน้าที่', 'องค์การปกครองส่วนท้องถิ่นบำนาญ(เบิกจ่ายตรง)': 'ข้าราชการ'}
    merged_df["สิทธิ์"] = merged_df["สิทธิ์"].map(direct_map).fillna(merged_df["สิทธิ์"])
    
    # 5. Data Splitting and Unique Counts (OPD/IPD)
    opd_merged_df = merged_df[merged_df['เลขที่เอกสาร'].isna() | (merged_df['เลขที่เอกสาร'].astype(str).str.strip().isin(['', '0']))]
    opd_2409 = opd_merged_df[opd_merged_df['Store'].notna() & (opd_merged_df['Store'] == 2409)]
    opd_not_2409 = opd_merged_df[opd_merged_df['Store'].notna() & (opd_merged_df['Store'] != 2409)]
    
    ipd_merged_df = merged_df[merged_df['Clinic'].isna() | (merged_df['Clinic'].astype(str).str.strip().isin(['', '0']))]
    ipd_2409 = ipd_merged_df[ipd_merged_df['Store'].notna() & (ipd_merged_df['Store'] == 2409)]
    ipd_not_2409 = ipd_merged_df[ipd_merged_df['Store'].notna() & (ipd_merged_df['Store'] != 2409)]
    
    def count_unique_by_month(df, subset_cols):
        return df.drop_duplicates(subset=subset_cols).groupby('Month').size().reset_index(name='Unique_Count')

    uniqueOPD = count_unique_by_month(opd_not_2409, ['VN / AN', 'HN', 'Clinic', 'Month'])
    uniqueOPD2409 = count_unique_by_month(opd_2409, ['VN / AN', 'HN', 'Clinic', 'Month'])
    uniqueIPD = count_unique_by_month(ipd_not_2409, ['เลขที่เอกสาร', 'HN', 'Ward', 'Month'])
    uniqueIPD2409 = count_unique_by_month(ipd_2409, ['เลขที่เอกสาร', 'HN', 'Ward', 'Month'])

    # 6. Final Cleaning and Pivot Table Generation
    merged_df["หน่วย"] = pd.to_numeric(merged_df["หน่วย"].astype(str).str.replace(r'.*/ ', '', regex=True), errors='coerce').fillna(1).astype(int)
    merged_df["จำนวน"] = merged_df["จำนวน"] * merged_df["หน่วย"]
    merged_df['HN'] = merged_df['HN'].astype(str).str.replace('.0', '', regex=False)

    grouped_countHN_df = merged_df.pivot_table(index=['Material', 'Material description'], columns='Month', values='HN', aggfunc=pd.Series.nunique).reset_index()
    grouped_sumRate_df = merged_df.pivot_table(index=['Material', 'Material description', 'หน่วยย่อย'], columns='Month', values='จำนวน', aggfunc='sum').reset_index()
    grouped_sumRateSplit_df = merged_df.pivot_table(index=['Material', "Store", 'Material description', 'หน่วยย่อย'], columns='Month', values='จำนวน', aggfunc='sum').reset_index()

    output_dfs = {
        "Rate แยกเดือน": grouped_sumRate_df,
        "Rate (M-Sloc)": grouped_sumRateSplit_df,
        "จำนวนเคสต่อเดือน": grouped_countHN_df,
        "Raw": merged_df,
        "Summary_Data": {
            'จำนวนใบยา OPD': uniqueOPD, 'จำนวนใบยา OPD 2409': uniqueOPD2409,
            'จำนวนใบยา IPD': uniqueIPD, 'จำนวนใบยา IPD 2409': uniqueIPD2409,
        }
    }
    
    return merged_df, output_dfs

# ==============================================================================
# FUNCTION 3: รายงานยา EPI (EPI Usage Report)
# ==============================================================================
def process_epi_usage(uploaded_files):
    """
    Processes uploaded files to generate a summary of EPI drug usage.
    """
    # 1. Load and combine all uploaded raw data files (*.xls)
    dfs = []
    for file_obj in uploaded_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0:
                    df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_obj.name}: {e}")
            return None

    if not dfs:
        st.warning("ไม่พบข้อมูลในไฟล์ที่อัปโหลด")
        return None
    
    stacked_df = pd.concat(dfs, ignore_index=True)

    # 2. Basic data cleaning
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')
    
    # 3. Rename columns
    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "อายุ", 
        "สิทธิ์", "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]
    if len(stacked_df.columns) != len(new_column_labels):
        st.error(f"จำนวนคอลัมน์ไม่ตรงกัน คาดว่าต้องมี {len(new_column_labels)} แต่พบ {len(stacked_df.columns)}")
        return None
    stacked_df.columns = new_column_labels

    # 4. Filter for specific EPI materials
    valid_epi_materials = [
        1400000084, 1400000083, 1400000087, 1400000086, 1400000088,
        1400000081, 1400000082, 1400000090, 1400000085, 1400000089
    ]
    epi_df = stacked_df[stacked_df["Material"].isin(valid_epi_materials)].copy()

    # 5. Group by Material and sum the quantity
    summary_df = epi_df.groupby(['Material', 'รายการยา'])['จำนวน'].sum().reset_index()
    summary_df.rename(columns={'จำนวน': 'จำนวนรวม'}, inplace=True)
    
    return summary_df

# ==============================================================================
# STREAMLIT USER INTERFACE (UI)
# ==============================================================================
st.set_page_config(layout="wide")

st.sidebar.title("⚙️ เลือกเมนู")
app_mode = st.sidebar.selectbox(
    "โปรดเลือกฟังก์ชันที่ต้องการ:",
    ["หน้าหลัก", "1. รายงานยา จ2", "2. รายงานขายยาประจำเดือน", "3. รายงานยา EPI"]
)

if app_mode == "หน้าหลัก":
    st.title("ยินดีต้อนรับสู่แอปพลิเคชันประมวลผลข้อมูล")
    st.markdown("กรุณาเลือกฟังก์ชันจากเมนูด้านซ้ายเพื่อเริ่มต้น")
    st.markdown("- **1. รายงานยา จ2**: รวมไฟล์และกรองยาตามรายการที่กำหนด (J2)")
    st.markdown("- **2. รายงานขายยาประจำเดือน**: วิเคราะห์ข้อมูลยาโดยละเอียดพร้อมไฟล์ Master")
    st.markdown("- **3. รายงานยา EPI**: สรุปยอดการใช้ยาตามรายการ EPI")

elif app_mode == "1. รายงานยา จ2":
    st.title("Tool 1: รายงานยา จ2")
    st.info("ฟังก์ชันนี้จะรวมไฟล์ข้อมูลดิบ (.xls) จากนั้นกรองรายการยาตามที่กำหนดสำหรับรายงาน จ2")

    uploaded_files_j2 = st.file_uploader(
        "อัปโหลดไฟล์ข้อมูลดิบของคุณ (*.xls)",
        type="xls", accept_multiple_files=True, key="j2_uploader"
    )
    
    if st.button("ประมวลผลรายงาน จ2", key="j2_button"):
        if uploaded_files_j2:
            with st.spinner("กำลังประมวลผล..."):
                final_df = process_j2_report(uploaded_files_j2)
            if final_df is not None:
                st.success("✅ ประมวลผลสำเร็จ!")
                st.dataframe(final_df)
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='Raw', index=False)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ J2.xlsx", data=output_buffer.getvalue(),
                    file_name="J2.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ข้อมูล")

elif app_mode == "2. รายงานขายยาประจำเดือน":
    st.title("Tool 2: รายงานขายยาประจำเดือน")
    st.info("ฟังก์ชันนี้ต้องการทั้งไฟล์ข้อมูลดิบ (.xls) และไฟล์ Drug Master (.xlsx)")

    col1, col2 = st.columns(2)
    with col1:
        uploaded_files_raw = st.file_uploader(
            "1. อัปโหลดไฟล์ข้อมูลดิบ (*.xls)",
            type="xls", accept_multiple_files=True, key="raw_uploader"
        )
    with col2:
        master_file = st.file_uploader(
            "2. อัปโหลดไฟล์ Drug Master (*.xlsx)",
            type=["xlsx"], key="master_uploader"
        )
        
    if st.button("🚀 เริ่มการวิเคราะห์", key="analysis_button"):
        if uploaded_files_raw and master_file:
            with st.spinner("กำลังวิเคราะห์ข้อมูล..."):
                raw_df, output_dfs = process_drug_rate_analysis(uploaded_files_raw, master_file)
            if raw_df is not None:
                st.success("✅ วิเคราะห์สำเร็จ!")
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    for sheet_name, df_to_save in output_dfs.items():
                        if sheet_name != "Summary_Data":
                            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
                    startrow = 0
                    for label, df_summary in output_dfs["Summary_Data"].items():
                        summary_pivot = df_summary.set_index('Month').T
                        summary_pivot.index = [label]
                        summary_pivot.to_excel(writer, sheet_name='Summary', startrow=startrow)
                        startrow += summary_pivot.shape[0] + 2
                st.download_button(
                    label="📥 ดาวน์โหลดรายงานวิเคราะห์ (Drugstore_Rate.xlsx)", data=output_buffer.getvalue(),
                    file_name="Drugstore_Rate.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.subheader("📊 ดูตัวอย่างผลลัพธ์")
                tab1, tab2, tab3 = st.tabs(["Rate by Month", "Cases per Month", "Raw Merged Data"])
                with tab1: st.dataframe(output_dfs["Rate แยกเดือน"])
                with tab2: st.dataframe(output_dfs["จำนวนเคสต่อเดือน"])
                with tab3: st.dataframe(raw_df)
        else:
            st.warning("⚠️ กรุณาอัปโหลดทั้งไฟล์ข้อมูลดิบและไฟล์ Drug Master")

elif app_mode == "3. รายงานยา EPI":
    st.title("Tool 3: รายงานยา EPI")
    st.info("ฟังก์ชันนี้จะรวมไฟล์ข้อมูลดิบ (.xls) จากนั้นกรองเฉพาะรายการยา EPI และสรุปยอดการใช้งานทั้งหมด")

    uploaded_files_epi = st.file_uploader(
        "อัปโหลดไฟล์ข้อมูลดิบของคุณ (*.xls)",
        type="xls", accept_multiple_files=True, key="epi_uploader"
    )
    
    if st.button("ประมวลผลรายงาน EPI", key="epi_button"):
        if uploaded_files_epi:
            with st.spinner("กำลังประมวลผล..."):
                final_df = process_epi_usage(uploaded_files_epi)
            if final_df is not None:
                st.success("✅ ประมวลผลสำเร็จ!")
                st.subheader("ตารางสรุปยอดใช้ยา EPI")
                st.dataframe(final_df)
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='Raw', index=False)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ EPI usage.xlsx", data=output_buffer.getvalue(),
                    file_name="EPI usage.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ข้อมูล")
            