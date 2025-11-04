import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- Main Processing Function (Based on your script) ---
# อัปเดต: เพิ่ม 'mode' parameter เพื่อควบคุมการประมวลผล
def process_files(rate_files_list, data_files_list, mode):
    """
    ประมวลผลไฟล์ Rate และ Data ตาม Logic เดิมของผู้ใช้
    mode: "OPD และ IPD (ทั้งหมด)", "เฉพาะ OPD", "เฉพาะ IPD"
    """
    
    # --- 1. ประมวลผลไฟล์ "Rate" (*.xlsx) ---
    # (ยังคงต้องอ่านทั้งหมด แม้จะทำแค่ OPD หรือ IPD)
    st.write("กำลังอ่านไฟล์ Rate...")
    combined_df = pd.DataFrame()
    for file in rate_files_list:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    # --- 2. ประมวลผลไฟล์ "Data" (*.xls) ---
    # (ยังคงต้องอ่านทั้งหมด เพื่อกรอง)
    st.write("กำลังอ่านไฟล์ข้อมูล *.xls ...")
    dfs = []
    for file in data_files_list:
        try:
            source_workbook = pd.ExcelFile(file)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                if i == 0:
                    df = df.iloc[2:]
                dfs.append(df)
        except Exception as e:
            st.warning(f"ไม่สามารถอ่านไฟล์ {file.name} (ชีท {sheet_name}): {e}")

    if not dfs:
        st.error("ไม่พบข้อมูลที่สามารถอ่านได้ในไฟล์ *.xls ที่อัปโหลด")
        return None, pd.DataFrame(), pd.DataFrame()

    stacked_df = pd.concat(dfs, ignore_index=True)

    # --- 3. การทำความสะอาดและเตรียมข้อมูล (Logic เดิม) ---
    st.write("กำลังทำความสะอาดข้อมูล...")
    
    if 18 not in stacked_df.columns:
        st.error(f"ข้อมูล *.xls ไม่มีคอลัมน์ที่ 12 หรือ 18 (index base 0)")
        return None, pd.DataFrame(), pd.DataFrame()

    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')

    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN",
        "HN", "ชื่อ", "อายุ", "สิทธิ์", "แพทย์", "Clinic",
        "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]
    
    if len(stacked_df.columns) != len(new_column_labels):
        st.error(f"เกิดข้อผิดพลาด: ไฟล์ข้อมูล *.xls มี {len(stacked_df.columns)} คอลัมน์ แต่คาดหวัง {len(new_column_labels)} คอลัมน์")
        return None, pd.DataFrame(), pd.DataFrame()

    stacked_df.columns = new_column_labels
    
    stacked_df = stacked_df.loc[:, ["วันที่จ่ายยา", "เลขที่เอกสาร", "VN / AN", "HN", "ชื่อ", "แพทย์", "Clinic", "Ward", "Material", "รายการยา", "จำนวน", "หน่วย", "Store"]]
    
    valid_store_values = [2403, 2401, 2408, 2409, 2417, 2402]
    stacked_df = stacked_df[stacked_df['Store'].isin(valid_store_values)]
    stacked_df = stacked_df[stacked_df['จำนวน'] >= 0]
    
    # แยก OPD / IPD (Logic เดิม)
    stacked_IPD = stacked_df.dropna(subset=["เลขที่เอกสาร"])
    stacked_OPD = stacked_df.dropna(subset=["Clinic"])

    st.write("ประมวลผลไฟล์ Rate (OPD/IPD)...")
    # --- 4. ประมวลผล Combined (Rate) Data (Logic เดิม) ---
    # (จำเป็นต้องเตรียมข้อมูลทั้ง 2 ส่วนก่อน)
    combined_OPD = combined_df.loc[:, ["Material Number", "Material description", "Batch Quantity", "Order Number", "VN Number", "Hospital Number"]]
    combined_IPD = combined_df.loc[:, ["Material Number", "Material description", "Batch Quantity", "Order Number", "Admit Number"]]
    
    combined_OPD = combined_OPD.groupby(["Material Number", "Order Number", "VN Number", "Hospital Number"])['Batch Quantity'].sum().reset_index()
    combined_IPD = combined_IPD.groupby(["Material Number", "Order Number", "Admit Number"])['Batch Quantity'].sum().reset_index()
    
    new_column_names_opd = ["Material", "Order Number", "VN / AN", "HN", "จำนวน Pick"]
    combined_OPD = combined_OPD.rename(columns=dict(zip(combined_OPD.columns, new_column_names_opd)))
    
    new_column_names_ipd = ["Material", "เลขที่เอกสาร", "VN / AN", "จำนวน Pick"]
    combined_IPD = combined_IPD.rename(columns=dict(zip(combined_IPD.columns, new_column_names_ipd)))

    # --- 5. Merge OPD (Logic เดิม - แต่มีเงื่อนไข) ---
    merged_OPD = pd.DataFrame() # สร้าง DataFrame ว่างไว้ก่อน
    
    if mode in ("OPD และ IPD (ทั้งหมด)", "เฉพาะ OPD"):
        st.write("กำลัง Merge ข้อมูล OPD...")
        try:
            stacked_OPD["HN"] = pd.to_numeric(stacked_OPD["HN"], errors='coerce').fillna(0).astype(np.int64)
            combined_OPD["HN"] = pd.to_numeric(combined_OPD["HN"], errors='coerce').fillna(0).astype(np.int64)
            stacked_OPD["VN / AN"] = stacked_OPD["VN / AN"].astype(str)
            combined_OPD["VN / AN"] = combined_OPD["VN / AN"].astype(str)
            stacked_OPD["Material"] = pd.to_numeric(stacked_OPD["Material"], errors='coerce').fillna(0).astype(np.int64)
            combined_OPD["Material"] = pd.to_numeric(combined_OPD["Material"], errors='coerce').fillna(0).astype(np.int64)

            merged_OPD = pd.merge(stacked_OPD, combined_OPD, on=["HN", "VN / AN", "Material"], how="left")
            merged_OPD.fillna(0, inplace=True)
            merged_OPD['ค้าง PickO'] = merged_OPD['จำนวน'] - merged_OPD['จำนวน Pick']
            merged_OPD = merged_OPD[merged_OPD['ค้าง PickO'] > 0]
            merged_OPD['วันที่จ่ายยา'] = pd.to_datetime(merged_OPD['วันที่จ่ายยา']).dt.strftime('%d/%m/%y')
            merged_OPD = merged_OPD.drop(columns=["เลขที่เอกสาร", "Ward", "Order Number", "จำนวน Pick", "แพทย์", "ชื่อ"])
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดขณะประมวลผล OPD: {e}")

    # --- 6. Merge IPD (Logic เดิม - แต่มีเงื่อนไข) ---
    merged_IPD = pd.DataFrame() # สร้าง DataFrame ว่างไว้ก่อน

    if mode in ("OPD และ IPD (ทั้งหมด)", "เฉพาะ IPD"):
        st.write("กำลัง Merge ข้อมูล IPD...")
        try:
            stacked_IPD["เลขที่เอกสาร"] = stacked_IPD["เลขที่เอกสาร"].astype(str)
            combined_IPD["เลขที่เอกสาร"] = combined_IPD["เลขที่เอกสาร"].astype(str)
            stacked_IPD["VN / AN"] = stacked_IPD["VN / AN"].astype(str)
            combined_IPD["VN / AN"] = combined_IPD["VN / AN"].astype(str)
            stacked_IPD["Material"] = pd.to_numeric(stacked_IPD["Material"], errors='coerce').fillna(0).astype(np.int64)
            combined_IPD["Material"] = pd.to_numeric(combined_IPD["Material"], errors='coerce').fillna(0).astype(np.int64)

            merged_IPD = pd.merge(stacked_IPD, combined_IPD, on=["เลขที่เอกสาร", "VN / AN", "Material"], how="left")
            merged_IPD.fillna(0, inplace=True)
            merged_IPD['ค้าง PickI'] = merged_IPD['จำนวน'] - merged_IPD['จำนวน Pick']
            merged_IPD = merged_IPD[merged_IPD['ค้าง PickI'] > 0]
            merged_IPD['วันที่จ่ายยา'] = pd.to_datetime(merged_IPD['วันที่จ่ายยา']).dt.strftime('%d/%m/%y')
            merged_IPD = merged_IPD.drop(columns=["Clinic", "จำนวน Pick", "แพทย์", "ชื่อ"])
            merged_IPD = merged_IPD[merged_IPD['Material'].between(1200000001, 1400099999, inclusive='both')].reset_index(drop=True)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดขณะประมวลผล IPD: {e}")

    # --- 7. สร้าง Excel Output ใน Memory (แบบมีเงื่อนไข) ---
    st.write("กำลังสร้างไฟล์ Excel...")
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if mode in ("OPD และ IPD (ทั้งหมด)", "เฉพาะ OPD"):
            merged_OPD.to_excel(writer, sheet_name='ค้าง Pick OPD', index=False)
            stacked_OPD.to_excel(writer, sheet_name='Order OPD', index=False)
            combined_OPD.to_excel(writer, sheet_name='OPD Pick', index=False)
        
        if mode in ("OPD และ IPD (ทั้งหมด)", "เฉพาะ IPD"):
            merged_IPD.to_excel(writer, sheet_name='ค้าง Pick IPD', index=False)
            stacked_IPD.to_excel(writer, sheet_name='Order IPD', index=False)
            combined_IPD.to_excel(writer, sheet_name='IPD Pick', index=False)
    
    output.seek(0)
    return output, merged_OPD, merged_IPD

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("💊 โปรแกรมตรวจสอบรายการค้าง Pick (OPD/IPD)")
st.markdown("โปรแกรมนี้จะช่วยรวมไฟล์ข้อมูลยาและไฟล์ Rate เพื่อค้นหารายการที่ยังค้าง Pick")

# --- File Uploaders ---
col1, col2 = st.columns(2)
with col1:
    st.header("ขั้นตอนที่ 1: อัปโหลดไฟล์ Rate")
    rate_files = st.file_uploader("เลือกไฟล์ 'Rate' (*.xlsx)", 
                                  type="xlsx", 
                                  accept_multiple_files=True, 
                                  help="เลือกไฟล์ Rate ทั้งหมดที่ต้องการประมวลผล")

with col2:
    st.header("ขั้นตอนที่ 2: อัปโหลดไฟล์ข้อมูล")
    data_files = st.file_uploader("เลือกไฟล์ข้อมูล (*.xls)", 
                                  type="xls", 
                                  accept_multiple_files=True, 
                                  help="เลือกไฟล์ข้อมูล *.xls ทั้งหมด (แทนการเลือกโฟลเดอร์)")

# --- 🌟 ขั้นตอนที่ 3: เลือกประเภทการประมวลผล (ใหม่) ---
st.header("ขั้นตอนที่ 3: เลือกประเภทการประมวลผล")
processing_mode = st.radio(
    "เลือกประเภทที่ต้องการ",
    ("OPD และ IPD (ทั้งหมด)", "เฉพาะ OPD", "เฉพาะ IPD"),
    horizontal=True,
    label_visibility="collapsed" # ซ่อน Label "เลือกประเภทที่ต้องการ" ซ้ำซ้อน
)

# --- Process Button ---
st.divider()
if st.button("🚀 เริ่มประมวลผล", use_container_width=True, type="primary"):
    
    if rate_files and data_files:
        try:
            with st.spinner("กำลังประมวลผลข้อมูล... กรุณารอสักครู่..."):
                excel_output, df_opd, df_ipd = process_files(rate_files, data_files, processing_mode)
            
            if excel_output:
                st.success("✅ ประมวลผลสำเร็จ!")
                
                # --- Download Button ---
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ 'รายการค้าง Pick.xlsx'",
                    data=excel_output,
                    file_name="รายการค้าง Pick.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # --- Display Results (อัปเดต: แสดงผลตาม mode ที่เลือก) ---
                st.header("สรุปผลการค้าง Pick")
                
                if processing_mode == "เฉพาะ OPD":
                    st.info(f"ค้าง Pick OPD ({len(df_opd)} รายการ)")
                    st.dataframe(df_opd)
                    
                elif processing_mode == "เฉพาะ IPD":
                    st.info(f"ค้าง Pick IPD ({len(df_ipd)} รายการ)")
                    st.dataframe(df_ipd)
                    
                else: # "OPD และ IPD (ทั้งหมด)"
                    tab_opd, tab_ipd = st.tabs([f"ค้าง Pick OPD ({len(df_opd)} รายการ)", 
                                                f"ค้าง Pick IPD ({len(df_ipd)} รายการ)"])
                    with tab_opd:
                        st.dataframe(df_opd)
                    with tab_ipd:
                        st.dataframe(df_ipd)

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดร้ายแรงระหว่างประมวลผล:")
            st.exception(e) # แสดงรายละเอียดข้อผิดพลาด
    
    else:
        st.warning("กรุณาอัปโหลดไฟล์ Rate และไฟล์ข้อมูลให้ครบถ้วน")