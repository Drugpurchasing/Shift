import streamlit as st
import pandas as pd
import io

# --- ฟังก์ชันสำหรับประมวลผลข้อมูล ---
# แยกส่วนการประมวลผลออกมาเป็นฟังก์ชันเพื่อให้โค้ดอ่านง่ายขึ้น
def process_data(df, sloc):
    """
    ฟังก์ชันนี้รับ DataFrame และรหัสคลัง (sloc) เข้ามา
    แล้วทำการประมวลผลข้อมูลตามขั้นตอนเดิมทั้งหมด
    จากนั้นคืนค่าเป็น DataFrame ที่ประมวลผลเสร็จแล้ว
    """
    # การกรองและจัดการข้อมูล (เหมือนโค้ดเดิมของคุณ)
    filterdf = df[df['Material Document'].between(7100000000, 7200000000, inclusive='both')].reset_index(drop=True)
    checkerdf = df[["Goods Receipt/Issue Slip", "Reference", "Material Document"]]
    checkerdf = checkerdf[checkerdf["Reference"] > 0]
    checkerdf = checkerdf.rename(columns={"Reference": "Material Document", "Material Document": "Reference"})
    
    df_processed = df.groupby(["Goods Receipt/Issue Slip", "Material", "Material description", "Batch"])["Quantity"].sum().reset_index()
    df_processed = df_processed[df_processed['Quantity'] < 0].reset_index(drop=True)
    df_processed = pd.merge(df_processed, filterdf, how='inner').drop(columns=['Reference', 'Movement type', 'Plant'])
    df_processed = pd.merge(df_processed, checkerdf, how='left').drop_duplicates()
    
    # การเปลี่ยนชื่อคอลัมน์
    df_processed = df_processed.rename(columns={
        "Goods Receipt/Issue Slip": "Reservation", 
        "Material Document": "Mat Doc", 
        "Storage location": "คลังจ่าย", 
        "Receiving stor. loc.": "คลังรับ",
        "Base Unit of Measure": "หน่วย", 
        "Quantity": "จำนวน"
    })
    
    df_processed = df_processed.reset_index(drop=True)
    df_processed['จำนวน'] = df_processed['จำนวน'].abs()
    
    # จัดเรียงและเลือกคอลัมน์สุดท้าย
    df_processed = df_processed[["Mat Doc", "Posting Date", "Reservation", 'Material', 'Material description', 'Batch', "จำนวน", 'คลังจ่าย', 'คลังรับ']]
    df_processed['Posting Date'] = pd.to_datetime(df_processed['Posting Date']).dt.strftime('%d.%m.%Y')

    # กรองตามรหัสคลัง (sloc) ที่ผู้ใช้ป้อน
    if sloc.lower() != 'all':
        try:
            # แปลง sloc เป็นตัวเลขเพื่อเปรียบเทียบ
            df_processed = df_processed[df_processed["คลังรับ"] == int(sloc)]
        except ValueError:
            # กรณีผู้ใช้ป้อนค่าที่ไม่ใช่ตัวเลขและไม่ใช่ 'All'
            st.error(f"รหัสคลัง '{sloc}' ไม่ถูกต้อง กรุณาใส่ตัวเลขหรือ 'All'")
            return None # คืนค่า None เพื่อหยุดการทำงาน

    return df_processed

# --- ส่วนของหน้าเว็บ Streamlit ---
st.set_page_config(page_title="Excel Data Processor", layout="wide")
st.title("📄 Excel File Processor for ZTRF")
st.write("อัปโหลดไฟล์ Excel ของคุณและใส่รหัสคลังเพื่อประมวลผลข้อมูล")

# 1. วิดเจ็ตสำหรับอัปโหลดไฟล์
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=["xlsx"])

# 2. วิดเจ็ตสำหรับรับข้อความ (รหัสคลัง)
sloc = st.text_input("ใส่รหัสคลังของท่าน (หากต้องการทั้งหมด พิมพ์ 'All')", placeholder="เช่น 1111 หรือ All")

# 3. ปุ่มสำหรับเริ่มการประมวลผล
if st.button("🚀 เริ่มประมวลผล"):
    if uploaded_file is not None and sloc:
        with st.spinner('กำลังประมวลผลไฟล์...'):
            # อ่านไฟล์ Excel ที่อัปโหลดเข้ามา
            df_original = pd.read_excel(uploaded_file)
            
            # เรียกใช้ฟังก์ชันประมวลผล
            df_final = process_data(df_original, sloc)

            if df_final is not None:
                st.success("✅ ประมวลผลไฟล์สำเร็จ!")
                
                # แสดงข้อมูล 5 แถวแรกของผลลัพธ์
                st.write("ตัวอย่างข้อมูลหลังการประมวลผล:")
                st.dataframe(df_final.head())
                st.info(f"พบข้อมูลทั้งหมด {len(df_final)} แถว")

                # 4. สร้างปุ่มสำหรับดาวน์โหลดไฟล์
                # แปลง DataFrame เป็นไฟล์ Excel ในหน่วยความจำ (in-memory)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='ProcessedData')
                
                # st.download_button ต้องการข้อมูลเป็น bytes
                processed_data = output.getvalue()
                
                new_filename = f"{sloc}_ZTRF.xlsx"
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ที่ประมวลผลแล้ว",
                    data=processed_data,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.warning("กรุณาอัปโหลดไฟล์และใส่รหัสคลังให้ครบถ้วน")