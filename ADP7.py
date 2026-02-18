import streamlit as st
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- 1. ฟังก์ชันเตรียมข้อมูล (Data Processing) ---
def process_dataframe(df, file_type):
    """
    ฟังก์ชันสำหรับคลีนข้อมูลตาม Logic เดิมของผู้ใช้
    """
    try:
        # Filter ตามเงื่อนไขเดิม
        if 'Flag Issue' in df.columns:
            df = df[df['Flag Issue'] != 'X']
        if 'M7 Log Exist' in df.columns:
            df = df[df['M7 Log Exist'] != 'X']
        
        processed_data = pd.DataFrame()

        # สร้าง Column Barcode ตาม Logic เดิม
        if file_type == "OPD":
            # Logic: 'O' + | + VN (4 digit) + | + Order Number
            vn_str = df['VN Number'].astype(str).str.zfill(4)
            order_str = df['Order Number'].astype(str)
            
            processed_data['barcode'] = 'O|' + vn_str + "|" + order_str
            processed_data['date'] = df['VN Date'].astype(str)
            processed_data['location'] = df['Storage location']

        elif file_type == "IPD":
            # Logic: 'i' + | + Admit Number + | + Order Number
            admit_str = df['Admit Number'].astype(str)
            order_str = df['Order Number'].astype(str)
            
            processed_data['barcode'] = 'i|' + admit_str + "|" + order_str
            processed_data['date'] = df['Order Date'].astype(str)
            processed_data['location'] = df['Storage location']
            
        return processed_data.drop_duplicates()

    except Exception as e:
        st.error(f"Error processing {file_type} data: {e}")
        return pd.DataFrame()

# --- 2. ฟังก์ชันเริ่มระบบ Automation (Selenium) ---
def run_automation(dataframe, user, password):
    # Setup Chrome Driver อัตโนมัติ (ไม่ต้องโหลดไฟล์เอง)
    try:
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        # options.add_argument("--headless") # เปิดบรรทัดนี้ถ้าไม่อยากให้เด้งหน้าต่าง Chrome ขึ้นมา
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 10) # รอสูงสุด 10 วินาทีในแต่ละขั้นตอน
    except Exception as e:
        st.error(f"ไม่สามารถเปิด Chrome ได้: {e}")
        return

    st.info("กำลังเปิด Browser...")
    
    # Progress Bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # 1. Login
        driver.get('http://172.16.61.11:8000/sap/bc/gui/sap/its/zismmhh0010?saml2=disabled')
        
        # รอให้ช่อง Username โผล่มาแล้วค่อยพิมพ์ (แทน time.sleep)
        user_field = wait.until(EC.presence_of_element_located((By.XPATH, '//input[contains(@name, "sap-user")]')))
        pass_field = driver.find_element(By.XPATH, '//input[contains(@name, "sap-password")]')
        
        user_field.send_keys(user)
        pass_field.send_keys(password)
        pass_field.send_keys(Keys.ENTER) # Submit

        # 2. Navigate Menu
        # รอหน้าโหลด แล้วกด m4[1]
        m4_input = wait.until(EC.presence_of_element_located((By.NAME, 'm4[1]')))
        m4_input.send_keys(Keys.ENTER)

        # รอหน้าโหลด แล้วกด m3[1]
        m3_input = wait.until(EC.presence_of_element_located((By.NAME, 'm3[1]')))
        m3_input.send_keys(Keys.ENTER)

        # 3. Loop Data
        total_rows = len(dataframe)
        
        # หา Element ที่เป็นช่อง Input ข้อมูล (ปรับ Xpath ให้แม่นยำขึ้นจากโค้ดเดิม)
        # หมายเหตุ: Xpath ของ SAP ITS Mobile มักจะเปลี่ยนถ้าหน้าจอเปลี่ยน 
        # แนะนำให้ใช้ ID หรือ Name ถ้าหาเจอ แต่ถ้าไม่มีใช้ Xpath เดิมที่เคยใช้ได้
        
        # รอให้หน้าจอ input พร้อม
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]')))

        for index, row in dataframe.iterrows():
            try:
                # Update progress
                status_text.text(f"Processing row {index + 1}/{total_rows}: {row['barcode']}")
                progress_bar.progress((index + 1) / total_rows)

                # เตรียมข้อมูลวันที่ (Format เดิม: 0:4 + 5:7 + 8:10 -> YYYYMMDD)
                raw_date = str(row['date'])
                if len(raw_date) >= 10:
                    formatted_date = raw_date[0:4] + raw_date[5:7] + raw_date[8:10]
                else:
                    formatted_date = raw_date # Fallback
                
                input_str = f"{row['barcode']}|{formatted_date}"
                
                # --- Fill Form ---
                # ช่อง Barcode/Data
                field_barcode = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]')
                field_barcode.clear()
                field_barcode.send_keys(input_str)
                
                # ช่อง Location
                field_loc = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[11]')
                field_loc.clear()
                field_loc.send_keys(str(row['location']))
                
                # กด Submit (Input 3 ในโค้ดเดิม)
                btn_submit = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[3]')
                btn_submit.click()
                
                # --- Handle Popups / Errors ---
                try:
                    # เช็คว่ามี Popup ให้กดเลือก option หรือไม่
                    popup_opt = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.NAME, "spop-option1[1]"))
                    )
                    popup_opt.click()
                except TimeoutException:
                    pass # ไม่มี Popup ก็ทำต่อ

                # เคลียร์ค่าเพื่อเตรียมรอบถัดไป (Logic เดิมมีการกด Tab แต่เราใช้ clear() ปลอดภัยกว่า)
                try:
                    field_barcode = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]')
                    field_barcode.clear()
                except:
                    pass

            except Exception as row_e:
                st.warning(f"Row {index} failed: {row_e}")
                # พยายามกลับไปหน้าเดิมหรือเคลียร์หน้าจอถ้าจำเป็น
                continue

        st.success("ทำงานเสร็จสิ้น!")
        time.sleep(5) # เปิดค้างไว้แป๊บนึงก่อนปิด
        driver.quit()

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดระหว่างรัน Automation: {e}")
        if 'driver' in locals():
            driver.quit()

# --- 3. ส่วนหน้าจอ UI (Streamlit Layout) ---
st.set_page_config(page_title="Auto Dispense V7", page_icon="💊")

st.title("💊 Auto Dispense V7 (Web Version)")
st.markdown("โปรแกรมช่วยคีย์ข้อมูล SAP อัตโนมัติ")

with st.sidebar:
    st.header("การตั้งค่า")
    user_in = st.text_input("Username", placeholder="SAP Username")
    pass_in = st.text_input("Password", type="password")
    
    st.divider()
    mode = st.radio("เลือกโหมดการทำงาน", ["OPD", "IPD", "BOTH"])

# Main Area
col1, col2 = st.columns(2)
opd_df = pd.DataFrame()
ipd_df = pd.DataFrame()

# File Uploader
if mode in ["OPD", "BOTH"]:
    with col1:
        st.subheader("ไฟล์ OPD")
        opd_file = st.file_uploader("Upload OPD.xlsx", type=['xlsx'])
        if opd_file:
            raw_opd = pd.read_excel(opd_file)
            opd_df = process_dataframe(raw_opd, "OPD")
            st.write(f"พบข้อมูล {len(opd_df)} รายการ")

if mode in ["IPD", "BOTH"]:
    with col2:
        st.subheader("ไฟล์ IPD")
        ipd_file = st.file_uploader("Upload IPD.xlsx", type=['xlsx'])
        if ipd_file:
            raw_ipd = pd.read_excel(ipd_file)
            ipd_df = process_dataframe(raw_ipd, "IPD")
            st.write(f"พบข้อมูล {len(ipd_df)} รายการ")

# รวมข้อมูล
final_df = pd.concat([opd_df, ipd_df], ignore_index=True)

if not final_df.empty:
    st.divider()
    st.subheader("ตัวอย่างข้อมูลที่จะรัน")
    st.dataframe(final_df.head())
    
    if st.button("🚀 เริ่มทำงาน (Run)", type="primary"):
        if not user_in or not pass_in:
            st.warning("กรุณากรอก Username และ Password")
        else:
            run_automation(final_df, user_in, pass_in)
else:
    st.info("กรุณาอัปโหลดไฟล์ข้อมูลเพื่อเริ่มทำงาน")