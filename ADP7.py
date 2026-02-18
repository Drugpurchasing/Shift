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

# ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(page_title="Auto Dispense V7", page_icon="💊", layout="wide")

# --- 1. ฟังก์ชันเตรียมข้อมูล (Data Processing) ---
def process_dataframe(df, file_type):
    try:
        # กรองข้อมูลตามเงื่อนไข (Clean Data)
        if 'Flag Issue' in df.columns:
            df = df[df['Flag Issue'] != 'X']
        if 'M7 Log Exist' in df.columns:
            df = df[df['M7 Log Exist'] != 'X']
        
        processed_data = pd.DataFrame()

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
            
        # คืนค่าข้อมูลที่ตัดตัวซ้ำออกแล้ว
        return processed_data.drop_duplicates()

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file_type}: {e}")
        return pd.DataFrame()

# --- 2. ฟังก์ชันเริ่มระบบ Automation (Selenium) ---
def run_automation(dataframe, user, password, show_browser):
    driver = None
    try:
        # Setup Chrome Driver
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        
        # เลือกโหมดแสดงผล Browser (ถ้าไม่ติ๊ก Show Browser จะรันแบบ Headless)
        if not show_browser:
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
        
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--start-maximized")

        st.info("กำลังเปิด Google Chrome...")
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 15)  # รอสูงสุด 15 วินาที

        # ---------------------------------------------------------
        # เริ่มขั้นตอนการทำงานกับ SAP
        # ---------------------------------------------------------
        
        # 1. เข้าสู่เว็บไซต์ (Login)
        target_url = 'http://172.16.61.11:8000/sap/bc/gui/sap/its/zismmhh0010?saml2=disabled'
        st.write(f"กำลังเชื่อมต่อ: {target_url}")
        driver.get(target_url)

        # รอช่อง User/Pass และ Login
        wait.until(EC.presence_of_element_located((By.XPATH, '//input[contains(@name, "sap-user")]'))).send_keys(user)
        pwd_box = driver.find_element(By.XPATH, '//input[contains(@name, "sap-password")]')
        pwd_box.send_keys(password)
        pwd_box.send_keys(Keys.ENTER)

        # 2. นำทางเมนู (Menu Navigation)
        st.write("กำลังเข้าสู่เมนู...")
        wait.until(EC.presence_of_element_located((By.NAME, 'm4[1]'))).send_keys(Keys.ENTER) # เมนู 1
        wait.until(EC.presence_of_element_located((By.NAME, 'm3[1]'))).send_keys(Keys.ENTER) # เมนู 2

        # 3. เริ่มวนลูปข้อมูล (Loop Data)
        st.divider()
        st.subheader("สถานะการทำงาน")
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_area = st.empty()
        
        total_rows = len(dataframe)
        success_count = 0
        fail_count = 0

        # รอหน้าฟอร์มพร้อมก่อนเริ่มลูป
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]')))

        for index, row in dataframe.iterrows():
            try:
                current_barcode = row['barcode']
                status_text.text(f"กำลังทำรายการที่ {index + 1}/{total_rows}: {current_barcode}")
                progress_bar.progress((index + 1) / total_rows)

                # แปลงวันที่
                raw_date = str(row['date'])
                if len(raw_date) >= 10:
                    fmt_date = raw_date[0:4] + raw_date[5:7] + raw_date[8:10]
                else:
                    fmt_date = raw_date
                
                input_str = f"{current_barcode}|{fmt_date}"

                # กรอกข้อมูล (Fill Form)
                # ช่อง Barcode
                inp_barcode = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]')
                inp_barcode.clear()
                inp_barcode.send_keys(input_str)

                # ช่อง Location
                inp_loc = driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[11]')
                inp_loc.clear()
                inp_loc.send_keys(str(row['location']))

                # กดปุ่มยืนยัน (Submit)
                driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[3]').click()

                # จัดการ Popup (ถ้ามี)
                try:
                    popup = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.NAME, "spop-option1[1]"))
                    )
                    popup.click()
                except TimeoutException:
                    pass # ไม่มี Popup

                # เคลียร์ค่าเตรียมรอบถัดไป
                try:
                    driver.find_element(By.XPATH, '//*[@id="mobileform"]/div[2]/input[6]').clear()
                except:
                    pass
                
                success_count += 1

            except Exception as e:
                fail_count += 1
                st.warning(f"Error รายการที่ {index + 1}: {e}")
                # พยายาม Reset กลับมาหน้าเดิมถ้าพัง
                continue
        
        progress_bar.progress(100)
        st.success(f"✅ เสร็จสิ้น! สำเร็จ: {success_count} | ผิดพลาด: {fail_count}")
        st.balloons()

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดร้ายแรง (Critical Error): {e}")
        st.error("คำแนะนำ: โปรดตรวจสอบว่า Username/Password ถูกต้อง และเครื่องคอมพิวเตอร์เชื่อมต่อเครือข่ายโรงพยาบาลอยู่")
    
    finally:
        if driver:
            time.sleep(5) # รอให้ดูผลลัพธ์แป๊บนึง
            driver.quit()
            st.info("ปิด Browser แล้ว")

# --- 3. ส่วนหน้าจอ UI (Streamlit Layout) ---
st.title("💊 Auto Dispense V7")
st.markdown("ระบบช่วยคีย์ข้อมูล SAP อัตโนมัติ (รันบน Local Network)")

with st.sidebar:
    st.header("🔑 เข้าสู่ระบบ SAP")
    user_in = st.text_input("Username", placeholder="กรอก Username")
    pass_in = st.text_input("Password", type="password", placeholder="กรอก Password")
    
    st.divider()
    st.header("⚙️ ตั้งค่า")
    mode = st.radio("เลือกโหมดข้อมูล", ["OPD", "IPD", "BOTH"])
    show_browser = st.checkbox("แสดงหน้าจอ Chrome ขณะทำงาน", value=True)

# Main Content
col1, col2 = st.columns(2)
opd_df = pd.DataFrame()
ipd_df = pd.DataFrame()

# ส่วนอัปโหลดไฟล์
if mode in ["OPD", "BOTH"]:
    with col1:
        st.subheader("📄 ไฟล์ OPD")
        opd_file = st.file_uploader("ลากไฟล์ OPD.xlsx มาวางที่นี่", type=['xlsx'])
        if opd_file:
            raw_opd = pd.read_excel(opd_file)
            opd_df = process_dataframe(raw_opd, "OPD")
            st.success(f"OPD: {len(opd_df)} รายการ")

if mode in ["IPD", "BOTH"]:
    with col2:
        st.subheader("📄 ไฟล์ IPD")
        ipd_file = st.file_uploader("ลากไฟล์ IPD.xlsx มาวางที่นี่", type=['xlsx'])
        if ipd_file:
            raw_ipd = pd.read_excel(ipd_file)
            ipd_df = process_dataframe(raw_ipd, "IPD")
            st.success(f"IPD: {len(ipd_df)} รายการ")

# รวมข้อมูลเพื่อเตรียมรัน
final_df = pd.concat([opd_df, ipd_df], ignore_index=True)

if not final_df.empty:
    st.divider()
    st.subheader(f"📊 พร้อมทำงาน: ทั้งหมด {len(final_df)} รายการ")
    with st.expander("ดูตัวอย่างข้อมูล (คลิกเพื่อขยาย)"):
        st.dataframe(final_df.head(10))
    
    start_btn = st.button("🚀 เริ่มทำงาน (Start Automation)", type="primary", use_container_width=True)
    
    if start_btn:
        if not user_in or not pass_in:
            st.warning("⚠️ กรุณากรอก Username และ Password ก่อนเริ่มทำงาน")
        else:
            run_automation(final_df, user_in, pass_in, show_browser)
else:
    st.info("👆 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้น")