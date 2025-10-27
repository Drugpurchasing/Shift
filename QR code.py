import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import urllib.error


# --- ฟังก์ชันสร้างฉบากยา (ไม่มีการเปลี่ยนแปลง) ---
def create_drug_label(drug_name, drug_code, expiry_date, batch_no, quantity, unit):
    # ฟังก์ชันนี้เหมือนเดิม แต่จะไม่มีพารามิเตอร์ formulation แล้ว
    # ... (โค้ดส่วนนี้เหมือนเดิมทุกประการ) ...
    qr_data = f"M|{drug_code}|{batch_no}|{quantity}|{unit}|{quantity}|{unit}"
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    try:
        main_font = ImageFont.truetype("arial.ttf", 36)
        cra_font = ImageFont.truetype("arial.ttf", 30)
    except IOError:
        main_font = ImageFont.load_default()
        cra_font = ImageFont.load_default()
    # [จุดที่แก้ไข] - เราจะไม่มี formulation ใน list นี้แล้ว
    text_lines_info = [
        (drug_name, main_font), (drug_code, main_font),
        (f"{expiry_date}    {batch_no}", main_font), (f"{quantity} {unit}", main_font),
    ]
    max_text_width = 0
    total_text_height = 0
    line_spacing = 10
    for text, font in text_lines_info:
        try:
            bbox = font.getbbox(text)
            text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]
        except AttributeError:
            text_width, text_height = font.getsize(text)
        if text_width > max_text_width: max_text_width = text_width
        total_text_height += text_height + line_spacing
    padding, vertical_text_width, qr_text_spacing = 20, 40, 20
    total_width = vertical_text_width + qr_img.width + qr_text_spacing + max_text_width + padding
    total_height = max(qr_img.height, total_text_height) + padding * 2  # ปรับการคำนวณเล็กน้อย
    canvas = Image.new('RGB', (total_width, total_height), 'white')
    draw = ImageDraw.Draw(canvas)
    cra_text_img = Image.new('RGB', (50, 200), 'white')
    cra_draw = ImageDraw.Draw(cra_text_img)
    cra_draw.text((0, 0), "CRA", font=cra_font, fill="black")
    cra_text_img = cra_text_img.rotate(90, expand=True)
    canvas.paste(cra_text_img, (5, int((total_height - cra_text_img.height) / 2)))
    canvas.paste(qr_img, (vertical_text_width, int((total_height - qr_img.height) / 2)))
    x_pos_text, current_y = vertical_text_width + qr_img.width + qr_text_spacing, padding
    for text, font in text_lines_info:
        draw.text((x_pos_text, current_y), text, font=font, fill="black")
        try:
            bbox = font.getbbox(text)
            current_y += (bbox[3] - bbox[1]) + line_spacing
        except AttributeError:
            _, text_height = font.getsize(text)
            current_y += text_height + line_spacing
    return canvas


# --- ฟังก์ชันดึงข้อมูลจาก URL ---
@st.cache_data(ttl=600)
def get_data_from_published_url(url):
    try:
        df = pd.read_csv(url)
        # [จุดที่แก้ไข] - เปลี่ยนชื่อคอลัมน์ให้ตรงกับชีทของคุณ
        # ตรวจสอบว่าคอลัมน์ที่จำเป็นมีอยู่หรือไม่
        required_columns = ['Material', 'Material description', 'Sale Unit']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Error: ไม่พบคอลัมน์ที่จำเป็นใน Google Sheet")
            st.info(f"โปรดตรวจสอบว่าไฟล์ CSV ของคุณมีคอลัมน์: {', '.join(required_columns)}")
            return None

        # [จุดที่แก้ไข] - ทำให้คอลัมน์ Material (รหัสยา) เป็น string
        df['Material'] = df['Material'].astype(str)
        return df
    except urllib.error.URLError:
        st.error("Connection Error: ไม่สามารถเข้าถึง URL ได้ โปรดตรวจสอบ Link และการเชื่อมต่ออินเทอร์เน็ต")
        return None
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านข้อมูล: {e}")
        return None


# --- ส่วนของ Streamlit App ---
st.set_page_config(page_title="Drug Label Generator", layout="wide")
st.title("⚕️ Drug Label Generator (from Published Google Sheet)")

# --- 1. ตั้งค่า URL ---
st.subheader("1. Connection Setup")
published_url = st.text_input(
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQJpIKf_q4h4h1VEIM0tT1MlMvoEw1PXLYMxMv_c3abXFvAIBS0tWHxLL0sDjuuBrPjbrTP7lJH-NQw/pub?gid=0&single=true&output=csv",
    help="ต้องเป็น URL ที่ได้จาก File -> Share -> Publish to web และเลือก output เป็น .csv"
)

if published_url:
    drug_df = get_data_from_published_url(published_url)

    if drug_df is not None:
        st.success("อ่านข้อมูลจาก Google Sheet สำเร็จ!")

        # --- 2. ฟอร์มสำหรับกรอกข้อมูล ---
        st.subheader("2. Generate Label")
        with st.form("drug_form"):
            col1, col2 = st.columns(2)
            with col1:
                # [จุดที่แก้ไข] - เปลี่ยน Label ให้สอดคล้องกัน
                drug_code_input = st.text_input("รหัสยา (Material Code)")
                batch_no_input = st.text_input("เลขที่ผลิต (Batch No.)")
            with col2:
                expiry_date_input = st.text_input("วันหมดอายุ (Expiry Date, e.g., 29.02.2028)")
                quantity_input = st.number_input("จำนวน (Quantity)", min_value=1, value=1, step=1)

            submitted = st.form_submit_button("สร้างฉลากยา")

        if submitted and drug_code_input:
            # [จุดที่แก้ไข] - ค้นหาข้อมูลโดยใช้ชื่อคอลัมน์ 'Material'
            drug_info = drug_df[drug_df['Material'] == drug_code_input]

            if not drug_info.empty:
                drug_data = drug_info.iloc[0]
                with st.spinner('กำลังสร้างฉลากยา...'):
                    # [จุดที่แก้ไข] - ส่งข้อมูลไปยังฟังก์ชันโดยใช้ชื่อคอลัมน์ใหม่
                    final_image = create_drug_label(
                        drug_name=drug_data['Material description'],  # <-- ชื่อยา
                        drug_code=drug_code_input,
                        expiry_date=expiry_date_input,
                        batch_no=batch_no_input,
                        quantity=quantity_input,
                        unit=drug_data['Sale Unit']  # <-- หน่วย
                    )

                st.success("สร้างฉลากยาเรียบร้อยแล้ว!")
                st.image(final_image, caption="ผลลัพธ์ฉลากยา")

                buf = BytesIO()
                final_image.save(buf, format="PNG")
                st.download_button(
                    label="📥 ดาวน์โหลดรูปภาพ",
                    data=buf.getvalue(),
                    file_name=f"label_{drug_code_input}_{batch_no_input}.png",
                    mime="image/png"
                )
            else:
                st.error(f"ไม่พบรหัสยา '{drug_code_input}' ในฐานข้อมูล")