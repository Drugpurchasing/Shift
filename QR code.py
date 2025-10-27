import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import urllib.error

# --- ตารางแปลงหน่วย (Unit Mapping Dictionary) ---
# [จุดแก้ไข] ฝังตารางแปลงหน่วยไว้ในโค้ด
UNIT_MAPPING = {
    'Ampule': 'AMP',
    'Bag': 'G01',
    'Bottle': 'E01',
    'Box': 'B01',
    'Canister': 'CAN',
    'Capsule': 'CAP',
    'Each': 'EA',
    'Jar': 'J01',
    'Nebule': 'NEB',
    'Sache': 'SAC',
    'Sheet': 'SHT',
    'Suppo': 'SUP',
    'Syringe': 'SYR',
    'Vial': 'VIA',
    'Tablet': 'TAB',
    'Tube': 'TU1',
    'Unit': 'UNT'
}


# --- ฟังก์ชันสร้างฉลากยา (ไม่เปลี่ยนแปลง) ---
def create_drug_label(drug_name, drug_code, expiry_date, batch_no, quantity, unit_abbr):
    # ใช้ unit_abbr (หน่วยย่อ) ในการสร้าง QR code และแสดงผล
    qr_data = f"M|{drug_code}|{batch_no}|{quantity}|{unit_abbr}|{quantity}|{unit_abbr}"
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
    text_lines_info = [
        (drug_name, main_font), (drug_code, main_font),
        (f"{expiry_date}    {batch_no}", main_font), (f"{quantity} {unit_abbr}", main_font),  # แสดงหน่วยย่อ
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
    total_height = max(qr_img.height, total_text_height) + padding * 2
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
@st.cache_data(ttl=600)  # โหลดข้อมูลใหม่ทุก 10 นาที
def get_data_from_published_url(url):
    try:
        df = pd.read_csv(url)
        required_columns = ['Material', 'Material description', 'Sale Unit']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Error: ไม่พบคอลัมน์ที่จำเป็น ({', '.join(required_columns)}) ใน Google Sheet")
            return None
        df['Material'] = df['Material'].astype(str)
        return df
    except Exception as e:
        st.error(f"ไม่สามารถโหลดหรืออ่านข้อมูลจาก Google Sheet ได้: {e}")
        return None


# --- ส่วนของ Streamlit App ---
st.set_page_config(page_title="Drug Label Generator", layout="centered")
st.title("⚕️ Drug Label Generator")

# [จุดแก้ไข] - ฝัง URL ของ Google Sheet ที่ Publish ไว้ที่นี่
PUBLISHED_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQJpIKf_q4h4h1VEIM0tT1MlMvoEw1PXLYMxMv_c3abXFvAIBS0tWHxLL0sDjuuBrPjbrTP7lJH-NQw/pub?gid=0&single=true&output=csv"

# โหลดข้อมูลทันทีเมื่อเปิดแอป
drug_df = get_data_from_published_url(PUBLISHED_URL)

if drug_df is not None:
    st.success("ฐานข้อมูลยาพร้อมใช้งานแล้ว")

    # --- ฟอร์มสำหรับกรอกข้อมูล ---
    with st.form("drug_form"):
        st.subheader("กรอกข้อมูลเพื่อสร้างฉลาก")
        drug_code_input = st.text_input("รหัสยา (Material Code)")
        batch_no_input = st.text_input("เลขที่ผลิต (Batch No.)")
        expiry_date_input = st.text_input("วันหมดอายุ (Expiry Date, e.g., 29.02.2028)")
        quantity_input = st.number_input("จำนวน (Quantity)", min_value=1, value=1, step=1)

        submitted = st.form_submit_button("สร้างฉลากยา")

    if submitted and drug_code_input:
        drug_info = drug_df[drug_df['Material'] == drug_code_input]

        if not drug_info.empty:
            drug_data = drug_info.iloc[0]

            # [จุดแก้ไข] - แปลงหน่วยเต็มเป็นหน่วยย่อ
            full_unit = drug_data['Sale Unit']
            unit_abbreviation = UNIT_MAPPING.get(full_unit, full_unit)  # ถ้าไม่พบให้ใช้ชื่อเต็มแทน

            with st.spinner('กำลังสร้างฉลากยา...'):
                final_image = create_drug_label(
                    drug_name=drug_data['Material description'],
                    drug_code=drug_code_input,
                    expiry_date=expiry_date_input,
                    batch_no=batch_no_input,
                    quantity=quantity_input,
                    unit_abbr=unit_abbreviation  # ส่งหน่วยย่อไปใช้งาน
                )

            st.success("สร้างฉลากยาเรียบร้อย!")
            st.image(final_image, caption=f"ผลลัพธ์: {drug_data['Material description']}")

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
else:
    st.error("ไม่สามารถเริ่มต้นโปรแกรมได้ โปรดตรวจสอบ URL ของ Google Sheet ในโค้ด")