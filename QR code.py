import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import urllib.error

# --- ตารางแปลงหน่วย (Unit Mapping Dictionary) ---
UNIT_MAPPING = {
    'Ampule': 'AMP', 'Bag': 'G01', 'Bottle': 'E01', 'Box': 'B01',
    'Canister': 'CAN', 'Capsule': 'CAP', 'Each': 'EA', 'Jar': 'J01',
    'Nebule': 'NEB', 'Sache': 'SAC', 'Sheet': 'SHT', 'Suppo': 'SUP',
    'Syringe': 'SYR', 'Vial': 'VIA', 'Tablet': 'TAB', 'Tube': 'TU1',
    'Unit': 'UNT'
}


# --- ฟังก์ชันสร้างฉลากยา (ปรับแก้สัดส่วนและการจัดวางข้อความ) ---
def create_drug_label(drug_name, drug_code, expiry_date, batch_no, quantity, unit_abbr):
    qr_data = f"M|{drug_code}|{batch_no}|{quantity}|{unit_abbr}|{quantity}|{unit_abbr}"
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')

    try:
        main_font = ImageFont.truetype("arial.ttf", 36)  # ฟอนต์หลัก
        small_font = ImageFont.truetype("arial.ttf", 30)  # ฟอนต์สำหรับ CRA
        batch_expiry_font = ImageFont.truetype("arial.ttf", 36)  # ฟอนต์สำหรับวันหมดอายุ/Batch
        quantity_font = ImageFont.truetype("arial.ttf", 36)  # ฟอนต์สำหรับจำนวน
    except IOError:
        st.warning("ไม่พบฟอนต์ 'arial.ttf', โปรดตรวจสอบว่ามีฟอนต์นี้ในระบบ หรือติดตั้ง 'ms-fonts' บน Linux")
        main_font = ImageFont.load_default()
        small_font = ImageFont.load_default()
        batch_expiry_font = ImageFont.load_default()
        quantity_font = ImageFont.load_default()

    # [จุดแก้ไขหลัก] - จัดเรียงข้อความแยกบรรทัดตามรูปภาพต้นฉบับ
    # แทนที่จะเป็น list เดียว เราจะจัดตำแหน่งด้วยตนเอง

    # --- กำหนดขนาดพื้นฐานของภาพ ---
    qr_width = qr_img.width
    qr_height = qr_img.height

    # กำหนดค่าคงที่สำหรับการจัดวาง
    padding_x = 20  # padding ซ้าย-ขวา โดยรวม
    padding_y = 20  # padding บน-ล่าง โดยรวม
    cra_text_width = 40  # ความกว้างของพื้นที่ CRA แนวตั้ง
    spacing_qr_to_text = 20  # ระยะห่างจาก QR code ถึงข้อความ
    line_height_multiplier = 1.2  # ตัวคูณความสูงบรรทัด (เพื่อให้มีช่องว่างระหว่างบรรทัด)

    # กำหนดความกว้างของฉลากโดยประมาณจากรูปต้นฉบับ (QR + spacing + ข้อความ)
    # เราจะให้ความกว้างของข้อความเป็นประมาณ 2 เท่าของ QR Code
    estimated_text_area_width = qr_width * 2

    total_width = cra_text_width + qr_width + spacing_qr_to_text + estimated_text_area_width + padding_x
    total_height = qr_height + (padding_y * 2)  # ให้สูงเท่า QR + padding

    canvas = Image.new('RGB', (total_width, total_height), 'white')
    draw = ImageDraw.Draw(canvas)

    # 1. วาดข้อความ "CRA" แนวตั้ง
    cra_text_img = Image.new('RGB', (50, 200), 'white')
    cra_draw = ImageDraw.Draw(cra_text_img)
    cra_draw.text((0, 0), "CRA", font=small_font, fill="black")
    cra_text_img = cra_text_img.rotate(90, expand=True)
    y_pos_cra = int((total_height - cra_text_img.height) / 2)
    canvas.paste(cra_text_img, (5, y_pos_cra))

    # 2. วาง QR Code
    x_pos_qr = cra_text_width
    y_pos_qr = int((total_height - qr_height) / 2)  # จัด QR ให้อยู่กึ่งกลางแนวตั้ง
    canvas.paste(qr_img, (x_pos_qr, y_pos_qr))

    # 3. วาดข้อความข้อมูลยา
    x_pos_text_start = x_pos_qr + qr_width + spacing_qr_to_text

    # คำนวณความสูงของแต่ละบรรทัด (เพื่อให้จัดตำแหน่งได้)
    # [จุดแก้ไข] - ใช้ getbbox เพื่อความแม่นยำในการวัดฟอนต์
    def get_text_height(text, font):
        try:
            bbox = font.getbbox(text)
            return bbox[3] - bbox[1]
        except AttributeError:  # Fallback for older Pillow versions
            _, h = font.getsize(text)
            return h

    # บรรทัดที่ 1: ชื่อยา (Avastin (Bevacizumab) 100 mg/4 mL)
    text1_y = padding_y
    draw.text((x_pos_text_start, text1_y), drug_name, font=main_font, fill="black")

    # บรรทัดที่ 2: inj. (หรือ formulation, แต่ตอนนี้เราไม่มีแล้ว)
    # ถ้าอยากให้มีคำว่า inj. ต้องเพิ่มเข้าไปใน drug_name หรือกำหนดแยก
    # จากตัวอย่างรูปภาพ "inj." อยู่บรรทัดถัดจากชื่อยา
    # เพื่อให้เหมือนรูป ผมจะแยก "inj." ออกจาก drug_name และแสดงเป็นบรรทัดใหม่
    # หรือถ้า drug_name มี inj. อยู่แล้วก็ไม่จำเป็นต้องเพิ่มแยกอีก
    # ในกรณีนี้ผมจะสมมติว่า drug_name มี "inj." รวมอยู่แล้ว หรือไม่ต้องแสดง
    # ถ้าต้องการแสดง 'inj.' แยก ให้เพิ่มพารามิเตอร์ formulation กลับมา หรือจัดการกับ drug_name ให้ดี

    # สำหรับตอนนี้ ผมจะเว้นช่องว่างเล็กน้อยเหมือนมีบรรทัด "inj." แต่ไม่แสดงข้อความ
    # หรือถ้า drug_name ไม่มี "inj." และอยากได้ "inj."
    # อาจจะต้องปรับ drug_name ให้เป็น "Avastin (Bevacizumab) 100 mg/4 mL\ninj." แล้วใช้ multiline text
    # แต่จากโค้ดเดิมเราส่ง drug_name มาเป็น string เดียว

    # ลองปรับความสูงบรรทัดที่ 2 เพื่อให้เหมือนรูปมากขึ้น
    # [จุดแก้ไข] - ปรับตำแหน่งบรรทัด
    text_height_name = get_text_height(drug_name, main_font)

    # บรรทัดที่ 3: รหัสยา (1200000639)
    # ให้รหัสยาอยู่ใต้ชื่อยาโดยมีช่องว่าง
    text2_y = text1_y + text_height_name + (line_height_multiplier * get_text_height(" ", main_font))
    draw.text((x_pos_text_start, text2_y), drug_code, font=main_font, fill="black")

    # บรรทัดที่ 4: วันหมดอายุ (29.02.2028) และ Batch No. (H7911B02U1)
    # [จุดแก้ไข] - จัดวางให้อยู่บรรทัดเดียวกัน
    text3_y = text2_y + get_text_height(drug_code, main_font) + (
                line_height_multiplier * get_text_height(" ", batch_expiry_font))
    expiry_text = expiry_date
    batch_text = batch_no

    # คำนวณตำแหน่งของ batch_no เพื่อให้ชิดขวาของพื้นที่ข้อความ
    text_expiry_width = get_text_height(expiry_text, batch_expiry_font)
    text_batch_width = get_text_height(batch_text, batch_expiry_font)

    # [จุดแก้ไข] - วาง expiry_date ชิดซ้าย และ batch_no ชิดขวาในพื้นที่ข้อความ
    # โดยอ้างอิงจากขอบขวาของ Canvas - padding - ความกว้างของ batch_no
    draw.text((x_pos_text_start, text3_y), expiry_text, font=batch_expiry_font, fill="black")

    # ตำแหน่ง X สำหรับ Batch No. (ชิดขวาของพื้นที่ข้อความ)
    x_pos_batch_no = total_width - padding_x - get_text_height(batch_text, batch_expiry_font)
    draw.text((x_pos_batch_no, text3_y), batch_text, font=batch_expiry_font, fill="black")

    # บรรทัดที่ 5: จำนวนและหน่วย (1 Vial)
    # [จุดแก้ไข] - จัดวาง
    text4_y = text3_y + get_text_height(expiry_text, batch_expiry_font) + (
                line_height_multiplier * get_text_height(" ", quantity_font))
    draw.text((x_pos_text_start, text4_y), f"{quantity} {unit_abbr}", font=quantity_font, fill="black")

    return canvas


# --- ฟังก์ชันดึงข้อมูลจาก URL (ไม่มีการเปลี่ยนแปลง) ---
@st.cache_data(ttl=600)
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
PUBLISHED_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQJpIKf_q4h4h1VEIM0tT1MlMvoEw1PXLYMxMv_c3abXFvAIBS0tWHxLL0sDjuuBrPjbrTP7lJH-NQw/pub?gid=0&single=true&output=csv"  # <--- แก้ไข URL ของคุณที่นี่ !!!

drug_df = get_data_from_published_url(PUBLISHED_URL)

if drug_df is not None:
    st.success("ฐานข้อมูลยาพร้อมใช้งานแล้ว")

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

            full_unit = drug_data['Sale Unit']
            unit_abbreviation = UNIT_MAPPING.get(full_unit, full_unit)

            with st.spinner('กำลังสร้างฉลากยา...'):
                final_image = create_drug_label(
                    drug_name=drug_data['Material description'],
                    drug_code=drug_code_input,
                    expiry_date=expiry_date_input,
                    batch_no=batch_no_input,
                    quantity=quantity_input,
                    unit_abbr=unit_abbreviation
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
