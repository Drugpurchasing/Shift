import streamlit as st
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO


def create_drug_label(
        drug_name,
        formulation,
        drug_code,
        expiry_date,
        batch_no,
        quantity,
        unit
):
    """
    ฟังก์ชันสำหรับสร้างภาพฉลากยาที่มีทั้ง QR Code และข้อความประกอบ
    """
    # 1. สร้างข้อมูลสำหรับ QR Code
    qr_data = f"M|{drug_code}|{batch_no}|{quantity}|{unit}|{quantity}|{unit}"

    # 2. สร้าง QR Code
    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')

    # 3. เตรียม Font (ใช้ฟอนต์เริ่มต้นของ PIL เพื่อความเข้ากันได้)
    try:
        # พยายามใช้ฟอนต์ที่อาจมีอยู่ แต่ถ้าไม่มีให้ใช้ฟอนต์ดีฟอลต์
        main_font = ImageFont.truetype("arial.ttf", 36)
        cra_font = ImageFont.truetype("arial.ttf", 30)
    except IOError:
        # ฟอนต์ดีฟอลต์จะทำงานได้ในทุกระบบ
        main_font = ImageFont.load_default()
        cra_font = ImageFont.load_default()

    # 4. จัดเรียงและคำนวณขนาดข้อความ
    text_lines_info = [
        (drug_name, main_font), (formulation, main_font), (drug_code, main_font),
        (f"{expiry_date}    {batch_no}", main_font), (f"{quantity} {unit}", main_font),
    ]

    max_text_width = 0
    total_text_height = 0
    line_spacing = 10
    for text, font in text_lines_info:
        try:
            bbox = font.getbbox(text)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except AttributeError:
            text_width, text_height = font.getsize(text)

        if text_width > max_text_width:
            max_text_width = text_width
        total_text_height += text_height + line_spacing

    # 5. สร้างภาพและประกอบชิ้นส่วน
    padding = 20
    vertical_text_width = 40
    qr_text_spacing = 20  # เพิ่มระยะห่าง

    total_width = vertical_text_width + qr_img.width + qr_text_spacing + max_text_width + padding
    total_height = max(qr_img.height, total_text_height) + (padding * 2)

    canvas = Image.new('RGB', (total_width, total_height), 'white')
    draw = ImageDraw.Draw(canvas)

    # วาดข้อความ "CRA"
    cra_text_img = Image.new('RGB', (50, 200), 'white')
    cra_draw = ImageDraw.Draw(cra_text_img)
    cra_draw.text((0, 0), "CRA", font=cra_font, fill="black")
    cra_text_img = cra_text_img.rotate(90, expand=True)
    y_pos_cra = int((total_height - cra_text_img.height) / 2)
    canvas.paste(cra_text_img, (5, y_pos_cra))

    # วาง QR Code
    x_pos_qr = vertical_text_width
    y_pos_qr = int((total_height - qr_img.height) / 2)
    canvas.paste(qr_img, (x_pos_qr, y_pos_qr))

    # วาดข้อความข้อมูลยา
    x_pos_text = x_pos_qr + qr_img.width + qr_text_spacing
    current_y = padding
    for text, font in text_lines_info:
        draw.text((x_pos_text, current_y), text, font=font, fill="black")
        try:
            bbox = font.getbbox(text)
            current_y += (bbox[3] - bbox[1]) + line_spacing
        except AttributeError:
            _, text_height = font.getsize(text)
            current_y += text_height + line_spacing

    return canvas


# --- ส่วนของ Streamlit App ---
st.set_page_config(page_title="Drug Label Generator", layout="centered")
st.title("⚕️ Drug Label & QR Code Generator")
st.write("กรอกข้อมูลยาเพื่อสร้างฉลากพร้อม QR Code ตามรูปแบบที่กำหนด")

with st.form("drug_form"):
    st.subheader("ข้อมูลยา (Drug Information)")

    # แบ่งหน้าจอเป็น 2 คอลัมน์
    col1, col2 = st.columns(2)

    with col1:
        drug_name = st.text_input("ชื่อยา (Drug Name)", "Avastin (Bevacizumab) 100 mg/4 mL")
        formulation = st.text_input("รูปแบบยา (Formulation)", "inj.")
        drug_code = st.text_input("รหัสยา (Drug Code)", "1200000639")
        expiry_date = st.text_input("วันหมดอายุ (Expiry Date)", "29.02.2028")

    with col2:
        batch_no = st.text_input("เลขที่ผลิต (Batch No.)", "H7911B02U1")
        quantity = st.number_input("จำนวน (Quantity)", min_value=1, value=1, step=1)
        unit = st.text_input("หน่วย (Unit)", "Vial")

    submitted = st.form_submit_button("สร้างฉลากยา (Generate Label)")

if submitted:
    # ตรวจสอบว่ากรอกข้อมูลครบหรือไม่
    if not all([drug_name, drug_code, expiry_date, batch_no, unit]):
        st.error("กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน")
    else:
        with st.spinner('กำลังสร้างฉลากยา...'):
            # เรียกฟังก์ชันเพื่อสร้างภาพ
            final_image = create_drug_label(
                drug_name, formulation, drug_code, expiry_date, batch_no, quantity, unit
            )

            st.success("สร้างฉลากยาเรียบร้อยแล้ว!")
            st.image(final_image, caption="ผลลัพธ์ฉลากยา", use_column_width=True)

            # เตรียมไฟล์สำหรับดาวน์โหลด
            buf = BytesIO()
            final_image.save(buf, format="PNG")
            byte_im = buf.getvalue()

            st.download_button(
                label="📥 ดาวน์โหลดรูปภาพ (Download Image)",
                data=byte_im,
                file_name=f"{drug_code}_{batch_no}.png",
                mime="image/png"
            )