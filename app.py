import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
import datetime

# --- ฟังก์ชันแปลงวันที่เป็นภาษาไทย ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def to_thai_datetime(dt, is_time_included=True):
    if not dt: return ""
    base_format = f"วันที่ {dt.day} {THAI_MONTHS[dt.month]} {dt.year + 543}"
    if is_time_included:
        return f"{base_format} เวลาประมาณ {dt.strftime('%H:%M')} น."
    return base_format

# --- ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="ระบบบันทึกจับกุม", layout="wide")
st.title("แบบฟอร์มบันทึกจับกุม")

# --- 1. ข้อมูลสถานที่และเวลา ---
st.header("1. ข้อมูลสถานที่และเวลา")
col1, col2 = st.columns(2)
with col1:
    record_location = st.text_input("สถานที่ทำการบันทึก", value="กองกำกับการ 3 กองบังคับการปราบปราม")
    arrest_location = st.text_area("สถานที่จับกุม", height=100)
with col2:
    record_date = st.date_input("วันที่บันทึก", datetime.date.today())
    record_time = st.time_input("เวลาที่บันทึก", datetime.datetime.now().time())
    arrest_date = st.date_input("วันที่จับกุม", datetime.date.today())
    arrest_time = st.time_input("เวลาที่จับกุม", datetime.datetime.now().time())

record_dt = datetime.datetime.combine(record_date, record_time)
arrest_dt = datetime.datetime.combine(arrest_date, arrest_time)
st.divider()

# --- 2. หน่วยจับกุมและเจ้าหน้าที่ ---
st.header("2. หน่วยจับกุมและเจ้าหน้าที่ (ข้อ 5, 6, 7)")
if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 1

def add_unit():
    st.session_state.unit_count += 1

units_data = []
default_commander = "พล.ต.ท.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พล.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."

for i in range(st.session_state.unit_count):
    with st.expander(f"🏢 หน่วยจับกุมที่ {i+1}", expanded=True):
        unit_name = st.text_input(f"ชื่อหน่วยงานที่ {i+1}", value="กก.๓ บก.ป." if i==0 else "", key=f"unit_name_{i}")
        commanders_text = st.text_area(f"ภายใต้อำนวยการสั่งการของ", value=default_commander if i==0 else "", key=f"cmd_{i}")
        
        st.write("รายชื่อเจ้าหน้าที่ผู้จับกุม:")
        default_df = pd.DataFrame([{"ยศ": "พ.ต.ท.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])
        edited_officers = st.data_editor(default_df, num_rows="dynamic", key=f"editor_{i}", use_container_width=True)
        
        valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].str.strip() != ""]
        officers_list = [{"full_name": f"{r['ยศ']}{r['ชื่อ-นามสกุล']}", "display": f"{r['ยศ']}{r['ชื่อ-นามสกุล']} {r['ตำแหน่ง']}"} for _, r in valid_officers.iterrows()]
        
        # จับคู่ลายเซ็นทีละ 2 คน
        signature_rows = []
        for j in range(0, len(officers_list), 2):
            officer1 = officers_list[j]['full_name']
            officer2 = officers_list[j+1]['full_name'] if j+1 < len(officers_list) else ""
            signature_rows.append({"officer1_name": officer1, "officer2_name": officer2})

        units_data.append({
            "unit_name": unit_name,
            "commanders_text": commanders_text,
            "arresting_officers_text": ", ".join([o["display"] for o in officers_list]),
            "signature_rows": signature_rows
        })

st.button("➕ เพิ่มหน่วยจับกุม (กรณีจับร่วม)", on_click=add_unit)
st.divider()

# --- 3. ข้อมูลผู้ถูกจับกุม และข้อหา ---
st.header("3. ข้อมูลผู้ถูกจับกุม และข้อหา (ข้อ 8, 9, 10)")
suspect_mode = st.radio("รูปแบบผู้ถูกจับกุม", ["1 คน (กรอกข้อมูล)", "หลายคน (อัปโหลด Excel)"], horizontal=True)

final_suspects = []
charges_text = ""

if suspect_mode == "1 คน (กรอกข้อมูล)":
    col1, col2, col3 = st.columns(3)
    with col1: s_name = st.text_input("ชื่อ-นามสกุล")
    with col2: s_id = st.text_input("เลขประจำตัวประชาชน")
    with col3: s_address = st.text_area("ที่อยู่", height=68)
    single_charge = st.text_area("ฐานความผิด")
    
    if s_name:
        final_suspects.append({"name": s_name, "id_card": s_id, "address": s_address})
    charges_text = f"ซึ่งต้องหาว่ากระทำความผิดฐาน {single_charge}" if single_charge else ""

else:
    st.info("ไฟล์ Excel ต้องมีหัวตาราง: ชื่อ-นามสกุล | เลขบัตรประจำตัวประชาชน | ที่อยู่ | ฐานความผิด")
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df, use_container_width=True)
        charge_list = []
        for idx, row in df.iterrows():
            final_suspects.append({
                "name": str(row.get("ชื่อ-นามสกุล", "-")),
                "id_card": str(row.get("เลขบัตรประจำตัวประชาชน", "-")),
                "address": str(row.get("ที่อยู่", "-"))
            })
            charge_list.append(f"ผู้ต้องหาที่ {idx+1} ซึ่งต้องหาว่ากระทำความผิดฐาน {str(row.get('ฐานความผิด', '-'))}")
        charges_text = " ".join(charge_list)

arrest_circumstances = st.text_area("พฤติการณ์และรายละเอียดเกี่ยวกับเหตุแห่งการจับ", height=150)
st.divider()

# --- 4. ถ้อยคำ และ พ.ร.บ.อุ้มหายฯ ---
st.header("4. ถ้อยคำ และ พ.ร.บ.อุ้มหายฯ (ข้อ 12, 13)")
col1, col2 = st.columns(2)
with col1:
    confession = st.radio("การให้ถ้อยคำในชั้นจับกุม", ["รับสารภาพตลอดข้อกล่าวหา", "ปฏิเสธตลอดข้อกล่าวหา"])
    additional_statement = st.text_area("ให้การเพิ่มเติมว่า...", value="รับว่าเป็นบุคคลตามหมายจับจริง และไม่เคยถูกจับตามหมายจับดังกล่าวมาก่อน")
with col2:
    st.write(f"วันแจ้งอัยการ/ปกครอง: {to_thai_datetime(arrest_dt, is_time_included=False)}")
    torture_time = st.time_input("เวลาที่แจ้ง (พ.ร.บ.ทรมาน)", datetime.datetime.now().time())
st.divider()

# --- 5. สร้างไฟล์ Word ---
st.header("5. ประมวลผลและสร้างเอกสาร")

if st.button("💾 สร้างไฟล์บันทึกจับกุม (Word)", type="primary", use_container_width=True):
    try:
        context = {
            "record_location": record_location,
            "record_datetime_th": to_thai_datetime(record_dt),
            "arrest_datetime_th": to_thai_datetime(arrest_dt),
            "arrest_location": arrest_location,
            "charges_text": charges_text,
            "arrest_circumstances": arrest_circumstances,
            "suspects": final_suspects,
            "units": units_data,
            "confession": confession,
            "additional_statement": additional_statement
        }

        doc = DocxTemplate("template_arrest.docx")
        doc.render(context)

        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)

        st.success("สร้างไฟล์สำเร็จ! กดปุ่มด้านล่างเพื่อดาวน์โหลด")
        st.download_button(
            label="⬇️ ดาวน์โหลด บันทึกจับกุม.docx",
            data=bio.getvalue(),
            file_name=f"บันทึกจับกุม_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")