import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
import datetime

# --- ฟังก์ชันช่วยเหลือ (Helper Functions) ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

def combine_date_time_text(date_obj, time_str):
    date_text = format_thai_date(date_obj)
    if time_str:
        return f"{date_text} เวลาประมาณ {time_str} น."
    return date_text

# --- ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="ระบบบันทึกจับกุม", layout="wide")
st.title("แบบฟอร์มบันทึกจับกุมอัจฉริยะ")

# --- 1. ข้อมูลสถานที่และเวลา ---
st.header("1. ข้อมูลสถานที่และเวลา")
col1, col2 = st.columns(2)
with col1:
    record_location = st.text_input("สถานที่ทำการบันทึก", value="กองกำกับการ 3 กองบังคับการปราบปราม")
    arrest_location = st.text_area("สถานที่จับกุม", height=100)
with col2:
    # ใช้วิธีแสดงปี พ.ศ. กำกับไว้ข้างๆ ช่องเลือกวันที่
    record_date = st.date_input("วันที่บันทึก (กดที่ไอคอนปฏิทิน)", datetime.date.today())
    st.caption(f"📌 ตรงกับ: {format_thai_date(record_date)}")
    record_time = st.text_input("เวลาที่บันทึก (กรอกมือ เช่น 19:30)", value=datetime.datetime.now().strftime('%H:%M'))
    
    arrest_date = st.date_input("วันที่จับกุม (กดที่ไอคอนปฏิทิน)", datetime.date.today())
    st.caption(f"📌 ตรงกับ: {format_thai_date(arrest_date)}")
    arrest_time = st.text_input("เวลาที่จับกุม (กรอกมือ เช่น 18:20)", value=datetime.datetime.now().strftime('%H:%M'))

record_datetime_th = combine_date_time_text(record_date, record_time)
arrest_datetime_th = combine_date_time_text(arrest_date, arrest_time)
st.divider()

# --- 2. หน่วยจับกุมและเจ้าหน้าที่ ---
st.header("2. หน่วยจับกุมและเจ้าหน้าที่")
if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 1

def add_unit():
    st.session_state.unit_count += 1

units_data = []
default_cmd = "พล.ต.ท.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พล.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."

for i in range(st.session_state.unit_count):
    with st.expander(f"🏢 หน่วยจับกุมที่ {i+1}", expanded=True):
        unit_name = st.text_input(f"ชื่อหน่วยงานที่ {i+1}", value="กก.๓ บก.ป." if i==0 else "", key=f"unit_name_{i}")
        commanders_text = st.text_area(f"ภายใต้อำนวยการสั่งการของ", value=default_cmd if i==0 else "", key=f"cmd_{i}")
        
        st.write("รายชื่อเจ้าหน้าที่ผู้จับกุม (เพิ่มแถวได้โดยกดที่ตาราง):")
        default_officers = pd.DataFrame([{"ยศ": "พ.ต.ท.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])
        edited_officers = st.data_editor(default_officers, num_rows="dynamic", key=f"editor_{i}", use_container_width=True)
        
        valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].str.strip() != ""]
        officers_list = []
        for _, r in valid_officers.iterrows():
            officers_list.append({
                "rank": r['ยศ'],
                "name_only": r['ชื่อ-นามสกุล'],
                "display": f"{r['ยศ']}{r['ชื่อ-นามสกุล']} {r['ตำแหน่ง']}"
            })
        
        # จับคู่ลายเซ็น แยกยศและชื่อ
        signature_rows = []
        for j in range(0, len(officers_list), 2):
            o1 = officers_list[j]
            o2 = officers_list[j+1] if j+1 < len(officers_list) else {"rank": "", "name_only": ""}
            signature_rows.append({
                "officer1_rank": o1["rank"], "officer1_name_only": o1["name_only"],
                "officer2_rank": o2["rank"], "officer2_name_only": o2["name_only"]
            })

        units_data.append({
            "unit_name": unit_name,
            "commanders_text": commanders_text,
            "arresting_officers_text": ", ".join([o["display"] for o in officers_list]),
            "signature_rows": signature_rows
        })
st.button("➕ เพิ่มหน่วยจับกุมอื่น (กรณีจับร่วม)", on_click=add_unit)
st.divider()

# --- 3. ข้อมูลผู้ถูกจับกุม และประเภทการจับ ---
st.header("3. ข้อมูลผู้ถูกจับกุม และข้อหา")

# ประเภทการจับกุม (จับสด / จับตามหมาย)
arrest_type = st.radio("ประเภทการจับกุม", ["จับสด", "จับตามหมายจับ"], horizontal=True)
warrant_text = ""
if arrest_type == "จับตามหมายจับ":
    st.write("รายละเอียดหมายจับ (กดเพิ่มแถวได้ถ้ามีหลายหมาย):")
    warrant_df = pd.DataFrame([{"ศาลที่ออกหมาย (เช่น อาญา)": "", "เลขที่หมาย (เช่น 787/2569)": "", "ลงวันที่ (เช่น 9 กุมภาพันธ์ 2569)": ""}])
    edited_warrants = st.data_editor(warrant_df, num_rows="dynamic", use_container_width=True)
    
    w_list = []
    for _, w in edited_warrants.iterrows():
        if w["ศาลที่ออกหมาย (เช่น อาญา)"]:
            w_list.append(f"ผู้ต้องหาตามหมายจับศาล{w['ศาลที่ออกหมาย (เช่น อาญา)']} ที่ {w['เลขที่หมาย (เช่น 787/2569)']} ลงวันที่ {w['ลงวันที่ (เช่น 9 กุมภาพันธ์ 2569)']}")
    if w_list:
        warrant_text = " ".join(w_list) + " "

# เลือกรูปแบบกรอกข้อมูล
suspect_mode = st.radio("รูปแบบกรอกผู้ถูกจับกุม", ["กรอกผ่านตารางในเว็บ (แนะนำ)", "อัปโหลดไฟล์ Excel"], horizontal=True)
final_suspects = []

if suspect_mode == "กรอกผ่านตารางในเว็บ (แนะนำ)":
    st.write("กรอกข้อมูลผู้ต้องหา (กดปุ่มบวกมุมซ้ายล่างของตารางเพื่อเพิ่มคน):")
    suspect_df = pd.DataFrame([{
        "ชื่อ-นามสกุล": "", "เลขประจำตัวประชาชน": "", 
        "ที่อยู่": "", "ฐานความผิด": "", "ชื่อญาติที่ประสงค์แจ้ง": ""
    }])
    edited_suspects = st.data_editor(suspect_df, num_rows="dynamic", use_container_width=True)
    
    for idx, row in edited_suspects.iterrows():
        if row["ชื่อ-นามสกุล"]:
            final_suspects.append({
                "index": idx + 1,
                "name": str(row["ชื่อ-นามสกุล"]),
                "id_card": str(row["เลขประจำตัวประชาชน"]),
                "address": str(row["ที่อยู่"]),
                "charge": str(row["ฐานความผิด"]),
                "relative_name": str(row["ชื่อญาติที่ประสงค์แจ้ง"]) if str(row["ชื่อญาติที่ประสงค์แจ้ง"]) else "ไม่ประสงค์แจ้ง"
            })
else:
    st.info("💡 ไฟล์ Excel ต้องมีหัวตาราง: ชื่อ-นามสกุล | เลขบัตรประจำตัวประชาชน | ที่อยู่ | ฐานความผิด | ชื่อญาติที่ประสงค์แจ้ง")
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df, use_container_width=True)
        for idx, row in df.iterrows():
            final_suspects.append({
                "index": idx + 1,
                "name": str(row.get("ชื่อ-นามสกุล", "-")),
                "id_card": str(row.get("เลขบัตรประจำตัวประชาชน", "-")),
                "address": str(row.get("ที่อยู่", "-")),
                "charge": str(row.get("ฐานความผิด", "-")),
                "relative_name": str(row.get("ชื่อญาติที่ประสงค์แจ้ง", "ไม่ประสงค์แจ้ง"))
            })

# รวมข้อหา
charge_list = []
for s in final_suspects:
    charge_list.append(f"ผู้ต้องหาที่ {s['index']} ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}")
charges_text = warrant_text + " ".join(charge_list)

arrest_circumstances = st.text_area("พฤติการณ์และรายละเอียดเกี่ยวกับเหตุแห่งการจับ", height=150)
st.divider()

# --- 4. ถ้อยคำ และ พ.ร.บ.อุ้มหายฯ ---
st.header("4. ถ้อยคำ และ พ.ร.บ.อุ้มหายฯ")
col1, col2 = st.columns(2)
with col1:
    confession = st.radio("การให้ถ้อยคำในชั้นจับกุม", ["รับสารภาพตลอดข้อกล่าวหา", "ปฏิเสธตลอดข้อกล่าวหา"])
    additional_statement = st.text_area("ให้การเพิ่มเติมว่า...", value="รับว่าเป็นบุคคลตามหมายจับจริง และไม่เคยถูกจับตามหมายจับดังกล่าวมาก่อน")
with col2:
    st.write(f"**วันแจ้งอัยการ/ปกครอง:** {format_thai_date(arrest_date)} (ดึงจากวันจับกุม)")
    torture_time = st.text_input("เวลาที่แจ้ง พ.ร.บ.ทรมาน (กรอกมือ เช่น 20:00)", value=datetime.datetime.now().strftime('%H:%M'))
st.divider()

# --- 5. สร้างไฟล์ Word ---
st.header("5. ประมวลผลและสร้างเอกสาร")
if st.button("💾 สร้างไฟล์บันทึกจับกุม (Word)", type="primary", use_container_width=True):
    try:
        context = {
            "record_location": record_location,
            "record_datetime_th": record_datetime_th,
            "arrest_datetime_th": arrest_datetime_th,
            "arrest_location": arrest_location,
            "charges_text": charges_text,
            "arrest_circumstances": arrest_circumstances,
            "suspects": final_suspects,
            "units": units_data,
            "confession": confession,
            "additional_statement": additional_statement,
            "arrest_date_text": format_thai_date(arrest_date), # ส่งวันที่จับกุมไปให้ข้อ 13
            "torture_time": torture_time
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
        st.error(f"เกิดข้อผิดพลาด: โปรดตรวจสอบไฟล์ Template ว่ามีการใส่ตัวแปรถูกต้องหรือไม่ (Error: {e})")
