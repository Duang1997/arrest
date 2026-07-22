import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import datetime

# --- ฟังก์ชันแปลงวันที่ ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

def combine_date_time_text(date_obj, time_str):
    date_text = format_thai_date(date_obj)
    if time_str:
        return f"{date_text} เวลาประมาณ {time_str} น."
    return date_text

st.set_page_config(page_title="ระบบบันทึกจับกุม & ม.22", layout="centered") 
st.title("แบบฟอร์มบันทึกจับกุม และแจ้ง ม.22")
st.divider()

# ==========================================
# ส่วนที่ 1: รายละเอียดเกี่ยวกับวันที่และสถานที่จับกุม
# ==========================================
st.header("ส่วนที่ 1: รายละเอียดเกี่ยวกับวันที่และสถานที่")
record_location = st.text_input("สถานที่ทำการบันทึก", value="กองกำกับการ 3 กองบังคับการปราบปราม")

st.markdown("**ข้อมูลการบันทึก**")
record_date = st.date_input("วันที่บันทึก (กดที่ไอคอนปฏิทิน)")
record_time = st.text_input("เวลาที่บันทึก (เช่น 19:30)", value=datetime.datetime.now().strftime('%H:%M'))

st.markdown("**ข้อมูลการจับกุม**")
arrest_location = st.text_area("สถานที่จับกุม", height=100)
arrest_date = st.date_input("วันที่จับกุม (กดที่ไอคอนปฏิทิน)")
arrest_time = st.text_input("เวลาที่จับกุม (เช่น 18:30)", value=datetime.datetime.now().strftime('%H:%M'))

record_datetime_th = combine_date_time_text(record_date, record_time)
arrest_datetime_th = combine_date_time_text(arrest_date, arrest_time)
st.divider()

# ==========================================
# ส่วนที่ 2: รายละเอียดเกี่ยวกับหน่วยการจับกุม
# ==========================================
st.header("ส่วนที่ 2: รายละเอียดเกี่ยวกับหน่วยการจับกุม")

if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 1

def add_unit():
    st.session_state.unit_count += 1

units_data = []
all_officer_names = [] # เก็บรายชื่อตำรวจทั้งหมดสำหรับทำ Dropdown
default_cmd = "พล.ต.ท.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พ.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."

for i in range(st.session_state.unit_count):
    with st.container(border=True):
        st.subheader(f"🏢 หน่วยจับกุมที่ {i+1}")
        unit_name = st.text_input(f"ชื่อหน่วยงานที่ {i+1}", value="กก.๓ บก.ป." if i==0 else "", key=f"unit_name_{i}")
        commanders_text = st.text_area(f"ภายใต้อำนวยการสั่งการของ", value=default_cmd if i==0 else "", key=f"cmd_{i}")
        
        st.markdown("**รายชื่อเจ้าหน้าที่ผู้จับกุม:**")
        default_officers = pd.DataFrame([{"ยศ": "พ.ต.ท.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])
        edited_officers = st.data_editor(default_officers, num_rows="dynamic", key=f"editor_{i}", use_container_width=True)
        
        valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].str.strip() != ""]
        officers_list = []
        for _, r in valid_officers.iterrows():
            # รวมยศและชื่อ
            officer_fullname = f"{r.get('ยศ', '')}{r.get('ชื่อ-นามสกุล', '')}"
            all_officer_names.append(officer_fullname)
            
            officers_list.append({
                "rank": str(r.get('ยศ', '')),
                "name_only": str(r.get('ชื่อ-นามสกุล', '')),
                "display": f"{officer_fullname} {r.get('ตำแหน่ง', '')}"
            })
        
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

# ==========================================
# ส่วนที่ 3: ข้อมูลผู้ถูกจับกุม
# ==========================================
st.header("ส่วนที่ 3: ข้อมูลผู้ถูกจับกุม")

arrest_type = st.radio("ประเภทการจับกุม", ["จับสด", "จับตามหมายจับ"])
warrant_text = ""
if arrest_type == "จับตามหมายจับ":
    warrant_df = pd.DataFrame([{"ศาลที่ออกหมาย (เช่น อาญา)": "", "เลขที่หมาย (เช่น 787/2569)": "", "ลงวันที่ (เช่น 9 ก.พ. 69)": ""}])
    edited_warrants = st.data_editor(warrant_df, num_rows="dynamic", use_container_width=True)
    w_list = []
    for _, w in edited_warrants.iterrows():
        if w["ศาลที่ออกหมาย (เช่น อาญา)"]:
            w_list.append(f"ผู้ต้องหาตามหมายจับศาล{w['ศาลที่ออกหมาย (เช่น อาญา)']} ที่ {w['เลขที่หมาย (เช่น 787/2569)']} ลงวันที่ {w['ลงวันที่ (เช่น 9 ก.พ. 69)']}")
    if w_list:
        warrant_text = " ".join(w_list) + "\n"

suspect_mode = st.radio("รูปแบบการเพิ่มผู้ถูกจับกุม", ["กรอกผ่านตารางในเว็บ", "อัปโหลดไฟล์ Excel"])
final_suspects = []

if suspect_mode == "กรอกผ่านตารางในเว็บ":
    suspect_df = pd.DataFrame([{"ชื่อ-นามสกุล": "", "เลขประจำตัวประชาชน": "", "ที่อยู่": "", "ฐานความผิด": ""}])
    edited_suspects = st.data_editor(suspect_df, num_rows="dynamic", use_container_width=True)
    
    for idx, row in edited_suspects.iterrows():
        name = str(row.get("ชื่อ-นามสกุล", "")).strip()
        if name != "" and name.lower() != "nan":
            final_suspects.append({
                "index": len(final_suspects) + 1,
                "name": name,
                "id_card": str(row.get("เลขประจำตัวประชาชน", "-")).replace(".0", ""),
                "address": str(row.get("ที่อยู่", "-")),
                "charge": str(row.get("ฐานความผิด", "-"))
            })
else:
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        st.dataframe(df, use_container_width=True)
        
        for idx, row in df.iterrows():
            name = str(row.get("ชื่อ-นามสกุล", "")).strip()
            if name != "" and name.lower() != "nan":
                final_suspects.append({
                    "index": len(final_suspects) + 1,
                    "name": name,
                    "id_card": str(row.get("เลขบัตรประจำตัวประชาชน", "-")).replace(".0", ""),
                    "address": str(row.get("ที่อยู่", "-")),
                    "charge": str(row.get("ฐานความผิด", "-"))
                })

arrest_circumstances = st.text_area("พฤติการณ์และรายละเอียดเกี่ยวกับเหตุแห่งการจับ", height=150)
st.divider()

# ==========================================
# ส่วนที่ 4: สิทธิ และแบบแจ้ง ม.22 (เพิ่มรูปภาพ)
# ==========================================
st.header("ส่วนที่ 4: ข้อมูลเพิ่มเติม ม.22 และรูปถ่าย")

# ข้อมูลสำหรับแบบแจ้ง ม.22
st.markdown("**1. สถานที่ควบคุม และเจ้าหน้าที่ (ม.22)**")
detention_location = st.text_area("สถานที่ควบคุมตัวไว้", value="กองกำกับการ 3 กองบังคับการปราบปราม")

col1, col2 = st.columns(2)
# ดึง Dropdown ชื่อตำรวจจาก ส่วนที่ 2
officer_choices = all_officer_names if all_officer_names else ["กรุณาเพิ่มตำรวจในส่วนที่ 2"]
with col1: 
    officer_m22_name = st.selectbox("เจ้าหน้าที่รัฐผู้รับผิดชอบ (ม.22)", officer_choices, key="m22_res")
    officer_m22_phone = st.text_input("เบอร์โทรศัพท์ (ผู้รับผิดชอบ)")
with col2: 
    notif_officer_name = st.selectbox("เจ้าหน้าที่ผู้แจ้ง (ม.22)", officer_choices, key="m22_notif")
    notif_phone = st.text_input("เบอร์โทรศัพท์ (ผู้แจ้ง)")

fm_choice = st.radio("เหตุสุดวิสัย (กรณีไม่สามารถบันทึกภาพ/เสียง)", ["ไม่มี", "ระหว่างฝากควบคุมตัวชั่วคราวชุดจับไม่สามารถบันทึกวิดีโอได้", "อื่นๆ ระบุ"])
force_majeure = st.text_input("ระบุเหตุสุดวิสัย") if fm_choice == "อื่นๆ ระบุ" else fm_choice

st.markdown("---")
st.markdown("**2. อัปโหลดรูปถ่าย และการแจ้งญาติ (แยกรายบุคคล)**")

if len(final_suspects) == 0:
    st.warning("⚠️ กรุณากรอกข้อมูลผู้ต้องหาใน ส่วนที่ 3 ก่อน")
else:
    is_multi = len(final_suspects) > 1
    for s in final_suspects:
        with st.container(border=True):
            st.subheader(f"🗣 ผู้ต้องหาที่ {s['index']}: {s['name']}" if is_multi else f"🗣 ข้อมูลของ: {s['name']}")
            
            # อัปโหลดรูปภาพ 4 มุม สำหรับ ม.22
            st.caption("📷 อัปโหลดรูปถ่ายสำหรับ ม.22 (หน้าตรง, ซ้าย, ขวา, หลัง)")
            c1, c2, c3, c4 = st.columns(4)
            s['img_front'] = c1.file_uploader("หน้าตรง", type=['png', 'jpg', 'jpeg'], key=f"img_f_{s['index']}")
            s['img_left']  = c2.file_uploader("หันซ้าย", type=['png', 'jpg', 'jpeg'], key=f"img_l_{s['index']}")
            s['img_right'] = c3.file_uploader("หันขวา", type=['png', 'jpg', 'jpeg'], key=f"img_r_{s['index']}")
            s['img_back']  = c4.file_uploader("หันหลัง", type=['png', 'jpg', 'jpeg'], key=f"img_b_{s['index']}")

            # ข้อมูลญาติและคำให้การ (สำหรับบันทึกจับกุม)
            st.caption("📝 ข้อมูลญาติ และคำให้การ (สำหรับบันทึกจับกุม)")
            rel_name = st.text_input("ชื่อญาติที่ประสงค์แจ้ง", placeholder="ถ้าไม่มีให้ปล่อยว่างไว้", key=f"rel_{s['index']}")
            relative_name_final = rel_name if rel_name.strip() else "............................................................................ ผู้ซึ่งตนไว้วางใจทราบถึงการจับกุมด้วยแล้ว"
            
            confession = st.radio("การให้ถ้อยคำในชั้นจับกุม", ["รับสารภาพตลอดข้อกล่าวหา", "ปฏิเสธตลอดข้อกล่าวหา"], key=f"conf_{s['index']}")
            additional_statement = st.text_area("ให้การเพิ่มเติมว่า...", value="รับว่าเป็นบุคคลตามหมายจับจริง และไม่เคยถูกจับตามหมายจับดังกล่าวมาก่อน", key=f"add_{s['index']}")
            
            s['display_index'] = f"{s['index']}. " if is_multi else ""
            s['charge_text'] = f"ผู้ต้องหาที่ {s['index']} ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}" if is_multi else f"ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}"
            s['relative_text'] = f"ผู้ต้องหาที่ {s['index']} {s['name']} แจ้งให้: {relative_name_final}" if is_multi else f" : {relative_name_final}"
            s['statement_prefix'] = f"ผู้ต้องหาที่ {s['index']} {s['name']} :" if is_multi else ""
            s['confession'] = confession
            s['additional_statement'] = additional_statement

st.divider()

# ==========================================
# ส่วนที่ 5: สร้างเอกสาร
# ==========================================
st.header("ส่วนที่ 5: สร้างเอกสาร")

col1, col2 = st.columns(2)

# ปุ่มดาวน์โหลด 1: บันทึกจับกุม
with col1:
    if st.button("📄 สร้างไฟล์ บันทึกจับกุม", type="primary", use_container_width=True):
        try:
            charge_list = [f"ผู้ต้องหาที่ {s['index']} ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}" for s in final_suspects]
            charges_text = warrant_text + " ".join(charge_list)

            context_arrest = {
                "record_location": record_location,
                "record_datetime_th": record_datetime_th,
                "arrest_datetime_th": arrest_datetime_th,
                "arrest_location": arrest_location,
                "arrest_circumstances": arrest_circumstances,
                "suspects": final_suspects,
                "units": units_data,
                "arrest_date_text": format_thai_date(arrest_date),
                "warrant_text": warrant_text,
                "charges_text": charges_text
            }

            doc_arrest = DocxTemplate("template_arrest.docx")
            doc_arrest.render(context_arrest)
            bio_arrest = BytesIO()
            doc_arrest.save(bio_arrest)
            bio_arrest.seek(0)
            
            st.success("✅ สร้างบันทึกจับกุมสำเร็จ!")
            st.download_button(
                label="⬇️ ดาวน์โหลด บันทึกจับกุม.docx",
                data=bio_arrest.getvalue(),
                file_name=f"บันทึกจับกุม_{datetime.datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด (บันทึกจับกุม): {e}")

# ปุ่มดาวน์โหลด 2: แบบแจ้ง ม.22
with col2:
    if st.button("📄 สร้างไฟล์ แบบแจ้ง ม.22", type="primary", use_container_width=True):
        try:
            doc_m22 = DocxTemplate("template_section22.docx")
            
            # ประมวลผลรูปภาพสำหรับ ม.22 ด้วย InlineImage
            # กำหนดความกว้างของรูปให้พอดีกับตาราง (ประมาณ 35mm จะเรียง 4 รูปได้พอดีในหน้า A4)
            for s in final_suspects:
                s['pic_front'] = InlineImage(doc_m22, s['img_front'], width=Mm(35)) if s['img_front'] else ""
                s['pic_left']  = InlineImage(doc_m22, s['img_left'], width=Mm(35)) if s['img_left'] else ""
                s['pic_right'] = InlineImage(doc_m22, s['img_right'], width=Mm(35)) if s['img_right'] else ""
                s['pic_back']  = InlineImage(doc_m22, s['img_back'], width=Mm(35)) if s['img_back'] else ""

            context_m22 = {
                "arrest_date_text": format_thai_date(arrest_date),
                "arrest_time": arrest_time,
                "arrest_location": arrest_location,
                "arrest_circumstances": arrest_circumstances,
                "suspects": final_suspects,
                "detention_location": detention_location,
                "officer_m22_name": officer_m22_name,
                "officer_m22_phone": officer_m22_phone,
                "notif_officer_name": notif_officer_name,
                "notif_phone": notif_phone,
                "force_majeure": force_majeure
            }

            doc_m22.render(context_m22)
            bio_m22 = BytesIO()
            doc_m22.save(bio_m22)
            bio_m22.seek(0)
            
            st.success("✅ สร้างแบบแจ้ง ม.22 สำเร็จ!")
            st.download_button(
                label="⬇️ ดาวน์โหลด แบบแจ้ง_ม.22.docx",
                data=bio_m22.getvalue(),
                file_name=f"แบบแจ้ง_ม22_{datetime.datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด (ม.22): โปรดตรวจสอบไฟล์ template_section22.docx ({e})")
