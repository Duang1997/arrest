import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import datetime

# --- ตั้งค่าหน้าเพจ ---
st.set_page_config(page_title="ระบบบันทึกจับกุม & ม.22, ม.23 (CIB)", layout="wide") 

# --- ตกแต่ง UI ธีม CIB ---
st.markdown("""
<style>
    .cib-header { background-color: #00204a; padding: 15px; border-radius: 5px; color: #f9bc0f; text-align: center; font-family: sans-serif; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    .cib-header h1 { color: #f9bc0f; margin: 0; font-size: 28px; font-weight: bold; }
    .cib-header p { color: #ffffff; margin: 0; font-size: 16px; }
    .stTabs [data-baseweb="tab-list"] { gap: 20px; background-color: #00204a; padding: 10px 20px; border-radius: 5px; }
    .stTabs [data-baseweb="tab"] { color: #ffffff; font-weight: bold; font-size: 16px; }
    .stTabs [aria-selected="true"] { color: #f9bc0f !important; border-bottom-color: #f9bc0f !important; }
</style>
<div class="cib-header">
    <h1>ตำรวจสอบสวนกลาง (Central Investigation Bureau)</h1>
    <p>ระบบจัดทำบันทึกจับกุม แบบแจ้ง ม.22 และ แบบบันทึก ม.23</p>
</div>
""", unsafe_allow_html=True)

# --- ฟังก์ชันจัดการวันที่ ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

def combine_date_time_text(date_obj, time_str):
    date_text = format_thai_date(date_obj)
    if time_str:
        return f"{date_text} เวลาประมาณ {time_str} น."
    return date_text

# --- จัดการ Session State ---
if 'shared_data' not in st.session_state:
    st.session_state.shared_data = {}
if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 1

def add_unit():
    st.session_state.unit_count += 1

# --- โครงสร้างระบบ Tabs (ยุบรวม ม.22 และ 23) ---
tab_arrest, tab_m22_23 = st.tabs(["📝 ฟังก์ชันบันทึกจับกุม", "⚖️ ฟังก์ชันแบบแจ้ง ม.22 และบันทึก ม.23"])

# ==========================================
# โหมดที่ 1: ฟังก์ชันบันทึกจับกุม
# ==========================================
with tab_arrest:
    st.header("ส่วนที่ 1: รายละเอียดเกี่ยวกับวันที่และสถานที่")
    record_location = st.text_input("สถานที่ทำการบันทึก", value="กองกำกับการ 3 กองบังคับการปราบปราม")

    col1, col2 = st.columns(2)
    with col1:
        record_date = st.date_input("วันที่บันทึก", key="rec_d")
        record_time = st.text_input("เวลาที่บันทึก", value=datetime.datetime.now().strftime('%H:%M'), key="rec_t")
    with col2:
        arrest_date = st.date_input("วันที่จับกุม", key="arr_d")
        arrest_time = st.text_input("เวลาที่จับกุม", value=datetime.datetime.now().strftime('%H:%M'), key="arr_t")
    
    arrest_location = st.text_area("สถานที่จับกุม", height=100, key="arr_loc")
    record_datetime_th = combine_date_time_text(record_date, record_time)
    arrest_datetime_th = combine_date_time_text(arrest_date, arrest_time)
    st.divider()

    st.header("ส่วนที่ 2: รายละเอียดเกี่ยวกับหน่วยการจับกุม")
    units_data = []
    officers_data = [] # เก็บเป็นชุดข้อมูลสำหรับดึงไปใช้ใน ม.23
    officer_displays = [] 
    default_cmd = "พล.ต.ท.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พ.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."

    for i in range(st.session_state.unit_count):
        with st.container(border=True):
            st.subheader(f"🏢 หน่วยจับกุมที่ {i+1}")
            unit_name = st.text_input(f"ชื่อหน่วยงาน", value="กก.๓ บก.ป." if i==0 else "", key=f"unit_name_{i}")
            commanders_text = st.text_area(f"ภายใต้อำนวยการสั่งการของ", value=default_cmd if i==0 else "", key=f"cmd_{i}")
            
            default_officers = pd.DataFrame([{"ยศ": "พ.ต.ท.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])
            edited_officers = st.data_editor(default_officers, num_rows="dynamic", key=f"editor_{i}", use_container_width=True)
            
            valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].str.strip() != ""]
            officers_list = []
            for _, r in valid_officers.iterrows():
                fullname = f"{r.get('ยศ', '')}{r.get('ชื่อ-นามสกุล', '')}"
                position = r.get('ตำแหน่ง', '')
                display = f"{fullname} {position}".strip()
                
                officers_data.append({"fullname": fullname, "position": position, "display": display})
                officer_displays.append(display)
                
                officers_list.append({"rank": str(r.get('ยศ', '')), "name_only": str(r.get('ชื่อ-นามสกุล', '')), "display": display})
            
            signature_rows = []
            for j in range(0, len(officers_list), 2):
                o1 = officers_list[j]
                o2 = officers_list[j+1] if j+1 < len(officers_list) else {"rank": "", "name_only": ""}
                signature_rows.append({"officer1_rank": o1["rank"], "officer1_name_only": o1["name_only"], "officer2_rank": o2["rank"], "officer2_name_only": o2["name_only"]})

            units_data.append({"unit_name": unit_name, "commanders_text": commanders_text, "arresting_officers_text": ", ".join([o["display"] for o in officers_list]), "signature_rows": signature_rows})
    st.button("➕ เพิ่มหน่วยจับกุมอื่น", on_click=add_unit)
    st.divider()

    st.header("ส่วนที่ 3: ข้อมูลผู้ถูกจับกุม")
    arrest_type = st.radio("ประเภทการจับกุม", ["จับสด", "จับตามหมายจับ"], horizontal=True)
    warrant_text, warrant_details = "", ""
    if arrest_type == "จับตามหมายจับ":
        warrant_df = pd.DataFrame([{"ศาลที่ออกหมาย": "", "เลขที่หมาย": "", "ลงวันที่": ""}])
        edited_warrants = st.data_editor(warrant_df, num_rows="dynamic", use_container_width=True)
        w_list = []
        for _, w in edited_warrants.iterrows():
            if w["ศาลที่ออกหมาย"]:
                w_list.append(f"ผู้ต้องหาตามหมายจับศาล{w['ศาลที่ออกหมาย']} ที่ {w['เลขที่หมาย']} ลงวันที่ {w['ลงวันที่']}")
                warrant_details = f"ศาล{w['ศาลที่ออกหมาย']} ที่ {w['เลขที่หมาย']} ลงวันที่ {w['ลงวันที่']}"
        if w_list: warrant_text = " ".join(w_list) + "\n"

    suspect_mode = st.radio("รูปแบบการเพิ่มผู้ถูกจับกุม", ["กรอกผ่านตารางในเว็บ", "อัปโหลดไฟล์ Excel"], horizontal=True)
    final_suspects = []

    if suspect_mode == "กรอกผ่านตารางในเว็บ":
        suspect_df = pd.DataFrame([{"ชื่อ-นามสกุล": "", "อายุ": "", "เลขประจำตัวประชาชน": "", "ที่อยู่": "", "ฐานความผิด": ""}])
        edited_suspects = st.data_editor(suspect_df, num_rows="dynamic", use_container_width=True)
        for idx, row in edited_suspects.iterrows():
            name = str(row.get("ชื่อ-นามสกุล", "")).strip()
            if name and name.lower() != "nan":
                final_suspects.append({"index": len(final_suspects) + 1, "name": name, "age": str(row.get("อายุ", "-")).replace(".0", ""), "id_card": str(row.get("เลขประจำตัวประชาชน", "-")).replace(".0", ""), "address": str(row.get("ที่อยู่", "-")), "charge": str(row.get("ฐานความผิด", "-"))})
    else:
        uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
        if uploaded_file:
            df = pd.read_excel(uploaded_file)
            st.dataframe(df, use_container_width=True)
            for idx, row in df.iterrows():
                name = str(row.get("ชื่อ-นามสกุล", "")).strip()
                if name and name.lower() != "nan":
                    final_suspects.append({"index": len(final_suspects) + 1, "name": name, "age": str(row.get("อายุ", "-")).replace(".0", ""), "id_card": str(row.get("เลขประจำตัวประชาชน", "-")).replace(".0", ""), "address": str(row.get("ที่อยู่", "-")), "charge": str(row.get("ฐานความผิด", "-"))})

    arrest_circumstances = st.text_area("พฤติการณ์และรายละเอียดเกี่ยวกับเหตุแห่งการจับ", height=150)
    st.divider()

    st.header("ส่วนที่ 4: สิทธิ และการแจ้งญาติ")
    if not final_suspects:
        st.warning("⚠️ กรุณากรอกข้อมูลผู้ต้องหาก่อน")
    else:
        is_multi = len(final_suspects) > 1
        for s in final_suspects:
            with st.container(border=True):
                st.subheader(f"🗣 ข้อมูลของ: {s['name']}")
                rel_name = st.text_input("ชื่อญาติที่ประสงค์แจ้ง", key=f"ar_rel_{s['index']}")
                relative_name_final = rel_name if rel_name.strip() else "............................................................................ ผู้ซึ่งตนไว้วางใจทราบถึงการจับกุมด้วยแล้ว"
                confession = st.radio("การให้ถ้อยคำในชั้นจับกุม", ["รับสารภาพตลอดข้อกล่าวหา", "ปฏิเสธตลอดข้อกล่าวหา"], key=f"ar_conf_{s['index']}", horizontal=True)
                
                s['display_index'] = f"{s['index']}. " if is_multi else ""
                s['charge_text'] = f"ผู้ต้องหาที่ {s['index']} ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}" if is_multi else f"ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}"
                s['relative_text'] = f"ผู้ต้องหาที่ {s['index']} {s['name']} แจ้งให้: {relative_name_final}" if is_multi else f" : {relative_name_final}"
                s['statement_prefix'] = f"ผู้ต้องหาที่ {s['index']} {s['name']} :" if is_multi else ""
                s['confession'] = confession
                s['additional_statement'] = st.text_area("ให้การเพิ่มเติมว่า...", value="รับว่าเป็นบุคคลตามหมายจับจริง และไม่เคยถูกจับตามหมายจับดังกล่าวมาก่อน", key=f"ar_add_{s['index']}")

    st.session_state.shared_data = {
        "arrest_location": arrest_location,
        "arrest_date_text": format_thai_date(arrest_date),
        "arrest_time": arrest_time,
        "arrest_circumstances": arrest_circumstances,
        "suspects": final_suspects,
        "officers_data": officers_data,
        "officer_displays": officer_displays if officer_displays else ["ไม่พบรายชื่อจากบันทึกจับกุม"],
        "warrant_details": warrant_details
    }

    st.divider()
    if st.button("💾 สร้างและดาวน์โหลด บันทึกจับกุม", type="primary", use_container_width=True):
        charge_list = [f"ผู้ต้องหาที่ {s['index']} ซึ่งต้องหาว่ากระทำความผิดฐาน {s['charge']}" for s in final_suspects]
        charges_text = warrant_text + " ".join(charge_list)
        context_arrest = {
            "record_location": record_location, "record_datetime_th": record_datetime_th, "arrest_datetime_th": arrest_datetime_th, "arrest_location": arrest_location,
            "arrest_circumstances": arrest_circumstances, "suspects": final_suspects, "units": units_data, "arrest_date_text": format_thai_date(arrest_date), "warrant_text": warrant_text, "charges_text": charges_text
        }
        doc_arrest = DocxTemplate("template_arrest.docx")
        doc_arrest.render(context_arrest)
        bio_arrest = BytesIO()
        doc_arrest.save(bio_arrest)
        bio_arrest.seek(0)
        st.download_button(label="⬇️ โหลดไฟล์ บันทึกจับกุม.docx", data=bio_arrest.getvalue(), file_name=f"บันทึกจับกุม_{datetime.datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ==========================================
# โหมดที่ 2: ฟังก์ชันแบบแจ้ง ม.22 และ ม.23
# ==========================================
with tab_m22_23:
    shared = st.session_state.shared_data
    suspects_from_arrest = shared.get("suspects", [])
    officer_choices = shared.get("officer_displays", ["ไม่พบรายชื่อ"])
    dropdown_opts = officer_choices + ["อื่นๆ (กรอกเพิ่มเติม)"]

    if not suspects_from_arrest:
        st.warning("⚠️ ไม่พบข้อมูลผู้ต้องหา กรุณากรอกตารางผู้ต้องหาในแท็บ 'บันทึกจับกุม' ก่อน")

    st.header("⚖️ ส่วนที่ 1: ข้อมูลกลางสำหรับ ม.22 และ ม.23")
    
    st.markdown("**► ข้อมูลฝั่ง ม.22 (แบบแจ้ง)**")
    col22_1, col22_2 = st.columns(2)
    with col22_1:
        detention_location = st.text_input("สถานที่ควบคุมตัวไว้ (ม.22)", value="กองกำกับการ 3 กองบังคับการปราบปราม")
        m22_officer_sel = st.selectbox("เจ้าหน้าที่รัฐผู้รับผิดชอบ (ม.22)", officer_choices, key="m22_res_s")
        m22_officer_phone = st.text_input("เบอร์โทรศัพท์ ผู้รับผิดชอบ (ม.22)", value="065-558-5054")
    with col22_2:
        fm_choice = st.radio("เหตุสุดวิสัยที่ไม่สามารถบันทึกภาพ/เสียง (ใช้ร่วมกัน)", ["ไม่มี", "อื่นๆ ระบุ"], horizontal=True)
        force_majeure = st.text_input("ระบุเหตุสุดวิสัย") if fm_choice == "อื่นๆ ระบุ" else fm_choice
        m22_notif_sel = st.selectbox("เจ้าหน้าที่ผู้แจ้ง (ม.22)", officer_choices, key="m22_notif_s")
        m22_notif_phone = st.text_input("เบอร์โทรศัพท์ ผู้แจ้ง (ม.22)", value="065-558-5054")

    # ประมวลผลดึง ยศ/ชื่อ/ตำแหน่ง ของผู้รับผิดชอบ ม.22 เพื่อส่งไป ม.23 อัตโนมัติ
    m22_dict = next((item for item in shared.get("officers_data", []) if item["display"] == m22_officer_sel), {"fullname": m22_officer_sel, "position": ""})

    st.markdown("**► ข้อมูลฝั่ง ม.23 (บันทึกการควบคุมตัว)**")
    st.info(f"📌 วันที่/เวลา/สถานที่ถูกควบคุมตัว: ดึงจากบันทึกจับกุมอัตโนมัติ\n📌 เจ้าหน้าที่ผู้ทำการควบคุมตัว: **{m22_dict['fullname']}** ตำแหน่ง **{m22_dict['position']}** (ดึงจาก ม.22)")
    
    col23_1, col23_2 = st.columns(2)
    with col23_1:
        dest_location = st.text_input("สถานที่ปลายทางที่รับตัว (ศาล/เรือนจำ)", value="ศาลจังหวัดสกลนคร")
        
        st.markdown("**เจ้าหน้าที่ผู้รับผิดชอบการย้ายตัว:**")
        transfer_sel = st.selectbox("เลือกเจ้าหน้าที่ผู้ย้ายตัว", dropdown_opts, key="m23_tr_s")
        if transfer_sel == "อื่นๆ (กรอกเพิ่มเติม)":
            transfer_name = st.text_input("ชื่อ-สกุล (ผู้ย้ายตัว)")
            transfer_pos = st.text_input("ตำแหน่ง (ผู้ย้ายตัว)")
        else:
            tr_dict = next((item for item in shared.get("officers_data", []) if item["display"] == transfer_sel), {})
            transfer_name = tr_dict.get("fullname", transfer_sel)
            transfer_pos = tr_dict.get("position", "")
        transfer_phone = st.text_input("เบอร์โทรศัพท์ (ผู้ย้ายตัว)")

    with col23_2:
        st.markdown("**วัน/เวลา/สถานที่ของการปล่อยตัว หรือ มอบตัว:**")
        release_date = st.date_input("วันที่ปล่อย/มอบตัว", key="rel_d")
        release_time = st.text_input("เวลาปล่อย/มอบตัว", value=datetime.datetime.now().strftime('%H:%M'))
        release_location = st.text_input("สถานที่ปล่อย/มอบตัว", value="ศาลจังหวัดสกลนคร")
        
        st.markdown("**เจ้าหน้าที่ผู้ออกคำสั่งให้ควบคุมตัว:**")
        cmd_sel = st.selectbox("เลือกผู้ออกคำสั่ง", dropdown_opts, key="m23_cmd_s")
        if cmd_sel == "อื่นๆ (กรอกเพิ่มเติม)":
            cmd_name = st.text_input("ชื่อ-สกุล (ผู้ออกคำสั่ง)")
            cmd_pos = st.text_input("ตำแหน่ง (ผู้ออกคำสั่ง)")
        else:
            cmd_dict = next((item for item in shared.get("officers_data", []) if item["display"] == cmd_sel), {})
            cmd_name = cmd_dict.get("fullname", cmd_sel)
            cmd_pos = cmd_dict.get("position", "")
        cmd_phone = st.text_input("เบอร์โทรศัพท์ (ผู้ออกคำสั่ง)")

    default_note = "ในการควบคุม เจ้าหน้าที่ตำรวจชุดจับกุม มิได้ทำร้าย ขู่เข็ญ อันเป็นการทรมาน การกระทำที่โหดร้าย ไร้มนุษยธรรม หรือย่ำยี่ศักดิ์ศรีความเป็นมนุษย์ หรือการกระทำให้บุคคลสูญหายแต่อย่างใด และไฟล์ภาพเคลื่อนไหวได้จัดเก็บไว้ที่หน่วยงานแล้ว"
    note_choice = st.radio("บันทึกอื่น ๆ เพิ่มเติม (ม.23)", ["ใช้ข้อความมาตรฐาน (Defalet)", "อื่นๆ (ระบุข้อความเอง)"])
    m23_additional_notes = default_note if note_choice == "ใช้ข้อความมาตรฐาน (Defalet)" else st.text_area("ระบุบันทึกเพิ่มเติม")

    st.divider()
    st.header("👤 ส่วนที่ 2: ข้อมูลผู้ต้องหา และแนบรูปถ่าย (ทำแยกรายบุคคล)")

    for s in suspects_from_arrest:
        with st.container(border=True):
            st.subheader(f"ผู้ต้องหา: {s['name']}")
            
            c_marks, c_pass = st.columns(2)
            s['marks'] = c_marks.text_input(f"ตำหนิรูปพรรณ (เช่น แผลเป็นที่แขน, รอยสักที่คอ)", value="ไม่มี", key=f"marks_{s['index']}")
            s['passport'] = c_pass.text_input(f"หนังสือเดินทาง", value="-", key=f"pass_{s['index']}")
            
            st.caption("📷 อัปโหลดรูปถ่าย 4 มุม (รูปชุดนี้จะใช้แนบท้ายทั้งเอกสาร ม.22 และ ม.23 อัตโนมัติ)")
            c1, c2, c3, c4 = st.columns(4)
            s['img_f'] = c1.file_uploader("หน้าตรง", type=['png', 'jpg'], key=f"pic_f_{s['index']}")
            s['img_l'] = c2.file_uploader("หันซ้าย", type=['png', 'jpg'], key=f"pic_l_{s['index']}")
            s['img_r'] = c3.file_uploader("หันขวา", type=['png', 'jpg'], key=f"pic_r_{s['index']}")
            s['img_b'] = c4.file_uploader("หันหลัง", type=['png', 'jpg'], key=f"pic_b_{s['index']}")

            if st.button(f"📄 สร้างและดาวน์โหลดเอกสาร ม.22 และ ม.23 ของ {s['name']}", key=f"btn_m2223_{s['index']}", type="primary"):
                try:
                    # -- จัดเตรียม Docx ม.22 --
                    doc_m22 = DocxTemplate("template_section22.docx")
                    ctx_m22 = {
                        "suspect": s, "arrest_date_text": shared.get("arrest_date_text", ""), "arrest_time": shared.get("arrest_time", ""),
                        "arrest_location": shared.get("arrest_location", ""), "arrest_circumstances": shared.get("arrest_circumstances", ""),
                        "detention_location": detention_location, "officer_m22_name": m22_officer_sel, "officer_m22_phone": m22_officer_phone,
                        "notif_officer_name": m22_notif_sel, "notif_phone": m22_notif_phone, "force_majeure": force_majeure
                    }
                    doc_m22.render(ctx_m22)
                    bio_m22 = BytesIO(); doc_m22.save(bio_m22); bio_m22.seek(0)
                    
                    # -- จัดเตรียมรูปภาพแนบท้าย ม.22 --
                    doc_att = DocxTemplate("template_attachment.docx")
                    ctx_att = {
                        "suspect": s, "officer_m22_name": m22_officer_sel,
                        "pic_front": InlineImage(doc_att, s['img_f'], width=Mm(75)) if s['img_f'] else "",
                        "pic_left": InlineImage(doc_att, s['img_l'], width=Mm(75)) if s['img_l'] else "",
                        "pic_right": InlineImage(doc_att, s['img_r'], width=Mm(75)) if s['img_r'] else "",
                        "pic_back": InlineImage(doc_att, s['img_b'], width=Mm(75)) if s['img_b'] else ""
                    }
                    doc_att.render(ctx_att)
                    bio_att = BytesIO(); doc_att.save(bio_att); bio_att.seek(0)

                    # -- จัดเตรียม Docx ม.23 --
                    doc_m23 = DocxTemplate("template_section23.docx")
                    ctx_m23 = {
                        "suspect": s, "arrest_date_text": shared.get("arrest_date_text", ""), "arrest_time": shared.get("arrest_time", ""),
                        "arrest_location": shared.get("arrest_location", ""), "warrant_details": shared.get("warrant_details", ""),
                        "ctrl_name": m22_dict['fullname'], "ctrl_pos": m22_dict['position'], "ctrl_phone": m22_officer_phone,
                        "cmd_name": cmd_name, "cmd_pos": cmd_pos, "cmd_phone": cmd_phone,
                        "dest_location": dest_location,
                        "trans_name": transfer_name, "trans_pos": transfer_pos, "trans_phone": transfer_phone,
                        "release_date_text": format_thai_date(release_date), "release_time": release_time, "release_location": release_location,
                        "force_majeure": force_majeure, "additional_notes": m23_additional_notes,
                        "pic_front": InlineImage(doc_m23, s['img_f'], width=Mm(75)) if s['img_f'] else "",
                        "pic_left": InlineImage(doc_m23, s['img_l'], width=Mm(75)) if s['img_l'] else "",
                        "pic_right": InlineImage(doc_m23, s['img_r'], width=Mm(75)) if s['img_r'] else "",
                        "pic_back": InlineImage(doc_m23, s['img_b'], width=Mm(75)) if s['img_b'] else ""
                    }
                    doc_m23.render(ctx_m23)
                    bio_m23 = BytesIO(); doc_m23.save(bio_m23); bio_m23.seek(0)
                    
                    st.success(f"✅ ประมวลผลเอกสารของ {s['name']} เสร็จสิ้น")
                    dl1, dl2, dl3 = st.columns(3)
                    with dl1: st.download_button(label="⬇️ 1. โหลด ม.22", data=bio_m22.getvalue(), file_name=f"ม22_{s['name']}.docx", use_container_width=True)
                    with dl2: st.download_button(label="⬇️ 2. โหลดแนบท้าย ม.22", data=bio_att.getvalue(), file_name=f"แนบท้าย_{s['name']}.docx", use_container_width=True)
                    with dl3: st.download_button(label="⬇️ 3. โหลด ม.23", data=bio_m23.getvalue(), file_name=f"ม23_{s['name']}.docx", use_container_width=True)
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
