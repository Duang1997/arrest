import os
import subprocess
import streamlit as st
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from PIL import Image, ImageOps

st.set_page_config(
    page_title="ระบบบันทึกภาพและจัดทำแฟ้มผู้ต้องหา", page_icon="📋", layout="centered"
)

st.title("ระบบจัดทำบันทึกแนบท้ายภาพถ่ายผู้ถูกจับและผู้ถูกควบคุม")
st.markdown("---")

# ส่วนที่ 1: ข้อมูลผู้ต้องหา (เพิ่มช่องอายุ)
st.subheader("1. ข้อมูลผู้ต้องหา")
col1, col2, col3 = st.columns([2, 1, 2])
with col1:
  full_name = st.text_input("ชื่อ - นามสกุล").strip()
with col2:
  age = st.text_input("อายุ (ปี)").strip()
with col3:
  national_id = st.text_input(
      "เลขประจำตัวประชาชน (13 หลัก)", max_chars=13
  ).strip()

st.markdown("---")

# ส่วนที่ 2: บันทึกภาพถ่าย 4 ด้าน
st.subheader("2. บันทึกภาพถ่าย 4 ด้าน")
angles = {
    "pic_front": "ภาพหน้าตรง (pic_front)",
    "pic_left": "ภาพด้านซ้าย (pic_left)",
    "pic_right": "ภาพด้านขวา (pic_right)",
    "pic_back": "ภาพด้านหลัง (pic_back)",
}

captured_images = {}
for key, label in angles.items():
  st.markdown(f"**{label}**")
  input_method = st.radio(
      f"วิธีนำเข้า {key}",
      ("ถ่ายภาพจากกล้อง", "อัปโหลดไฟล์ภาพ"),
      key=f"method_{key}",
      horizontal=True,
  )
  if input_method == "ถ่ายภาพจากกล้อง":
    img_file = st.camera_input(f"ถ่าย {key}", key=f"cam_{key}")
  else:
    img_file = st.file_uploader(
        f"เลือกไฟล์ {key}", type=["jpg", "jpeg", "png"], key=f"file_{key}"
    )

  captured_images[key] = img_file
  st.markdown("---")


# ฟังก์ชันสร้างไฟล์ Word และแปลงเป็น PDF
def generate_files_word_and_pdf(nid, name, suspect_age, images):
  template_path = "template.docx"
  if not os.path.exists(template_path):
    st.error("ไม่พบไฟล์ template.docx ในระบบ กรุณาอัปโหลดไฟล์ Template")
    return None, None

  doc = DocxTemplate(template_path)
  image_context = {}
  temp_files = []

  for key, img_file in images.items():
    if img_file is not None:
      temp_path = f"temp_{key}.jpg"
      img = Image.open(img_file)
      img = ImageOps.exif_transpose(img)

      if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
      img.save(temp_path, "JPEG")
      temp_files.append(temp_path)

      # กำหนดความสูงรูปภาพเท่ากับ 6 เซนติเมตร
      image_context[key] = InlineImage(doc, temp_path, height=Cm(6))
    else:
      image_context[key] = ""

  context = {
      "suspect_name": name,
      "suspect_age": suspect_age,
      "suspect_id": nid,
      **image_context,
  }

  doc.render(context)

  base_name = f"ภาพแนบ-{nid}-{name}"
  docx_filename = f"{base_name}.docx"
  pdf_filename = f"{base_name}.pdf"

  doc.save(docx_filename)

  try:
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", docx_filename],
        check=True,
    )
  except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการแปลงไฟล์เป็น PDF: {e}")
    return docx_filename, None

  for tf in temp_files:
    if os.path.exists(tf):
      os.remove(tf)

  return docx_filename, pdf_filename


# ส่วนที่ 3: ปุ่มบันทึกและดาวน์โหลด
if st.button("บันทึกข้อมูลและสร้างไฟล์รายงาน", type="primary"):
  if not national_id or not full_name or not age:
    st.error("กรุณากรอกข้อมูล ชื่อ-นามสกุล อายุ และเลขประจำตัวประชาชนให้ครบถ้วน")
  else:
    with st.spinner("กำลังประมวลผลข้อมูล ปรับทิศทางภาพ จัดวาง และแปลงไฟล์..."):
      docx_file, pdf_file = generate_files_word_and_pdf(
          national_id, full_name, age, captured_images
      )

      if docx_file:
        st.success("สร้างไฟล์รายงานสำเร็จเรียบร้อยแล้ว")
        st.markdown("---")

        col_btn1, col_btn2 = st.columns(2)

        with col_btn1:
          with open(docx_file, "rb") as f:
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word (.docx)",
                data=f,
                file_name=docx_file,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        if pdf_file and os.path.exists(pdf_file):
          with col_btn2:
            with open(pdf_file, "rb") as f:
              st.download_button(
                  label="📥 ดาวน์โหลดไฟล์ PDF (.pdf)",
                  data=f,
                  file_name=pdf_file,
                  mime="application/pdf",
              )
