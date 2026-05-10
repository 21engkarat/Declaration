import streamlit as st
import pdfplumber
from docx import Document
import io
import zipfile

st.title("📌 PDF to Word Batch Converter")
st.write("อัปโหลดไฟล์ PDF หลายไฟล์เพื่อแปลงเป็น Word พร้อมกัน")

# 1. ส่วนการอัปโหลดไฟล์ (รองรับหลายไฟล์)
uploaded_files = st.file_uploader("เลือกไฟล์ PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # สร้าง Buffer สำหรับเก็บไฟล์ ZIP ในหน่วยความจำ
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for uploaded_file in uploaded_files:
            st.info(f"กำลังประมวลผล: {uploaded_file.name}")
            
            # อ่าน PDF
            text_content = ""
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text_content += page.extract_text() + "\n"
            
            # สร้าง Word ในหน่วยความจำ (ไม่ต้องบันทึกลง Disk)
            doc = Document()
            doc.add_heading(f'ข้อมูลจากไฟล์: {uploaded_file.name}', 0)
            doc.add_paragraph(text_content)
            
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            
            # เพิ่มไฟล์ Word ลงใน ZIP
            zip_file.writestr(uploaded_file.name.replace(".pdf", ".docx"), doc_buffer.getvalue())

    # 2. ปุ่มดาวน์โหลดไฟล์ ZIP ทั้งหมด
    st.success("ประมวลผลเสร็จสิ้น!")
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Word ทั้งหมด (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="converted_files.zip",
        mime="application/zip"
    )