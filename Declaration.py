import streamlit as st
import pdfplumber
from docx import Document
import io
import re

st.title("📄 PDF to Word (Direct Download)")
st.write("ดึงข้อมูลจาก PDF ลง Template Word โดยตรง ไม่ต้องผ่าน ZIP")

# ฟังก์ชันสำหรับแทนที่คำใน Word
def replace_text_in_doc(doc, old_text, new_text):
    for p in doc.paragraphs:
        if old_text in p.text:
            for run in p.runs:
                run.text = run.text.replace(old_text, new_text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            run.text = run.text.replace(old_text, new_text)

# 1. ส่วนรับไฟล์
template_file = st.file_uploader("1. อัปโหลด Word Template (.docx)", type="docx")
uploaded_pdfs = st.file_uploader("2. อัปโหลด PDF", type="pdf", accept_multiple_files=True)

if template_file and uploaded_pdfs:
    st.subheader("📦 ผลลัพธ์การประมวลผล")
    
    for pdf_file in uploaded_pdfs:
        # อ่านข้อความจาก PDF
        text_content = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text_content += page.extract_text() + "\n"
        
        # --- ส่วนการดึงข้อมูลแบบ "อยู่ด้านล่างหัวข้อ" ---
        # ใช้คำสั่งค้นหาคำสำคัญ แล้วตามด้วยตัวเลขที่อยู่ในบรรทัดถัดไป
        # \s*[\n\r]\s* หมายถึง ค้นหาช่องว่างและตัวขึ้นบรรทัดใหม่
        amount_pattern = r'จำนวนเงินที่จ่าย\s*[\n\r]\s*([\d,]+\.?\d*)'
        tax_pattern = r'ภาษีที่หักไว้\s*[\n\r]\s*([\d,]+\.?\d*)'
        
        amount_match = re.search(amount_pattern, text_content)
        tax_match = re.search(tax_pattern, text_content)
        
        amount_val = amount_match.group(1) if amount_match else "ไม่พบข้อมูล"
        tax_val = tax_match.group(1) if tax_match else "ไม่พบข้อมูล"
        
        # สร้าง Word ใหม่จาก Template
        template_file.seek(0) # รีเซ็ต cursor ของไฟล์ template
        doc = Document(template_file)
        replace_text_in_doc(doc, "XXX", amount_val)
        replace_text_in_doc(doc, "YYY", tax_val)
        
        # บันทึกลงหน่วยความจำ
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        # 2. สร้างปุ่มดาวน์โหลดแยกตามไฟล์ (ไม่ต้องโหลด ZIP)
        new_filename = pdf_file.name.replace(".pdf", "_Success.docx")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.write(f"✅ {pdf_file.name} (เงิน: {amount_val} / ภาษี: {tax_val})")
        with col2:
            st.download_button(
                label="ดาวน์โหลดไฟล์ Word",
                data=doc_io,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=pdf_file.name # ป้องกันปุ่มซ้ำกัน
            )
