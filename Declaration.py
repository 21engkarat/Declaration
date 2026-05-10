import streamlit as st
import pdfplumber
from docx import Document
import io
import zipfile
import re

st.title("📄 PDF Data Extractor to Word Template")
st.write("ดึงข้อมูลจาก PDF ไปใส่ใน Word Template อัตโนมัติ")

# ---------------------------------------------------------
# ฟังก์ชันสำหรับแทนที่คำใน Word (ทั้งในย่อหน้าปกติ และในตาราง)
# ---------------------------------------------------------
def replace_text_in_doc(doc, old_text, new_text):
    # ค้นหาและแทนที่ในข้อความปกติ
    for p in doc.paragraphs:
        if old_text in p.text:
            for run in p.runs:
                run.text = run.text.replace(old_text, new_text)
                
    # ค้นหาและแทนที่ในตาราง
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        for run in p.runs:
                            run.text = run.text.replace(old_text, new_text)

# ---------------------------------------------------------
# ส่วนที่ 1: รับไฟล์จากผู้ใช้งาน
# ---------------------------------------------------------
template_file = st.file_uploader("1. อัปโหลดไฟล์ Word Template (.docx)", type="docx")
uploaded_pdfs = st.file_uploader("2. อัปโหลดไฟล์ PDF (เลือกได้หลายไฟล์)", type="pdf", accept_multiple_files=True)

# ---------------------------------------------------------
# ส่วนที่ 2: เริ่มประมวลผล
# ---------------------------------------------------------
if template_file and uploaded_pdfs:
    if st.button("🚀 เริ่มดึงข้อมูลและสร้างเอกสาร"):
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            
            for pdf_file in uploaded_pdfs:
                st.info(f"กำลังประมวลผล: {pdf_file.name}")
                
                # 1. อ่านข้อความจาก PDF
                text_content = ""
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        text_content += page.extract_text() + "\n"
                
                # 2. ค้นหา "จำนวนเงิน" และ "ภาษี" ด้วย Regex (ค้นหาตัวเลขที่อยู่หลังคำสำคัญ)
                # หมายเหตุ: อาจจะต้องปรับคำว่า "จำนวนเงินที่จ่าย" ตามที่เขียนเป๊ะๆ ใน PDF ของคุณ
                amount_match = re.search(r'จำนวนเงินที่จ่าย\s*([\d,]+\.?\d*)', text_content)
                tax_match = re.search(r'ภาษีที่หักไว้\s*([\d,]+\.?\d*)', text_content)
                
                # ถ้าหาเจอให้เก็บค่าไว้ ถ้าหาไม่เจอให้ใส่ข้อความแจ้งเตือน
                amount_val = amount_match.group(1) if amount_match else "(ไม่พบยอดเงิน)"
                tax_val = tax_match.group(1) if tax_match else "(ไม่พบยอดภาษี)"
                
                # 3. เปิดไฟล์ Word Template ขึ้นมา
                doc = Document(template_file)
                
                # 4. แทนที่คำว่า XXX และ YYY ด้วยข้อมูลที่ดึงมาได้
                replace_text_in_doc(doc, "XXX", amount_val)
                replace_text_in_doc(doc, "YYY", tax_val)
                
                # 5. เซฟไฟล์ Word ลงหน่วยความจำ และจับใส่ ZIP
                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                
                new_filename = pdf_file.name.replace(".pdf", "_สำเร็จ.docx")
                zip_file.writestr(new_filename, doc_buffer.getvalue())

        st.success("✅ ประมวลผลเสร็จสิ้นทุกไฟล์!")
        
        # ปุ่มดาวน์โหลด
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Word ทั้งหมด (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="completed_documents.zip",
            mime="application/zip"
        )