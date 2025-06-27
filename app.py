# --- โค้ดฉบับสมบูรณ์ (เวอร์ชัน 8.4 - สรุปผลแบบตารางละเอียด) ---

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re

# --- การตั้งค่า Session State ---
if 'file_uploader_key' not in st.session_state:
    st.session_state.file_uploader_key = 0

# --- การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="ตรวจสอบใบรายงานผลสอบเทียบอัตโนมัติ", layout="wide")

# --- ฟังก์ชันสำหรับโหลด CSS และฟอนต์ ---
def load_custom_css():
    st.markdown("""
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;600&display=swap" rel="stylesheet">
        
        <style>
            html, body, [class*="st-"], h1, h2, h3, h4, h5, h6 {
                font-family: 'Prompt', sans-serif !important;
            }
            .st-emotion-cache-1avpkeq { /* รูปแบบ Card */
                box-shadow: 0 4px 8px 0 rgba(0,0,0,0.02);
                transition: 0.3s;
                border-radius: 10px;
                padding: 1rem 1rem 1.5rem 1rem;
            }
            .footer {
                position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                background-color: #F0F2F6;
                color: #555;
                text-align: center;
                padding: 8px;
                font-size: 14px;
                z-index: 100;
                border-top: 1px solid #E6E6E6;
            }
        </style>
    """, unsafe_allow_html=True)

load_custom_css()

# --- ส่วนของฟังก์ชัน (Backend) ---
@st.cache_data
def find_database_file(directory="."):
    for filename in os.listdir(directory):
        if filename.startswith("ใบขอรับบริการ") and filename.endswith((".xlsx", ".xls")): return os.path.join(directory, filename)
    return None

@st.cache_data
def load_database(path):
    return pd.read_excel(path)

def extract_data_from_pdf(file_object):
    try:
        file_bytes = file_object.read(); doc = fitz.open(stream=file_bytes, filetype="pdf")
        extracted_data, full_text = {}, ""
        keyword_variations = {
            'Equipment': ['Equipment'], 'Manufacturer': ['Manufacturer'], 'Model': ['Model'],
            'Serial No.': ['Serial No.', 'SERIAL No'], 'ID No.': ['ID No.', 'ID CODE'],
            'Customer': ['Customer', 'SUBMITTED BY'], 'Calibration Date': ['Calibration Date', 'DATE OF CALIBRATION']
        }
        for page in doc:
            full_text += page.get_text() + "\n"; words = page.get_text("words")
            for standard_name, variations in keyword_variations.items():
                if standard_name in extracted_data: continue
                for keyword_text in variations:
                    for i, w in enumerate(words):
                        current_word_sequence = w[4]
                        if " " in keyword_text:
                            parts = keyword_text.split(" ");
                            if w[4].lower() == parts[0].lower():
                                match_complete = True
                                for j in range(1, len(parts)):
                                    if (i + j) < len(words) and words[i + j][4].lower() == parts[j].lower(): current_word_sequence += " " + words[i + j][4]
                                    else: match_complete = False; break
                                if not match_complete: continue
                        if current_word_sequence.lower() == keyword_text.lower():
                            keyword_y, keyword_x1 = w[1], words[i + len(keyword_text.split()) - 1][2]
                            value_text = "";
                            for v_word in words:
                                if v_word[0] > keyword_x1 and abs(v_word[1] - keyword_y) < 5: value_text += v_word[4] + " "
                            if value_text: extracted_data[standard_name] = value_text.strip(); break
                    if standard_name in extracted_data: break
        cert_pattern = r"((?:AC|DB|WB|CF|EB|PH|PP|LG|CH|DT)\d{8})"; cert_match = re.search(cert_pattern, full_text, re.IGNORECASE)
        if cert_match: extracted_data['Certificate No.'] = cert_match.group(1)
        else:
            cert_match_fallback = re.search(r"Certificate No\.\s*:\s*(\S+)", full_text, re.IGNORECASE)
            if cert_match_fallback: extracted_data['Certificate No.'] = cert_match_fallback.group(1)
        if extracted_data.get('Customer') and len(extracted_data.get('Customer', '')) < 30:
            customer_match = re.search(r"(?:Customer|SUBMITTED BY)\s*[:\n\s]*(.*?)(?:\nLocation of Calibration|\nENVIRONMENT|\nReceived Date)", full_text, re.DOTALL | re.IGNORECASE)
            if customer_match: extracted_data['Customer'] = " ".join(customer_match.group(1).split()).strip()
        final_data = {};
        for key, value in extracted_data.items():
            if not isinstance(value, str): final_data[key] = value; continue
            cleaned_value = value.strip().lstrip(':').strip()
            if key in ['Serial No.', 'ID No.'] and ':' in cleaned_value: cleaned_value = cleaned_value.split(':')[-1].strip()
            if key == 'Calibration Date':
                date_match = re.search(r'(\d{1,2}\s+\w+\s+\d{4})', cleaned_value)
                if date_match: cleaned_value = date_match.group(1)
            final_data[key] = cleaned_value
        doc.close(); return final_data
    except Exception as e: st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ PDF: {e}"); return None

def verify_report(pdf_data, db_dataframe):
    if not pdf_data or not pdf_data.get('Certificate No.'): return {"status": "error", "message": "ไม่สามารถสกัด Certificate No. จาก PDF ได้"}
    cert_no_from_pdf = pdf_data.get('Certificate No.')
    if not cert_no_from_pdf: return {"status": "error", "message": "Certificate No. ใน PDF เป็นค่าว่าง"}
    try: record = db_dataframe[db_dataframe['Certificate No.'] == cert_no_from_pdf]
    except KeyError: return {"status": "error", "message": "ไม่พบคอลัมน์ 'Certificate No.' ในไฟล์ Excel"}
    if record.empty: return {"status": "invalid", "message": f"ไม่พบข้อมูลของใบรับรอง '{cert_no_from_pdf}'"}
    record_data, comparison_results, mismatched_fields = record.iloc[0], {}, []
    fields_to_check = ['Equipment', 'Manufacturer', 'Model', 'Serial No.', 'ID No.', 'Customer', 'Calibration Date']
    for field in fields_to_check:
        pdf_val, db_val = str(pdf_data.get(field, "")).strip(), str(record_data.get(field, "")).strip()
        if field == 'Calibration Date':
            try: db_val = pd.to_datetime(db_val).strftime('%d %b %Y')
            except Exception: pass
        pdf_comp, db_comp = " ".join(pdf_val.lower().split()), " ".join(db_val.lower().split())
        match = (pdf_comp == db_comp)
        comparison_results[field] = {'pdf': pdf_val, 'db': db_val, 'match': match}
        if not match: mismatched_fields.append(field)
    if not mismatched_fields: return {"status": "valid", "message": "ข้อมูลถูกต้องตรงกันทุกรายการ", "details": comparison_results}
    else: return {"status": "invalid", "message": f"ข้อมูลไม่ตรงกันในฟิลด์: {', '.join(mismatched_fields)}", "details": comparison_results}

# --- ส่วนหลักของโปรแกรม ---
st.title("ตรวจสอบใบรายงานผลสอบเทียบอัตโนมัติ")
st.write("อัปโหลดไฟล์ใบรายงานผล PDF ของคุณเพื่อเริ่มต้นการตรวจสอบ")
st.divider()

db_path = find_database_file()

if db_path:
    df_database = load_database(db_path)
    uploaded_files = st.file_uploader(
        "เลือกไฟล์ PDF ที่ต้องการ", type="pdf", accept_multiple_files=True, 
        label_visibility="collapsed", key=f"file_uploader_{st.session_state.file_uploader_key}"
    )
    if uploaded_files:
        st.divider()
        col1, col2 = st.columns([3, 1])
        with col1:
            st.info(f"พบ {len(uploaded_files)} ไฟล์ จะเริ่มทำการตรวจสอบตามลำดับ")
        with col2:
            if st.button("ล้างไฟล์ทั้งหมด", use_container_width=True, type="primary"):
                st.session_state.file_uploader_key += 1; st.rerun()
        
        st.subheader("ผลการตรวจสอบ:")
        
        # --- จุดที่แก้ไข 1: สร้าง dictionary สำหรับเก็บรายละเอียดไฟล์ที่ผิดพลาด ---
        mismatched_files_summary = {}

        for uploaded_file in uploaded_files:
            with st.container(border=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**ไฟล์:** `{uploaded_file.name}`")
                
                with st.spinner('⏳ กำลังวิเคราะห์ไฟล์...'):
                    uploaded_file.seek(0)
                    pdf_data = extract_data_from_pdf(uploaded_file)
                
                if not pdf_data:
                    st.warning("ไม่สามารถสกัดข้อมูลจากไฟล์นี้ได้อย่างสมบูรณ์"); continue

                verification_result = verify_report(pdf_data, df_database)
                
                with col2:
                    if verification_result['status'] == 'valid': 
                        st.success(f"**สถานะ:** ตรงกันทุกรายการ")
                    elif verification_result['status'] == 'invalid': 
                        st.error(f"**สถานะ:** ข้อมูลไม่ตรงกัน")
                        # --- จุดที่แก้ไข 2: เก็บรายละเอียดของไฟล์ที่ไม่ตรงกัน ---
                        mismatched_files_summary[uploaded_file.name] = verification_result.get('details', {})
                    else: 
                        st.warning(f"**สถานะ:** เกิดข้อผิดพลาด")
                
                with st.expander("▶ คลิกเพื่อดูรายละเอียด"):
                    if 'details' in verification_result:
                        details_df = pd.DataFrame([(f, values['pdf'], values['db'], '✅' if values['match'] else '❌') for f, values in verification_result['details'].items()], columns=["ฟิลด์", "จาก PDF", "จากฐานข้อมูล", "ผลลัพธ์"])
                        st.table(details_df)
                    else:
                         st.warning(verification_result['message'])
        
        # --- จุดที่แก้ไข 3: แสดงผลสรุปแบบตารางละเอียด ---
        if mismatched_files_summary:
            st.divider()
            st.warning("สรุปรายการไฟล์และฟิลด์ที่ข้อมูลไม่ถูกต้อง:")
            
            for filename, details in mismatched_files_summary.items():
                st.markdown(f"**ไฟล์:** `{filename}`")
                
                mismatched_rows = []
                for field, data in details.items():
                    if not data['match']:
                        mismatched_rows.append({
                            "ฟิลด์": field,
                            "ข้อมูลจาก PDF": data['pdf'],
                            "ข้อมูลจากฐานข้อมูล": data['db']
                        })
                
                if mismatched_rows:
                    summary_df = pd.DataFrame(mismatched_rows)
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                st.markdown("---") # คั่นระหว่างไฟล์

        elif uploaded_files:
            st.divider()
            st.success("🎉 ยอดเยี่ยม! ไฟล์ทั้งหมดที่ตรวจสอบถูกต้องตรงกันทุกรายการ")
            st.balloons()

else:
    st.error("ไม่พบไฟล์ฐานข้อมูล กรุณาวางไฟล์ Excel ที่ชื่อขึ้นต้นด้วย 'ใบขอรับบริการ' ในโฟลเดอร์เดียวกับโปรแกรม")


# --- Footer ---
if db_path:
    footer_text = f"สถานะฐานข้อมูล: พบไฟล์ `{os.path.basename(db_path)}`"
else:
    footer_text = "สถานะฐานข้อมูล: ไม่พบไฟล์ฐานข้อมูล"

st.markdown(f'<div class="footer">{footer_text}</div>', unsafe_allow_html=True)
