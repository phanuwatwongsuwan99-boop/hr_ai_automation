import streamlit as st
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd
from PIL import Image

# 1. ตั้งค่าโครงสร้างหน้าเว็บและธีมส่วนหัว
st.set_page_config(
    page_title="HR AI Automation System", 
    page_icon="🤖", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ตกแต่ง CSS ตกแต่งปุ่มและฟอนต์เพิ่มเติมให้ดูพรีเมียมขึ้น
st.markdown("""
    <style>
    .main-title {
        font-size: 2.5rem !important;
        font-weight: 700 !important;
        color: #1E3A8A;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 1.1rem !important;
        color: #4B5563;
        margin-bottom: 25px;
    }
    .step-box {
        background-color: #F3F4F6;
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid #3B82F6;
        height: 100%;
    }
    div.stButton > button:first-child {
        background-color: #1E3A8A !important;
        color: white !important;
        border-radius: 8px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        border: none !important;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1) !important;
        transition: all 0.2s !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #2563EB !important;
        transform: translateY(-1px) !important;
    }
    </style>
""", unsafe_allow_html=True)

# เรียกใช้งาน API Key จากระบบหลังบ้านอย่างปลอดภัย
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error("❌ ไม่พบ API Key ในระบบหลังบ้าน กรุณาตั้งค่าใน Advanced Settings ของ Streamlit Cloud")
    st.stop()

TEMPLATES_FOLDER = "hr_templates"

# ฟังก์ชันเรียนรู้โครงสร้างคอลัมน์
def learn_templates():
    structures = {}
    if not os.path.exists(TEMPLATES_FOLDER):
        return structures
    for filename in os.listdir(TEMPLATES_FOLDER):
        if filename.endswith((".xlsx", ".xls")) and not filename.startswith((".", "~$")):
            file_path = os.path.join(TEMPLATES_FOLDER, filename)
            try:
                wb = load_workbook(file_path, data_only=True)
                ws = wb.active
                headers = [str(cell.value).strip() for cell in ws[1] if cell.value is not None]
                
                existing_rows = []
                for row_idx in range(2, ws.max_row + 1):
                    row_data = {}
                    has_data = False
                    for col_idx, header in enumerate(headers, 1):
                        val = ws.cell(row=row_idx, column=col_idx).value
                        val_str = str(val).strip() if val is not None else ""
                        row_data[header] = val_str
                        if val_str: has_data = True
                    if has_data:
                        row_data["__row_index__"] = row_idx
                        existing_rows.append(row_data)
                        
                if headers:
                    structures[filename] = {"headers": headers, "existing_data": existing_rows}
            except:
                pass
    return structures

# --- ส่วนจัดแสดงหน้าต่างเว็บแอป (UI) ---

st.markdown('<div class="main-title">🤖 HR AI Automation System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">ระบบผู้ช่วยอัจฉริยะ คัดแยกประเภทเอกสารรายได้ และคีย์ข้อมูลลงตารางพนักงานอัตโนมัติ</div>', unsafe_allow_html=True)

st.markdown("### 📋 ขั้นตอนการทำงานง่าย ๆ ใน 3 สเต็ป")
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    st.markdown('<div class="step-box"><b>1. ตรวจสอบแม่แบบ</b><br>ระบบจะดึงไฟล์ตารางพนักงานหลักจากระบบหลังบ้านมาเตรียมรอไว้ หากยังไม่มีข้อมูลพนักงานสามารถดาวน์โหลดตารางไปกรอกก่อนได้</div>', unsafe_allow_html=True)
with col_s2:
    st.markdown('<div class="step-box"><b>2. อัปโหลดหลักฐานดิบ</b><br>โยนไฟล์ Excel รายได้ หรือถ่ายรูปรูปภาพสรุปยอดคอมมิชชั่น / OT ที่ได้มาจากไลน์หรือแผนกอื่น ๆ เข้าสู่ระบบได้ทันที</div>', unsafe_allow_html=True)
with col_s3:
    st.markdown('<div class="step-box"><b>3. AI คีย์ให้อัตโนมัติ</b><br>สมองกล AI จะอ่านสแกนชื่อและรหัสพนักงาน แล้วนำตัวเลขไปหยอดในช่องที่เว้นว่างไว้ของคนๆ นั้นให้ตรงแถวเป๊ะๆ</div>', unsafe_allow_html=True)

st.write("")
st.write("")

col_left, col_right = st.columns([1, 2], gap="large")

with col_left:
    st.markdown("### 📥 คลังตารางแม่แบบ (Templates)")
    templates_available = learn_templates()
    
    if templates_available:
        st.success(f"🟢 ระบบตรวจพบตารางพร้อมใช้งาน {len(templates_available)} ไฟล์")
        selected_t_download = st.selectbox("เลือกตารางแม่แบบเพื่อดูโครงสร้าง:", list(templates_available.keys()))
        st.caption(f"คอลัมน์ในตาราง: {str(templates_available[selected_t_download]['headers'])}")
        
        t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
        with open(t_download_path, "rb") as f:
            st.download_button(
                label=f"📥 ดาวน์โหลดไฟล์แม่แบบนี้ไปใช้",
                data=f,
                file_name=selected_t_download,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.warning("⚠️ ไม่พบไฟล์ตารางในโฟลเดอร์ hr_templates")

with col_right:
    st.markdown("### 📤 นำเข้าข้อมูลดิบเพื่อประมวลผล")
    uploaded_file = st.file_uploader(
        "ลากไฟล์ข้อมูลรายได้มาวางตรงนี้ (รองรับ Excel: .xlsx และไฟล์รูปภาพ: .png, .jpg, .jpeg)", 
        type=["xlsx", "xls", "png", "jpg", "jpeg"]
    )
    
    if uploaded_file is not None:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        is_image = file_ext in ["png", "jpg", "jpeg"]
        
        with st.container(border=True):
            st.markdown(f"**📄 ชื่อไฟล์:** `{uploaded_file.name}` | **ประเภท:** `{file_ext.upper()}`")
            if is_image:
                image_obj = Image.open(uploaded_file)
                st.image(image_obj, caption="🖼️ พรีวิวรูปภาพหลักฐานที่นำเข้าสู่ระบบ", width=350)
        
        st.write("")
        if st.button("🚀 สั่งให้ AI เริ่มแมตช์และคีย์ข้อมูลลงตาราง", use_container_width=True):
            with st.spinner("🤖 AI กำลังกวาดสายตาอ่านข้อมูลและประมวลผลลงช่องตาราง... โปรดรอสักครู่"):
                try:
                    ai_contents = []
                    if is_image:
                        uploaded_file.seek(0)
                        ai_contents.append(Image.open(uploaded_file))
                        input_desc = "รูปภาพหลักฐาน"
                    else:
                        wb_in = load_workbook(uploaded_file, data_only=True)
                        input_text = ""
                        for sheet_name in wb_in.sheetnames:
                            ws_in = wb_in[sheet_name]
                            input_text += f"\n[หน้าตาราง: {sheet_name}]\n"
                            for row in ws_in.iter_rows(values_only=True):
                                row_clean = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
                                if row_clean:
                                    input_text += " | ".join(row_clean) + "\n"
                        ai_contents.append(f"[ข้อมูลเอกสารดิบ]\n{input_text}")
                        input_desc = "ตาราง Excel อินพุต"

                    blueprint = json.dumps(templates_available, ensure_ascii=False, indent=2)
                    prompt = f"""
                    คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลพนักงานจากเอกสารดิบที่แนบมา
                    แล้วนำข้อมูลไปเติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

                    นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
                    {blueprint}

                    คำสั่ง:
                    1. วิเคราะห์ว่าเนื้อหาข้อมูลอินพุตนี้ สอดคล้องหรือควรนำไปเติมในไฟล์เทมเพลตไหนมากที่สุด (ตอบชื่อไฟล์ในช่อง "selected_template")
                    2. ตรวจสอบ 'รหัสพนักงาน' หรือ 'รายชื่อ' เพื่อจับคู่ระหว่างข้อมูลในหลักฐานอินพุต กับพนักงานเดิมในตารางให้ตรงคน
                    3. ดึงค่าในเอกสารดิบ/รูปภาพที่ตรงกับหัวตารางปลายทาง (เช่น ค่าคอมมิชชั่น, OT) นำไปเติมลงในแถวที่มีตัวเลข "__row_index__" ของพนักงานคนนั้นๆ

                    ให้ตอบกลับเป็นรูปแบบ JSON โครงสร้างแบบนี้เท่านั้น ห้ามมีคำอธิบายอื่นเด็ดขาด:
                    {{
                      "selected_template": "ชื่อไฟล์เทมเพลตปลายทาง.xlsx",
                      "updates": [
                         {{
                           "row_index": 2, 
                           "data_to_fill": {{
                              "ชื่อคอลัมน์": "ค่าตัวเลขหรือข้อความที่ดึงได้"
                           }}
                         }}
                      ]
                    }}
                    """
                    
                    ai_contents.insert(0, prompt)
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=ai_contents)
                    clean_text = response.text.replace("```json", "").replace("```", "").strip()
                    ai_result = json.loads(clean_text)
                    
                    template_name = ai_result.get("selected_template")
                    updates = ai_result.get("updates", [])
                    
                    template_path = os.path.join(TEMPLATES_FOLDER, template_name)
                    wb_target = load_workbook(template_path)
                    ws_target = wb_target.active
                    
                    header_map = {str(ws_target.cell(row=1, column=c).value).strip(): c for c in range(1, ws_target.max_column + 1)}
                    preview_data = []
                    
                    for item in updates:
                        r_idx = item.get("row_index")
                        data_to_fill = item.get("data_to_fill", {})
                        
                        emp_name_col = header_map.get("รายชื่อ") or header_map.get("ชื่อพนักงาน") or 1
                        emp_name_str = ws_target.cell(row=r_idx, column=emp_name_col).value
                        
                        for h_name, new_val in data_to_fill.items():
                            if h_name in header_map:
                                # --- แก้ไขจุดบั๊กจาก new_value เป็น new_val เรียบร้อยครับ ---
                                ws_target.cell(row=r_idx, column=header_map[h_name]).value = new_val
                                preview_data.append({
                                    "แถวตารางที่": r_idx, 
                                    "รายชื่อพนักงาน": emp_name_str, 
                                    "หัวข้อคอลัมน์ที่เติม": h_name, 
                                    "จำนวนเงิน / ข้อมูลใหม่": new_val
                                })

                    if preview_data:
                        st.balloons()
                        st.success(f"🎯 AI วิเคราะห์สำเร็จ! นำข้อมูลจาก {input_desc} วิ่งไปกรอกลงไฟล์ '{template_name}' เรียบร้อยแล้วค่ะ")
                        
                        st.markdown("### 👀 ตารางสรุปผลข้อมูลที่ถูกอัปเดต (Data Preview)")
                        df_preview = pd.DataFrame(preview_data)
                        st.dataframe(
                            df_preview.style.set_properties(**{'background-color': '#EFF6FF', 'color': '#1E3A8A'}),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        out_stream = io.BytesIO()
                        wb_target.save(out_stream)
                        out_stream.seek(0)
                        
                        st.write("")
                        st.download_button(
                            label="📥 คลิกตรงนี้เพื่อดาวน์โหลดไฟล์ Excel ปลายทางที่กรอกข้อมูลเสร็จสมบูรณ์",
                            data=out_stream,
                            file_name=f"HR_UPDATED_{template_name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ AI ตรวจสอบไฟล์แล้ว แต่ไม่พบรายชื่อพนักงานคนไหนในหลักฐานตรงกับรายชื่อพนักงานในระบบเลย กรุณาตรวจสอบรายชื่ออีกครั้ง")
                        
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลระบบ: {e}")
