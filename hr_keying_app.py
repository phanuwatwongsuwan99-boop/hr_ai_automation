import streamlit as st
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd
from PIL import Image

# 1. ตั้งค่าโครงสร้างหน้าเว็บสไตล์ Global Dashboard
st.set_page_config(
    page_title="HR Smart Matrix - AI Automation", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ตกแต่ง CSS ให้ดูทันสมัย ระดับสากล (Clean, Modern, Corporate)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
    }
    .card-box {
        background-color: #FFFFFF;
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
    .badge-personal {
        background-color: #E0F2FE;
        color: #0369A1;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        border: 1px solid #BAE6FD;
    }
    .badge-commission {
        background-color: #DCFCE7;
        color: #15803D;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        border: 1px solid #BBF7D0;
    }
    .badge-ot {
        background-color: #F3E8FF;
        color: #6B21A8;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        border: 1px solid #E9D5FF;
    }
    .badge-generic {
        background-color: #F3F4F6;
        color: #374151;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        border: 1px solid #E5E7EB;
    }
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%) !important;
        color: white !important;
        border-radius: 10px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        padding: 12px 30px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3) !important;
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

# --- ส่วนการแสดงผล (Global UI) ---

# แถบหัวข้อใหญ่หรูหราสไตล์เว็บอินเตอร์
st.markdown("""
    <div class="main-header">
        <h1 style='margin:0; font-size:2.2rem; font-weight:700;'>⚡ HR Smart Matrix Hub</h1>
        <p style='margin:5px 0 0 0; opacity:0.85; font-size:1rem;'>ระบบประมวลผลและคัดแยกหมวดหมู่เอกสารพนักงานอัจฉริยะด้วยสมองกล AI</p>
    </div>
""", unsafe_allow_html=True)

templates_available = learn_templates()

# ใช้ระบบ Grid แบ่งซ้าย-ขวาแบบ Dashboard สากล
col_control, col_display = st.columns([1, 2], gap="large")

# 🖥️ ฝั่งซ้าย: แผงควบคุมและจัดการไฟล์ (Control Workspace)
with col_control:
    st.markdown("### 🎛️ แผงควบคุมระบบ (Workspace)")
    
    with st.container():
        st.markdown('<div class="card-box">', unsafe_allow_html=True)
        st.markdown("📂 **1. คลังตารางแม่แบบปลายทาง**")
        if templates_available:
            selected_t_download = st.selectbox("เลือกแม่แบบเพื่อตรวจสอบ:", list(templates_available.keys()))
            st.caption(f"📌 คอลัมน์ที่ตรวจพบ: {', '.join(templates_available[selected_t_download]['headers'])}")
            
            t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
            with open(t_download_path, "rb") as f:
                st.download_button(
                    label="📥 ดาวน์โหลดตารางแม่แบบเปล่า",
                    data=f,
                    file_name=selected_t_download,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_btn"
                )
        else:
            st.error("ไม่พบไฟล์แม่แบบในระบบหลังบ้าน")
        st.markdown('</div>', unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="card-box">', unsafe_allow_html=True)
        st.markdown("📤 **2. นำเข้าไฟล์หลักฐานดิบ**")
        uploaded_file = st.file_uploader(
            "ลากและวางหลักฐานตรงนี้ (.xlsx, .png, .jpg, .jpeg)", 
            type=["xlsx", "xls", "png", "jpg", "jpeg"]
        )
        
        if uploaded_file is not None:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            is_image = file_ext in ["png", "jpg", "jpeg"]
            st.caption(f"⚡ พร้อมประมวลผลไฟล์: {uploaded_file.name}")
            if is_image:
                image_obj = Image.open(uploaded_file)
                st.image(image_obj, caption="พรีวิวหลักฐาน", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# 📊 ฝั่งขวา: พื้นที่จัดแสดงผลลัพธ์และคัดแยกหมวดหมู่ (Result Workspace)
with col_display:
    st.markdown("### 📑 ผลการวิเคราะห์และคัดแยกหมวดหมู่ (Analysis Output)")
    
    if uploaded_file is None:
        st.info("💡 ระบบสแตนด์บาย: กรุณาอัปโหลดไฟล์หลักฐานและกดปุ่มประมวลผลที่แผงควบคุมด้านซ้ายมือเพื่อเริ่มงานค่ะ")
    else:
        if st.button("🚀 เริ่มวิเคราะห์ข้อมูลสัมพันธ์ด้วย AI"):
            with st.spinner("🤖 AI กำลังสแกนตรวจสอบประเภทข้อมูล คัดแยกหมวดหมู่ และลงบันทึกช่องตาราง..."):
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
                    
                    # สั่งให้ AI คัดแยกหมวดหมู่เอกสารเพิ่มเข้ามาในโครงสร้าง JSON ปลายทางด้วย
                    prompt = f"""
                    คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลพนักงานจากเอกสารดิบที่แนบมา
                    แล้วนำข้อมูลไปเติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

                    นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
                    {blueprint}

                    คำสั่งในการวิเคราะห์:
                    1. ตรวจสอบเนื้อหาของเอกสารดิบ/รูปภาพ แล้วระบุว่ามันเป็นเอกสาร "หมวดหมู่ (Category)" ใด โดยเลือกจากหมวดหมู่เหล่านี้เท่านั้น:
                       - "PERSONAL_DATA" (หากเป็นบัตรประชาชน, ทะเบียนบ้าน, ใบขับขี่, ข้อมูลส่วนตัว)
                       - "COMMISSION" (หากเป็นไฟล์ค่าคอมมิชชั่น, ยอดขาย, ค่านายหน้า)
                       - "OVERTIME" (หากเป็นเวลาเข้างาน, สรุปโอที, ข้อมูลเวลาทำงาน)
                       - "GENERIC_HR" (หากเป็นเอกสารอื่นๆ)
                    2. วิเคราะห์ว่าเนื้อหาข้อมูลอินพุตนี้ สอดคล้องหรือควรนำไปเติมในไฟล์เทมเพลตไหนมากที่สุด (ตอบชื่อไฟล์ในช่อง "selected_template")
                    3. จับคู่รหัสหรือชื่อพนักงาน แล้วนำค่าไปเติมลงในแถวที่มีตัวเลข "__row_index__" ของพนักงานคนนั้นๆ

                    ให้ตอบกลับเป็นรูปแบบ JSON โครงสร้างแบบนี้เท่านั้น ห้ามมีคำอธิบายอื่นเด็ดขาด:
                    {{
                      "category": "หมวดหมู่ที่เลือกตามข้อ 1 เช่น COMMISSION",
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
                    
                    category = ai_result.get("category", "GENERIC_HR")
                    template_name = ai_result.get("selected_template")
                    updates = ai_result.get("updates", [])
                    
                    # ส่วนแสดงกล่องหมวดหมู่สีสวยงามบนหน้าเว็บแอป (UX Highlight)
                    st.markdown("#### 🎯 ผลการคัดแยกหมวดหมู่เอกสาร")
                    if category == "PERSONAL_DATA":
                        st.markdown('<span class="badge-personal">🆔 หมวดหมู่: ข้อมูลส่วนตัว / บัตรประชาชนพนักงาน</span>', unsafe_allow_html=True)
                    elif category == "COMMISSION":
                        st.markdown('<span class="badge-commission">💰 หมวดหมู่: รายได้ / ค่าคอมมิชชั่นพนักงาน</span>', unsafe_allow_html=True)
                    elif category == "OVERTIME":
                        st.markdown('<span class="badge-ot">⏱️ หมวดหมู่: เวลาทำงาน / ข้อมูลโอที (OT)</span>', unsafe_allow_html=True)
                    else:
                        st.markdown('<span class="badge-generic">📦 หมวดหมู่: เอกสารทั่วไป / ทั่วไปในระบบ HR</span>', unsafe_allow_html=True)
                        
                    st.write("")
                    
                    # ดำเนินการคีย์ข้อมูลลงไฟล์ Excel
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
                                ws_target.cell(row=r_idx, column=header_map[h_name]).value = new_val
                                preview_data.append({
                                    "แถวที่": r_idx, 
                                    "รายชื่อพนักงาน": emp_name_str, 
                                    "คอลัมน์ที่คีย์ลงตาราง": h_name, 
                                    "ข้อมูลที่เติมสำเร็จ": new_val
                                })

                    if preview_data:
                        st.balloons()
                        st.success(f"กรอกข้อมูลลงในไฟล์ตารางหลัก '{template_name}' สำเร็จแล้วค่ะ")
                        
                        st.markdown("#### 👀 ตารางตรวจสอบความถูกต้องก่อนบันทึก (Preview)")
                        df_preview = pd.DataFrame(preview_data)
                        st.dataframe(
                            df_preview.style.set_properties(**{'background-color': '#FAFAFA', 'color': '#0F172A'}),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        out_stream = io.BytesIO()
                        wb_target.save(out_stream)
                        out_stream.seek(0)
                        
                        st.markdown("#### 📥 ดาวน์โหลดผลลัพธ์เพื่อนำไปใช้งาน")
                        st.download_button(
                            label="📥 คลิกตรงนี้เพื่อโหลดไฟล์ Excel สำเร็จรูป",
                            data=out_stream,
                            file_name=f"HR_EXPORT_{template_name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ AI ตรวจสอบเอกสารนี้แล้ว แต่ไม่พบรายชื่อพนักงานคนไหนในหลักฐานตรงกับรายชื่อพนักงานในระบบเลย กรุณาตรวจสอบอีกครั้ง")
                        
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลระบบ: {e}")
