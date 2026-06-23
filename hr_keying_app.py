import streamlit as st
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd
from PIL import Image

# 1. ตั้งค่าหน้าเว็บระดับพรีเมียม
st.set_page_config(
    page_title="HR Smart Matrix Hub", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ปรับปรุง CSS สไตล์โกบอล โทนสี แดง-ดำ-ขาว (Premium Red & Dark Theme)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #FAFAFA;
    }
    .main-header {
        background: linear-gradient(135deg, #111827 0%, #991B1B 100%);
        padding: 35px;
        border-radius: 12px;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        border-bottom: 4px solid #DC2626;
    }
    .card-box {
        background-color: #FFFFFF;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 20px;
    }
    .section-title {
        font-size: 1.2rem;
        font-weight: 600;
        color: #111827;
        margin-bottom: 15px;
        border-left: 4px solid #991B1B;
        padding-left: 10px;
    }
    /* สไตล์ของ Badge แยกตามหมวดหมู่ */
    .badge-personal {
        background-color: #FEF2F2;
        color: #991B1B;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        display: inline-block;
        border: 1px solid #FEE2E2;
    }
    .badge-commission {
        background-color: #F0FDF4;
        color: #166534;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        display: inline-block;
        border: 1px solid #DCFCE7;
    }
    .badge-ot {
        background-color: #FFFBEB;
        color: #92400E;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        display: inline-block;
        border: 1px solid #FEF3C7;
    }
    /* ปรับแต่งปุ่มกดหลัก */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #111827 0%, #991B1B 100%) !important;
        color: white !important;
        border-radius: 6px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        padding: 12px 30px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 4px 12px rgba(153, 27, 27, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(135deg, #1F2937 0%, #B91C1C 100%) !important;
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

# --- ส่วนการแสดงผล (Global Red & Dark UI) ---

st.markdown("""
    <div class="main-header">
        <h1 style='margin:0; font-size:2.2rem; font-weight:700; letter-spacing: -0.5px;'>⚡ HR Smart Matrix Hub</h1>
        <p style='margin:6px 0 0 0; opacity:0.85; font-size:1rem; font-weight: 300;'>ระบบประมวลผล คัดแยกหมวดหมู่อัจฉริยะ และจับคู่ข้อมูลพนักงานอัตโนมัติด้วย AI</p>
    </div>
""", unsafe_allow_html=True)

templates_available = learn_templates()

# ใช้ระบบ Grid แบ่งซ้าย-ขวาแบบ Dashboard สากล
col_control, col_display = st.columns([1, 2], gap="large")

# 🖥️ ฝั่งซ้าย: Workspace ควบคุมระบบ
with col_control:
    st.markdown('<div class="section-title">🎛️ แผงควบคุม (Workspace)</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="card-box">', unsafe_allow_html=True)
        st.markdown("📁 **1. คลังตารางแม่แบบปลายทาง**")
        if templates_available:
            selected_t_download = st.selectbox("เลือกตารางแม่แบบปลายทาง:", list(templates_available.keys()))
            st.caption(f"📌 หัวตาราง: {', '.join(templates_available[selected_t_download]['headers'])}")
            
            t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
            with open(t_download_path, "rb") as f:
                st.download_button(
                    label="📥 ดาวน์โหลดตารางแม่แบบเปล่า",
                    data=f,
                    file_name=selected_t_download,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("ไม่พบไฟล์แม่แบบในระบบหลังบ้าน")
        st.markdown('</div>', unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="card-box">', unsafe_allow_html=True)
        st.markdown("📤 **2. นำเข้าไฟล์หลักฐานดิบ**")
        uploaded_file = st.file_uploader(
            "ลากและวางไฟล์หรือรูปภาพตรงนี้", 
            type=["xlsx", "xls", "png", "jpg", "jpeg"]
        )
        
        if uploaded_file is not None:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            is_image = file_ext in ["png", "jpg", "jpeg"]
            if is_image:
                image_obj = Image.open(uploaded_file)
                st.image(image_obj, caption="รูปหลักฐานที่นำเข้าสู่ระบบ", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# 📊 ฝั่งขวา: พื้นที่แสดงผลการวิเคราะห์หมวดหมู่และข้อมูลสากล
with col_display:
    st.markdown('<div class="section-title">📑 ศูนย์แสดงผลข้อมูลวิเคราะห์ (Analysis Output)</div>', unsafe_allow_html=True)
    
    if uploaded_file is None:
        st.info("💡 ระบบพร้อมทำงาน: กรุณาเตรียมไฟล์และอัปโหลดข้อมูลดิบที่แถบเมนูด้านซ้ายมือเพื่อเริ่มการวิเคราะห์")
    else:
        if st.button("🚀 สั่งเริ่มวิเคราะห์ข้อมูลสัมพันธ์ด้วย AI", use_container_width=True):
            with st.spinner("🤖 สมองกล AI กำลังกวาดสายตาอ่านข้อมูลและประมวลผล..."):
                try:
                    ai_contents = []
                    if is_image:
                        uploaded_file.seek(0)
                        ai_contents.append(Image.open(uploaded_file))
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

                    blueprint = json.dumps(templates_available, ensure_ascii=False, indent=2)
                    
                    prompt = f"""
                    คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลพนักงานจากเอกสารดิบหรือรูปภาพที่แนบมา
                    แล้วนำข้อมูลไปจับคู่เติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

                    นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
                    {blueprint}

                    คำสั่งในการวิเคราะห์และจัดหมวดหมู่:
                    1. ระบุประเภทหมวดหมู่ (Category) ของหลักฐานนี้ โดยเลือกค่าต่อไปนี้: "PERSONAL_DATA", "COMMISSION", "OVERTIME", "GENERIC_HR"
                    2. คัดเลือกเทมเพลตปลายทางที่เหมาะสมที่สุดกรอกลงช่อง "selected_template"
                    3. ตรวจสอบรหัสหรือชื่อเพื่อทำการจับคู่ หากในตารางปลายทางช่องรายชื่อยังไม่มีค่า (เป็นค่าว่างหรือพนักงานใหม่) ให้ดึงข้อมูล "ชื่อ-นามสกุล" ที่สแกนได้จากหลักฐานมาเติมใส่ใน Preview ด้วย เพื่อป้องกันคำว่า None บนหน้าจอ

                    ให้ตอบกลับเป็นรูปแบบ JSON โครงสร้างแบบนี้เท่านั้น ห้ามมีคำอธิบายอื่นเด็ดขาด:
                    {{
                      "category": "หมวดหมู่เอกสาร",
                      "selected_template": "ชื่อไฟล์เทมเพลตปลายทาง.xlsx",
                      "updates": [
                         {{
                           "row_index": 2, 
                           "data_to_fill": {{
                              "ชื่อคอลัมน์": "ค่าที่ดึงได้"
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
                    
                    # 🎯 ส่วนที่ปรับปรุง: กล่อง Badge หมวดหมู่ออกแบบใหม่หรูหรา
                    st.write("")
                    if category == "PERSONAL_DATA":
                        st.markdown('<div class="badge-personal">🆔 เอกสารคัดแยกสำเร็จ: หมวดหมู่ข้อมูลส่วนตัว / บัตรประชาชนพนักงาน</div>', unsafe_allow_html=True)
                    elif category == "COMMISSION":
                        st.markdown('<div class="badge-commission">💰 เอกสารคัดแยกสำเร็จ: หมวดหมู่รายได้สัมพันธ์ / ค่าคอมมิชชั่นพนักงาน</div>', unsafe_allow_html=True)
                    elif category == "OVERTIME":
                        st.markdown('<div class="badge-ot">⏱️ เอกสารคัดแยกสำเร็จ: หมวดหมู่ข้อมูลเวลาทำงาน / โอทีพนักงาน (OT)</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="badge-generic">📦 เอกสารคัดแยกสำเร็จ: หมวดหมู่ทั่วไปทั่วไปในระบบ HR</div>', unsafe_allow_html=True)
                    st.write("")
                    
                    template_path = os.path.join(TEMPLATES_FOLDER, template_name)
                    wb_target = load_workbook(template_path)
                    ws_target = wb_target.active
                    
                    header_map = {str(ws_target.cell(row=1, column=c).value).strip(): c for c in range(1, ws_target.max_column + 1)}
                    
                    # 🎯 ส่วนที่ปรับปรุง: ปรับ Logic ข้อมูลให้ออกมาเป็นตารางแนวนอน สากล อ่านง่าย ไม่ขึ้น None
                    row_preview_list = []
                    
                    for item in updates:
                        r_idx = item.get("row_index")
                        data_to_fill = item.get("data_to_fill", {})
                        
                        # ดึงชื่อเดิมจากตาราง ถ้าไม่มีให้ดึงค่าที่ได้มาจาก AI แทนเพื่อไม่ให้ขึ้น None ค้างบนหน้าจอ
                        emp_name_col = header_map.get("รายชื่อ") or header_map.get("ชื่อพนักงาน") or header_map.get("ชื่อ") or 1
                        existing_name = ws_target.cell(row=r_idx, column=emp_name_col).value
                        
                        display_name = existing_name if existing_name else data_to_fill.get("รายชื่อ") or data_to_fill.get("ชื่อพนักงาน") or data_to_fill.get("ชื่อ") or "พนักงานรายใหม่"
                        
                        # สร้างพิกัดตารางแนวนอนและอัปเดตลง Excel จริง
                        preview_row = {"แถวตารางที่": r_idx, "ชื่อ-นามสกุล": display_name}
                        
                        for h_name, new_val in data_to_fill.items():
                            if h_name in header_map:
                                ws_target.cell(row=r_idx, column=header_map[h_name]).value = new_val
                                preview_row[h_name] = new_val
                                
                        row_preview_list.append(preview_row)

                    if row_preview_list:
                        st.balloons()
                        st.success(f"🎉 ประมวลผลและนำข้อมูลจัดระเบียบกรอกลงไฟล์ '{template_name}' สำเร็จแล้วค่ะ")
                        
                        # แสดงตารางแบบแนวนอน (Wide Table) สไตล์ Global Dashboard อ่านง่ายขึ้น 1,000%
                        st.markdown("#### 👀 ตารางตรวจสอบความถูกต้องก่อนบันทึก (Data Preview)")
                        df_wide = pd.DataFrame(row_preview_list)
                        
                        # ย้ายคอลัมน์ แถวตาราง และ ชื่อ-นามสกุล ไปไว้ซ้ายสุดเสมอเพื่อความสวยงาม
                        cols = list(df_wide.columns)
                        if "แถวตารางที่" in cols: cols.insert(0, cols.pop(cols.index("แถวตารางที่")))
                        if "ชื่อ-นามสกุล" in cols: cols.insert(1, cols.pop(cols.index("ชื่อ-นามสกุล")))
                        df_wide = df_wide[cols]
                        
                        st.dataframe(
                            df_wide.style.set_properties(**{
                                'background-color': '#FFFFFF', 
                                'color': '#111827',
                                'border-color': '#E5E7EB'
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        out_stream = io.BytesIO()
                        wb_target.save(out_stream)
                        out_stream.seek(0)
                        
                        st.write("")
                        st.download_button(
                            label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel ผลลัพธ์สมบูรณ์",
                            data=out_stream,
                            file_name=f"HR_SUCCESS_{template_name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ ไม่สามารถบันทึกข้อมูลได้ เนื่องจากโครงสร้างข้อมูลไม่แมตช์กับหัวตารางในเทมเพลต")
                        
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลระบบ: {e}")
