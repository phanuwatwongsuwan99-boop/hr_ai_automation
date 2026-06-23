import streamlit as st
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd
from PIL import Image

# 1. ตั้งค่าหน้าเว็บให้เต็มจอและเรียบหรู
st.set_page_config(
    page_title="HR Smart Matrix Hub", 
    page_icon="⚡", 
    layout="wide"
)

# คุมธีมสไตล์ Global Corporate: แดงเบอร์กันดี ดำถ่าน และขาวสว่าง
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F9FAFB;
    }
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #7F1D1D 100%);
        padding: 40px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
        border-bottom: 5px solid #DC2626;
        text-align: center;
    }
    .global-container {
        background-color: #FFFFFF;
        padding: 30px;
        border-radius: 12px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        margin-bottom: 25px;
    }
    .badge-personal {
        background-color: #7F1D1D;
        color: #FFFFFF;
        padding: 8px 18px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 1rem;
        display: inline-block;
        border-left: 5px solid #EF4444;
        margin-bottom: 15px;
    }
    .badge-commission {
        background-color: #14532D;
        color: #FFFFFF;
        padding: 8px 18px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 1rem;
        display: inline-block;
        border-left: 5px solid #22C55E;
        margin-bottom: 15px;
    }
    .badge-ot {
        background-color: #78350F;
        color: #FFFFFF;
        padding: 8px 18px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 1rem;
        display: inline-block;
        border-left: 5px solid #F59E0B;
        margin-bottom: 15px;
    }
    .badge-generic {
        background-color: #1F2937;
        color: #FFFFFF;
        padding: 8px 18px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 1rem;
        display: inline-block;
        border-left: 5px solid #9CA3AF;
        margin-bottom: 15px;
    }
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #111827 0%, #991B1B 100%) !important;
        color: white !important;
        border-radius: 6px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        padding: 14px 40px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(153, 27, 27, 0.2) !important;
        transition: all 0.2s ease !important;
        display: block;
        margin: 0 auto !important;
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

# --- บล็อกส่วนหัวแอปพลิเคชัน (แบนเนอร์กลางหน้าจอ) ---
st.markdown("""
    <div class="main-header">
        <h1 style='margin:0; font-size:2.4rem; font-weight:700; letter-spacing: -0.5px;'>⚡ HR Smart Matrix Hub</h1>
        <p style='margin:8px 0 0 0; opacity:0.85; font-size:1.05rem; font-weight: 300;'>ระบบประมวลผล คัดแยกหมวดหมู่ และลงบันทึกเอกสารพนักงานอัตโนมัติด้วยพลัง AI</p>
    </div>
""", unsafe_allow_html=True)

templates_available = learn_templates()

# 📥 โซนที่ 1: ส่วนนำเข้าข้อมูล (เต็มหน้าจอ ไม่มีช่องโล่งว่างฝั่งขวา)
st.markdown('<div class="global-container">', unsafe_allow_html=True)
st.subheader("📤 นำเข้าไฟล์หลักฐานดิบเพื่อเริ่มงาน")
uploaded_file = st.file_uploader(
    "ลากและวางเอกสารที่นี่ รองรับทั้งไฟล์ Excel (.xlsx, .xls) และรูปภาพหลักฐานจากไลน์ (.png, .jpg, .jpeg)", 
    type=["xlsx", "xls", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    is_image = file_ext in ["png", "jpg", "jpeg"]
    
    if is_image:
        # ย่อขนาดรูปภาพแสดงพรีวิวตรงกลางให้สมดุล
        image_obj = Image.open(uploaded_file)
        col_img_l, col_img_c, col_img_r = st.columns([1, 1.2, 1])
        with col_img_c:
            st.image(image_obj, caption="🔍 รูปภาพหลักฐานที่ส่งเข้าสู่ระบบ", use_container_width=True)
    
    st.write("")
    # ปุ่มประมวลผลใหญ่อยู่ตรงกลางหน้าจอ
    execute_click = st.button("🚀 สั่งเริ่มวิเคราะห์และจับคู่ข้อมูลอัตโนมัติ")
st.markdown('</div>', unsafe_allow_html=True)

# 📊 โซนที่ 2: ส่วนประมวลผลและการแสดงเอาต์พุต (จะโผล่ขึ้นมาเต็มหน้าจอเมื่อกดรันเท่านั้น)
if uploaded_file is not None and 'execute_click' in locals() and execute_click:
    st.markdown('<div class="global-container">', unsafe_allow_html=True)
    st.subheader("📑 ผลการวิเคราะห์และตรวจสอบข้อมูล")
    
    with st.spinner("🤖 AI กำลังสแกนข้อมูล คัดแยกหมวดหมู่ และลงบันทึกในตารางพนักงาน..."):
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
            คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลจากเอกสารดิบหรือรูปภาพที่แนบมา
            แล้วนำข้อมูลไปจับคู่เติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

            นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
            {blueprint}

            คำสั่งในการวิเคราะห์และจัดหมวดหมู่:
            1. ระบุประเภทหมวดหมู่ (Category) ของหลักฐานนี้ โดยเลือกค่าต่อไปนี้: "PERSONAL_DATA", "COMMISSION", "OVERTIME", "GENERIC_HR"
            2. คัดเลือกเทมเพลตปลายทางที่เหมาะสมที่สุดกรอกลงช่อง "selected_template"
            3. ตรวจสอบรหัสหรือชื่อเพื่อทำการจับคู่ หากในตารางปลายทางช่องรายชื่อยังไม่มีค่า (เป็นค่าว่างหรือพนักงานใหม่) ให้ดึงข้อมูล "ชื่อ-นามสกุล" ที่สแกนได้จากหลักฐานมาเติมใส่ด้วย เพื่อไม่ให้เกิดค่า None

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
            clean_text = response.text.replace("```json", "").replace("
```", "").strip()
            ai_result = json.loads(clean_text)
            
            category = ai_result.get("category", "GENERIC_HR")
            template_name = ai_result.get("selected_template")
            updates = ai_result.get("updates", [])
            
            # บล็อกหมวดหมู่แสดงผลเต็มความกว้าง
            if category == "PERSONAL_DATA":
                st.markdown('<div class="badge-personal">🆔 คัดแยกเอกสารสำเร็จ: หมวดหมู่ข้อมูลส่วนตัว / บัตรประชาชนพนักงาน</div>', unsafe_allow_html=True)
            elif category == "COMMISSION":
                st.markdown('<div class="badge-commission">💰 คัดแยกเอกสารสำเร็จ: หมวดหมู่รายได้สัมพันธ์ / ค่าคอมมิชชั่นพนักงาน</div>', unsafe_allow_html=True)
            elif category == "OVERTIME":
                st.markdown('<div class="badge-ot">⏱️ คัดแยกเอกสารสำเร็จ: หมวดหมู่ข้อมูลเวลาทำงาน / โอทีพนักงาน (OT)</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="badge-generic">📦 คัดแยกเอกสารสำเร็จ: หมวดหมู่ทั่วไปในระบบงาน HR</div>', unsafe_allow_html=True)
                
            template_path = os.path.join(TEMPLATES_FOLDER, template_name)
            wb_target = load_workbook(template_path)
            ws_target = wb_target.active
            
            header_map = {str(ws_target.cell(row=1, column=c).value).strip(): c for c in range(1, ws_target.max_column + 1)}
            
            row_preview_list = []
            for item in updates:
                r_idx = item.get("row_index")
                data_to_fill = item.get("data_to_fill", {})
                
                emp_name_col = header_map.get("รายชื่อ") or header_map.get("ชื่อพนักงาน") or header_map.get("ชื่อ") or 1
                existing_name = ws_target.cell(row=r_idx, column=emp_name_col).value
                
                display_name = existing_name if existing_name else data_to_fill.get("รายชื่อ") or data_to_fill.get("ชื่อพนักงาน") or data_to_fill.get("ชื่อ") or "พนักงานรายใหม่"
                
                preview_row = {"แถวตารางที่": r_idx, "ชื่อ-นามสกุล": display_name}
                for h_name, new_val in data_to_fill.items():
                    if h_name in header_map:
                        ws_target.cell(row=r_idx, column=header_map[h_name]).value = new_val
                        preview_row[h_name] = new_val
                        
                row_preview_list.append(preview_row)

            if row_preview_list:
                st.balloons()
                st.success(f"🤖 AI นำข้อมูลบันทึกลงไฟล์หลัก '{template_name}' เรียบร้อยแล้ว สามารถตรวจสอบความถูกต้องด้านล่างนี้ได้เลยค่ะ")
                
                # แสดงผลตารางแนวนอน สากล สะอาดตา ขยายเต็มหน้าจอ
                df_wide = pd.DataFrame(row_preview_list)
                cols = list(df_wide.columns)
                if "แถวตารางที่" in cols: cols.insert(0, cols.pop(cols.index("แถวตารางที่")))
                if "ชื่อ-นามสกุล" in cols: cols.insert(1, cols.pop(cols.index("ชื่อ-นามสกุล")))
                df_wide = df_wide[cols]
                
                st.dataframe(df_wide, use_container_width=True, hide_index=True)
                
                out_stream = io.BytesIO()
                wb_target.save(out_stream)
                out_stream.seek(0)
                
                st.write("")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel ที่อัปเดตข้อมูลแล้ว",
                    data=out_stream,
                    file_name=f"HR_UPDATED_{template_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ โครงสร้างข้อมูลในหลักฐานไม่สัมพันธ์กับหัวตารางปลายทาง กรุณาตรวจสอบไฟล์แม่แบบ")
                
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในระบบ: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# 📥 โซนที่ 3: ซ่อนคลังไฟล์แม่แบบปลายทางไว้ในสปอยเลอร์ด้านล่างสุด ไม่ให้รกสายตาหน้าแรก
st.write("")
with st.expander("🛠️ สำหรับผู้ดูแลระบบ: ตรวจสอบและดาวน์โหลดไฟล์ตารางแม่แบบเปล่า"):
    if templates_available:
        selected_t_download = st.selectbox("เลือกตารางแม่แบบปลายทางเพื่อตรวจดูหัวคอลัมน์:", list(templates_available.keys()))
        t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
        with open(t_download_path, "rb") as f:
            st.download_button(
                label="📥 ดาวน์โหลดตารางต้นแบบนี้",
                data=f,
                file_name=selected_t_download,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
