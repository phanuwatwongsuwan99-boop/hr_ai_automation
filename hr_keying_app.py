import streamlit as st
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd

# 1. ตั้งค่าหน้าตาของเว็บแอปให้สวยงามน่าใช้งาน
st.set_page_config(page_title="HR AI Automation System", page_icon="🤖", layout="wide")

# เรียกใช้งาน API Key ที่เราซ่อนไว้ในระบบหลังบ้านของ Cloud อย่างปลอดภัย
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error("❌ ไม่พบ API Key ในระบบหลังบ้าน กรุณาตั้งค่าใน Advanced Settings ของ Streamlit Cloud")
    st.stop()

# โฟลเดอร์ที่ต้องใช้ในระบบ
TEMPLATES_FOLDER = "hr_templates"

# 2. ฟังก์ชันเรียนรู้โครงสร้างคอลัมน์และรายชื่อพนักงานจากทุกเทมเพลต
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
            except Exception as e:
                pass
    return structures

# --- หน้าตาเว็บแอปฝั่งพนักงานใช้งาน (UI) ---
st.title("🤖 ระบบ HR AI Automation คีย์ข้อมูลสัมพันธ์อัตโนมัติ")
st.write("ยินดีต้อนรับ! ระบบนี้จะช่วยคุณดึงข้อมูลจากเอกสารดิบ (เช่น ไฟล์ค่าคอมมิชชั่น, ไฟล์โอที) แล้วนำไปจับคู่กรอกลงในตารางเทมเพลต HR ของคุณให้อัตโนมัติโดยอ้างอิงจากรหัสหรือชื่อพนักงาน")
st.write("---")

# ส่วนที่ 1: แนะนำผู้ใช้งานและให้ดาวน์โหลดเทมเพลต
st.subheader("📋 ขั้นตอนการใช้งานสำหรับเจ้าหน้าที่ HR")
col_step1, col_step2, col_step3 = st.columns(3)
with col_step1:
    st.markdown("**1. เตรียมตารางแม่แบบ**\nตรวจสอบให้แน่ใจว่าตารางในระบบมีรายชื่อพนักงานแล้ว หรือกดดาวน์โหลดแม่แบบด้านล่างไปเตรียมข้อมูล")
with col_step2:
    st.markdown("**2. อัปโหลดไฟล์ข้อมูลดิบ**\nนำไฟล์ข้อมูลดิบที่ได้มาจากแผนกอื่น (เช่น ไฟล์สรุปยอดค่านายหน้าประจำเดือน) มาวางในระบบ")
with col_step3:
    st.markdown("**3. ตรวจสอบและดาวน์โหลด**\nระบบ AI จะคำนวณและกรอกข้อมูลให้ตรงแถวของพนักงานคนนั้น ๆ คุณสามารถตรวจทานและโหลดไฟล์ได้ทันที")

# สร้างปุ่มให้กดดาวน์โหลดเทมเพลตต้นแบบจากหน้าเว็บได้เลย
templates_available = learn_templates()
if templates_available:
    selected_t_download = st.selectbox("เลือกไฟล์เทมเพลตที่ต้องการดาวน์โหลดไปใช้งานตั้งต้น:", list(templates_available.keys()))
    t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
    with open(t_download_path, "rb") as f:
        st.download_button(
            label=f"📥 ดาวน์โหลดไฟล์แม่แบบ: {selected_t_download}",
            data=f,
            file_name=selected_t_download,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
st.write("---")

# ส่วนที่ 2: ช่องอัปโหลดไฟล์งานจริง
st.subheader("📤 ส่วนอัปโหลดไฟล์งานเพื่อคีย์ข้อมูล")
uploaded_file = st.file_uploader("ลากไฟล์ข้อมูลดิบ (เช่น ไฟล์ Excel ยอดคอมมิชชั่น หรือไฟล์สรุปโอทีประจำเดือน) มาวางตรงนี้ได้เลยครับ", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info(f"📂 รับไฟล์สำเร็จ: {uploaded_file.name} กำลังเตรียมความพร้อมของระบบ...")
    
    if st.button("🚀 สั่งให้ AI เริ่มคีย์ข้อมูลลงตารางอัตโนมัติ"):
        if not templates_available:
            st.error("❌ ไม่พบไฟล์เทมเพลตต้นแบบในระบบหลังบ้าน (โฟลเดอร์ hr_templates/) กรุณาตรวจสอบไฟล์บน GitHub")
        else:
            with st.spinner("🤖 AI กำลังเปิดอ่านเนื้อหาในไฟล์ คัดแยกประเภท และจับคู่ข้อมูลพนักงานอย่างละเอียด โปรดรอสักครู่..."):
                try:
                    # 1. แกะข้อมูลจากไฟล์ดิบที่อัปโหลดแปลงเป็น Text
                    wb_in = load_workbook(uploaded_file, data_only=True)
                    input_text = ""
                    for sheet_name in wb_in.sheetnames:
                        ws_in = wb_in[sheet_name]
                        input_text += f"\n[หน้าตาราง: {sheet_name}]\n"
                        for row in ws_in.iter_rows(values_only=True):
                            row_clean = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
                            if row_clean:
                                input_text += " | ".join(row_clean) + "\n"

                    # 2. ป้อนข้อมูลส่งให้สมองกล Gemini วิเคราะห์จับคู่ความหมาย
                    blueprint = json.dumps(templates_available, ensure_ascii=False, indent=2)
                    prompt = f"""
                    คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลพนักงานจากเอกสารดิบที่แนบมา 
                    แล้วนำข้อมูลไปเติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

                    นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
                    {blueprint}

                    คำสั่ง:
                    1. วิเคราะห์ว่าข้อมูลอินพุตนี้ สอดคล้องกับไฟล์เทมเพลตไหนมากที่สุด (ตอบชื่อไฟล์ในช่อง "selected_template")
                    2. ตรวจสอบ 'รหัสพนักงาน' หรือ 'รายชื่อ' เพื่อจับคู่ระหว่างข้อมูลอินพุตกับพนักงานเดิมในตารางให้ตรงคน
                    3. ดึงค่าในเอกสารดิบที่ตรงกับหัวตารางปลายทาง (เช่น ค่าคอมมิชชั่น หรือ เงินได้) นำไปเติมลงในแถวที่มีตัวเลข "__row_index__" ของพนักงานคนนั้นๆ

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
                    
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=f"{prompt}\n\n[ข้อมูลเอกสารดิบ]\n{input_text}")
                    clean_text = response.text.replace("```json", "").replace("```", "").strip()
                    ai_result = json.loads(clean_text)
                    
                    # 3. นำผลลัพธ์จาก AI มากรอกลงตาราง Excel ต้นฉบับ
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
                                ws_target.cell(row=r_idx, column=header_map[h_name]).value = new_value
                                preview_data.append({"แถวที่": r_idx, "พนักงาน": emp_name_str, "หัวข้อที่เติม": h_name, "จำนวนเงิน/ข้อมูล": new_value})

                    if preview_data:
                        st.success(f"🎯 AI วิเคราะห์สำเร็จ! ข้อมูลนี้สอดคล้องกับตาราง '{template_name}' และทำการจับคู่กรอกข้อมูลให้พนักงานเรียบร้อยแล้ว")
                        
                        # แสดงผลลัพธ์เป็นตารางจำลองให้ User ตรวจทานความถูกต้องก่อนดาวน์โหลด
                        st.subheader("👀 ตารางตรวจทานความถูกต้อง (Preview)")
                        st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
                        
                        # เซฟไฟล์ลง Memory
                        out_stream = io.BytesIO()
                        wb_target.save(out_stream)
                        out_stream.seek(0)
                        
                        st.write("---")
                        st.download_button(
                            label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel ที่คีย์ข้อมูลเสร็จแล้ว",
                            data=out_stream,
                            file_name=f"HR_SUCCESS_{template_name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("⚠️ AI อ่านไฟล์ได้ แต่ไม่พบรายชื่อพนักงานหรือรหัสพนักงานที่ตรงกับตารางแม่แบบในระบบเลย กรุณาตรวจสอบความถูกต้องของไฟล์อินพุต")
                        
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลระบบ: {e}")
                    st.info("💡 คำแนะนำ: กรุณากดปุ่มเพื่อลองใหม่อีกครั้ง")
