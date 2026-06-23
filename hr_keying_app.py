import streamlit as st  # นำคำสั่งนี้มาไว้บรรทัดที่ 1 เสมอ!
import os
import json
import io
from google import genai
from openpyxl import load_workbook
import pandas as pd

# ตั้งค่าหน้าตาของเว็บแอป
st.set_page_config(page_title="HR AI Automation System", page_icon="🤖", layout="wide")

# เรียกใช้งาน API Key จากระบบหลังบ้าน (Secrets)
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error("❌ ไม่พบ API Key ในระบบหลังบ้าน กรุณาตั้งค่าใน Advanced Settings ของ Streamlit Cloud")
    st.stop()

# โฟลเดอร์ที่ต้องใช้ในระบบ
def learn_templates():
    structures = {}
    if not os.path.exists(TEMPLATES_FOLDER):
        return structures

# ... โค้ดส่วนที่เหลือด้านล่างปล่อยไว้เหมือนเดิมได้เลยครับ ...
        
    print("🤖 [ระบบเรียนรู้] กำลังอ่านโครงสร้างคอลัมน์และรายชื่อพนักงาน...")
    
    for filename in os.listdir(TEMPLATES_FOLDER):
        if filename.endswith((".xlsx", ".xls")) and not filename.startswith((".", "~$")):
            try:
                # ด้านใน try ต้องเยื้องเข้าไปอีก 4 ช่อง (รวมเป็น 16 ช่องจากซ้ายสุด)
                wb = load_workbook(os.path.join(TEMPLATES_FOLDER, filename), data_only=True)
                ws = wb.active
            
            # ตรวจสอบก่อนว่าเคยรันไฟล์นี้ไปแล้วหรือยังในโฟลเดอร์ Output 
            # ถ้าเคยรันแล้ว จะหยิบไฟล์ล่าสุดจาก Output มาใช้เรียนรู้ต่อ เพื่อป้องกันการเขียนทับกันมั่ว
            output_file_name = f"HR_UPDATED_{filename}"
            output_path = os.path.join(output_folder, output_file_name)
            template_path = os.path.join(templates_folder, filename)
            
            active_path = output_path if os.path.exists(output_path) else template_path
            
            try:
                wb = load_workbook(active_path, data_only=True)
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
                    template_structures[filename] = {
                        "headers": headers,
                        "existing_data": existing_rows,
                        "source_used": "Output (งานเก่าทำต่อ)" if os.path.exists(output_path) else "Template ต้นแบบ"
                    }
                    print(f"  ✅ เรียนรู้: '{filename}' [{template_structures[filename]['source_used']}] (พบ {len(headers)} คอลัมน์, มีข้อมูลเดิม {len(existing_rows)} แถว)")
            except Exception as e:
                print(f"  ❌ ไม่สามารถอ่านตาราง {filename} ได้: {e}")
    return template_structures

# 3. ฟังก์ชันแกะเนื้อหาจากไฟล์ Input ทุกประเภท
def extract_content_from_input(file_path):
    file_ext = file_path.split(".")[-1].lower()
    if file_ext in ["jpg", "jpeg", "png"]:
        try: return Image.open(file_path)
        except: return None
    elif file_ext == "pdf":
        try:
            pdf_reader = PdfReader(file_path)
            pdf_text = ""
            for page in pdf_reader.pages: pdf_text += page.extract_text() or ""
            return f"[เนื้อหาจากไฟล์ PDF]\n{pdf_text}"
        except: return None
    elif file_ext in ["xlsx", "xls"]:
        try:
            wb = load_workbook(file_path, data_only=True)
            excel_text = ""
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                excel_text += f"\n--- หน้า (Sheet): {sheet_name} ---\n"
                for row in ws.iter_rows(values_only=True):
                    row_clean = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
                    if row_clean:
                        excel_text += " | ".join(row_clean) + "\n"
            return f"[เนื้อหาจากไฟล์ Excel อินพุต]\n{excel_text}"
        except: return None
    return None

# 4. สมองกล AI: วิเคราะห์ไฟล์คัดแยก จับคู่ความหมาย
def ai_analyze_and_match_values(input_data, templates_info, input_filename):
    templates_blueprint = json.dumps(templates_info, ensure_ascii=False, indent=2)
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญด้านระบบ HR Automation หน้าที่ของคุณคือวิเคราะห์ข้อมูลพนักงานจากเอกสารดิบที่แนบมา 
    แล้วนำข้อมูลไปเติมลงในตารางพนักงานของแต่ละเทมเพลตในจุดที่ข้อมูลยังว่างอยู่ให้ถูกต้องและสมบูรณ์

    นี่คือโครงสร้างตารางและรายชื่อพนักงานปัจจุบันที่มีอยู่ในระบบแยกตามไฟล์เทมเพลตต่างๆ:
    {templates_blueprint}

    คำสั่งสำหรับการประมวลผลไฟล์อินพุตชื่อ "{input_filename}":
    1. วิเคราะห์ว่าเนื้อหาข้อมูลในไฟล์อินพุตนี้ สอดคล้องหรือควรนำไปกรอกลงในไฟล์เทมเพลตไหนมากที่สุด (ตอบชื่อไฟล์เทมเพลตในช่อง "selected_template" โดยเอาชื่อมาจาก key หลักของข้อมูลด้านบน)
    2. ทำการตรวจสอบความสัมพันธ์ (Key Matching) เช่น ตรวจดู 'รหัสพนักงาน', 'รายชื่อ' หรือตัวระบุตัวตนอื่นๆ ระหว่างข้อมูลในไฟล์อินพุต กับข้อมูลเดิมที่มีอยู่ในตารางเทมเพลต
    3. เมื่อจับคู่ตัวบุคคลได้ตรงกันแล้ว ให้ดึงข้อมูลค่าต่างๆ ในเอกสารอินพุตที่ตรงกับหัวตารางปลายทาง (พิจารณาจากความหมายของคอลัมน์ แม้คำจะไม่ตรงกันเป๊ะ) นำไปเติมลงในแถวที่มีตัวเลข "__row_index__" ของพนักงานคนนั้นๆ

    ให้ตอบกลับเป็นรูปแบบ JSON ตามโครงสร้างนี้เท่านั้น ห้ามมีคำอธิบายอื่นเด็ดขาด:
    {{
      "selected_template": "ชื่อไฟล์เทมเพลตปลายทางที่เลือก.xlsx",
      "updates": [
         {{
           "row_index": 2, 
           "data_to_fill": {{
              "ชื่อคอลัมน์ที่ต้องการเติมข้อมูลช่องที่ว่าง": "ค่าหรือตัวเลขที่ดึงได้"
           }}
         }}
      ]
    }}
    """
    
    for attempt in range(1, 4):
        try:
            if isinstance(input_data, Image.Image):
                response = client.models.generate_content(model='gemini-2.5-flash', contents=[prompt, input_data])
            else:
                response = client.models.generate_content(model='gemini-2.5-flash', contents=f"{prompt}\n\n[ข้อมูลเอกสารดิบ]\n{input_data}")
                
            clean_text = response.text.replace("```json", "").replace("```", "").strip()
            return json.loads(clean_text)
            
        except Exception as e:
            print(f"  ⚠️ [คำเตือน] รอบที่ {attempt} ระบบ Google หนาแน่นชั่วคราว (Error: {e})")
            if attempt < 3:
                print("  ⏳ กำลังรอระบบว่าง 5 วินาทีแล้วลองส่งใหม่...")
                time.sleep(5)
            else:
                return None

# 5. ฟังก์ชันอัปเดตข้อมูลลง Excel ปลายทางอย่างปลอดภัย
def update_existing_excel(ai_result, templates_folder="hr_templates", output_folder="output_files"):
    template_name = ai_result.get("selected_template")
    updates = ai_result.get("updates", [])
    
    template_path = os.path.join(templates_folder, template_name)
    output_path = os.path.join(output_folder, f"HR_UPDATED_{template_name}")
    
    # หากเคยสร้างไฟล์ผลลัพธ์ไว้แล้วใน output ให้ดึงไฟล์งานเก่ามาเขียนต่อ ห้ามแตะต้อง Template ต้นฉบับ
    if os.path.exists(output_path):
        wb = load_workbook(output_path)
    else:
        wb = load_workbook(template_path)
        
    ws = wb.active
    
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        h_name = str(ws.cell(row=1, column=col_idx).value).strip()
        header_map[h_name] = col_idx
        
    filled_cells = 0
    for item in updates:
        row_idx = item.get("row_index")
        data_to_fill = item.get("data_to_fill", {})
        
        for header_name, new_value in data_to_fill.items():
            if header_name in header_map:
                col_idx = header_map[header_name]
                ws.cell(row=row_idx, column=col_idx).value = new_value
                filled_cells += 1
                
    wb.save(output_path)
    print(f"  ✨ [อัปเดตสำเร็จ] เติมข้อมูลลงช่องว่างได้ทั้งหมด {filled_cells} ช่อง -> ไฟล์ผลลัพธ์: {output_path}\n")

# --- ส่วนควบคุมหลัก ---
if __name__ == "__main__":
    for folder in ["hr_templates", "input_documents", "output_files"]:
        if not os.path.exists(folder): os.makedirs(folder)

    templates_info = learn_all_templates()
    
    if not templates_info:
        print("\n❌ ไม่พบไฟล์เทมเพลต Excel ในโฟลเดอร์ hr_templates")
    else:
        input_folder = "input_documents"
        all_inputs = [f for f in os.listdir(input_folder) if not f.startswith(('.', '~$'))]
        
        print("\n🔍 [ระบบกำลังตรวจสอบโฟลเดอร์ input_documents]")
        print(f"  📦 พบไฟล์อินพุตที่รอการประมวลผล: {all_inputs}")
        print("="*60)

        if not all_inputs:
            print("\n❌ ไม่มีไฟล์ข้อมูลในโฟลเดอร์ input_documents")
        else:
            for idx, filename in enumerate(all_inputs, 1):
                file_path = os.path.join(input_folder, filename)
                print(f"[{idx}/{len(all_inputs)}] 🔍 กำลังประมวลผลดึงค่าจากไฟล์: {filename}")
                
                input_data = extract_content_from_input(file_path)
                if input_data:
                    # อัปเดตข้อมูลแม่แบบตารางแบบ Real-time ก่อนส่งให้ AI เผื่อกรณีลูปทำไฟล์แรกเสร็จแล้วไฟล์ที่สองจะจำได้
                    templates_info = learn_all_templates()
                    
                    result = ai_analyze_and_match_values(input_data, templates_info, filename)
                    if result and result.get("selected_template") in templates_info:
                        print(f"  🎯 AI เลือกกรอกลงไฟล์เทมเพลต: '{result.get('selected_template')}'")
                        update_existing_excel(result)
                    else:
                        print(f"  ⚠️ AI ไม่สามารถจับคู่ข้อมูลกับเทมเพลตเดิมในระบบได้")
                else:
                    print(f"  ⚠️ ไม่รองรับรูปแบบไฟล์ {filename}")
            print("="*60)
            print("🎉 ทำงานคัดแยกและเติมเต็มข้อมูลเสร็จสิ้นเรียบร้อยแล้วครับ!")
