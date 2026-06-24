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
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===================== CSS THEME: Indigo & Slate (Premium HR / Enterprise) =====================
# ปรับโทนสีจาก แดง-ดำ เดิม เป็น Indigo-Slate ซึ่งเป็นโทนมาตรฐานของระบบ HR/Enterprise
# ให้ความรู้สึกน่าเชื่อถือ สุภาพ อ่านง่าย ไม่ดูเหมือนข้อความเตือน/อันตราย
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F8FAFC;
    }

    /* ---------- Header ---------- */
    .main-header {
        background: linear-gradient(135deg, #1E293B 0%, #4338CA 100%);
        padding: 36px 40px;
        border-radius: 16px;
        color: white;
        margin-bottom: 28px;
        box-shadow: 0 8px 24px rgba(67, 56, 202, 0.18);
        position: relative;
        overflow: hidden;
    }
    .main-header::after {
        content: "";
        position: absolute;
        top: -40%;
        right: -10%;
        width: 260px;
        height: 260px;
        background: rgba(255,255,255,0.06);
        border-radius: 50%;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    .main-header p {
        margin: 8px 0 0 0;
        opacity: 0.88;
        font-size: 1rem;
        font-weight: 300;
    }

    /* ---------- Step Progress Pills ---------- */
    .step-track {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 28px;
        flex-wrap: wrap;
    }
    .step-pill {
        display: flex;
        align-items: center;
        gap: 8px;
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 999px;
        padding: 8px 18px 8px 10px;
        font-size: 0.88rem;
        font-weight: 600;
        color: #64748B;
    }
    .step-pill.active {
        background: #EEF2FF;
        border-color: #C7D2FE;
        color: #3730A3;
    }
    .step-pill.done {
        background: #F0FDF4;
        border-color: #BBF7D0;
        color: #166534;
    }
    .step-num {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 22px;
        height: 22px;
        border-radius: 50%;
        background: #CBD5E1;
        color: white;
        font-size: 0.75rem;
        font-weight: 700;
    }
    .step-pill.active .step-num { background: #4F46E5; }
    .step-pill.done .step-num { background: #16A34A; }
    .step-connector {
        width: 24px;
        height: 2px;
        background: #E2E8F0;
    }

    /* ---------- Card ---------- */
    .card-box {
        background-color: #FFFFFF;
        padding: 26px 28px;
        border-radius: 14px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.04);
        margin-bottom: 22px;
    }
    .card-box-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 18px;
    }
    .card-box-icon {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 38px;
        height: 38px;
        border-radius: 10px;
        background: #EEF2FF;
        font-size: 1.1rem;
        flex-shrink: 0;
    }
    .card-box-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #1E293B;
        margin: 0;
    }
    .card-box-subtitle {
        font-size: 0.85rem;
        color: #94A3B8;
        margin: 2px 0 0 0;
    }

    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 16px;
        padding-left: 12px;
        border-left: 4px solid #4F46E5;
    }

    /* ---------- Badges (หมวดหมู่เอกสาร) ---------- */
    .badge-personal, .badge-commission, .badge-ot, .badge-generic {
        padding: 10px 20px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.95rem;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    .badge-personal {
        background-color: #FEF2F2;
        color: #991B1B;
        border: 1px solid #FEE2E2;
    }
    .badge-commission {
        background-color: #F0FDF4;
        color: #166534;
        border: 1px solid #DCFCE7;
    }
    .badge-ot {
        background-color: #FFFBEB;
        color: #92400E;
        border: 1px solid #FEF3C7;
    }
    .badge-generic {
        background-color: #EEF2FF;
        color: #3730A3;
        border: 1px solid #E0E7FF;
    }

    /* ---------- Buttons ---------- */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #4338CA 0%, #4F46E5 100%) !important;
        color: white !important;
        border-radius: 10px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        padding: 14px 30px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 4px 14px rgba(79, 70, 229, 0.25) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(135deg, #3730A3 0%, #4338CA 100%) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 18px rgba(79, 70, 229, 0.32) !important;
    }

    /* ---------- Download buttons ---------- */
    div.stDownloadButton > button:first-child {
        background: #FFFFFF !important;
        color: #4338CA !important;
        border: 1.5px solid #C7D2FE !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 12px 24px !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background: #EEF2FF !important;
        border-color: #4F46E5 !important;
    }

    /* ---------- Misc ---------- */
    .helper-text {
        font-size: 0.85rem;
        color: #94A3B8;
        margin-top: 6px;
    }
    .divider-line {
        border-top: 1px dashed #E2E8F0;
        margin: 22px 0;
    }
    .empty-state {
        text-align: center;
        padding: 50px 20px;
        color: #94A3B8;
    }
    .empty-state-icon {
        font-size: 2.8rem;
        margin-bottom: 12px;
        opacity: 0.5;
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

# --- ส่วนการแสดงผล (Indigo & Slate Premium UI) ---

st.markdown("""
    <div class="main-header">
        <h1>🗂️ HR Smart Matrix Hub</h1>
        <p>อัปโหลดหลักฐาน ให้ AI คัดแยกหมวดหมู่ และจับคู่กรอกข้อมูลพนักงานลงเทมเพลตให้อัตโนมัติ — ไม่ต้องนั่งกรอกเอง</p>
    </div>
""", unsafe_allow_html=True)

templates_available = learn_templates()

# ใช้สถานะไฟล์ที่อัปโหลดเพื่อกำหนดว่าผู้ใช้อยู่ขั้นตอนไหนของ progress bar (ใช้แสดงผลอย่างเดียว ไม่กระทบ logic)
_uploaded_marker = st.session_state.get("_last_uploaded_name")
step2_state = "done" if _uploaded_marker else "active"
step3_state = "active" if _uploaded_marker else ""

st.markdown(f"""
    <div class="step-track">
        <div class="step-pill done"><span class="step-num">1</span>เลือกเทมเพลตปลายทาง</div>
        <div class="step-connector"></div>
        <div class="step-pill {step2_state}"><span class="step-num">2</span>อัปโหลดไฟล์หลักฐาน</div>
        <div class="step-connector"></div>
        <div class="step-pill {step3_state}"><span class="step-num">3</span>ให้ AI วิเคราะห์และดาวน์โหลดผลลัพธ์</div>
    </div>
""", unsafe_allow_html=True)

# ใช้ Layout แบบ step-by-step แนวตั้ง (เต็มความกว้าง) แทนแบบ 2 คอลัมน์เดิม
# เพื่อให้ผู้ใช้ไล่ทำตามขั้นตอนทีละจุดโดยไม่หลงทาง

# ===================== STEP 1: เลือกเทมเพลตปลายทาง =====================
st.markdown('<div class="card-box">', unsafe_allow_html=True)
st.markdown("""
    <div class="card-box-header">
        <div class="card-box-icon">📁</div>
        <div>
            <p class="card-box-title">ขั้นตอนที่ 1 · เลือกตารางแม่แบบปลายทาง</p>
            <p class="card-box-subtitle">เลือกเทมเพลต Excel ที่ต้องการให้ระบบกรอกข้อมูลลงไป</p>
        </div>
    </div>
""", unsafe_allow_html=True)

if templates_available:
    col_select, col_download = st.columns([2, 1], gap="medium")
    with col_select:
        selected_t_download = st.selectbox("เลือกตารางแม่แบบปลายทาง:", list(templates_available.keys()), label_visibility="collapsed")
        st.markdown(f"<p class='helper-text'>📌 คอลัมน์ในเทมเพลตนี้: {', '.join(templates_available[selected_t_download]['headers'])}</p>", unsafe_allow_html=True)

    t_download_path = os.path.join(TEMPLATES_FOLDER, selected_t_download)
    with open(t_download_path, "rb") as f:
        with col_download:
            st.download_button(
                label="📥 ดาวน์โหลดเทมเพลตเปล่า",
                data=f,
                file_name=selected_t_download,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:
    st.error("ไม่พบไฟล์แม่แบบในระบบหลังบ้าน")
st.markdown('</div>', unsafe_allow_html=True)

# ===================== STEP 2: นำเข้าไฟล์หลักฐาน =====================
st.markdown('<div class="card-box">', unsafe_allow_html=True)
st.markdown("""
    <div class="card-box-header">
        <div class="card-box-icon">📤</div>
        <div>
            <p class="card-box-title">ขั้นตอนที่ 2 · นำเข้าไฟล์หลักฐานดิบ</p>
            <p class="card-box-subtitle">รองรับไฟล์ Excel (.xlsx, .xls) หรือรูปภาพเอกสาร (.png, .jpg, .jpeg)</p>
        </div>
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "ลากและวางไฟล์หรือรูปภาพตรงนี้ หรือคลิกเพื่อเลือกไฟล์",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
    label_visibility="visible"
)

if uploaded_file is not None:
    st.session_state["_last_uploaded_name"] = uploaded_file.name
    file_ext = uploaded_file.name.split(".")[-1].lower()
    is_image = file_ext in ["png", "jpg", "jpeg"]
    if is_image:
        col_img, col_spacer = st.columns([1, 1], gap="medium")
        with col_img:
            image_obj = Image.open(uploaded_file)
            st.image(image_obj, caption="รูปหลักฐานที่นำเข้าสู่ระบบ", use_container_width=True)
    else:
        st.markdown(f"<p class='helper-text'>✅ พร้อมวิเคราะห์ไฟล์: <b>{uploaded_file.name}</b></p>", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ===================== STEP 3: วิเคราะห์ด้วย AI และแสดงผล =====================
st.markdown('<div class="card-box">', unsafe_allow_html=True)
st.markdown("""
    <div class="card-box-header">
        <div class="card-box-icon">🤖</div>
        <div>
            <p class="card-box-title">ขั้นตอนที่ 3 · วิเคราะห์ด้วย AI และตรวจสอบผลลัพธ์</p>
            <p class="card-box-subtitle">ระบบจะคัดแยกหมวดหมู่เอกสาร จับคู่ข้อมูล และกรอกลงเทมเพลตให้อัตโนมัติ</p>
        </div>
    </div>
""", unsafe_allow_html=True)

if uploaded_file is None:
    st.markdown("""
        <div class="empty-state">
            <div class="empty-state-icon">🗒️</div>
            <p style="margin:0; font-weight:600; color:#64748B;">ยังไม่มีไฟล์ให้วิเคราะห์</p>
            <p style="margin:4px 0 0 0; font-size:0.9rem;">กรุณาอัปโหลดไฟล์หลักฐานในขั้นตอนที่ 2 ก่อน ระบบจะเปิดให้กดวิเคราะห์ที่นี่</p>
        </div>
    """, unsafe_allow_html=True)
else:
    if st.button("🚀 สั่งเริ่มวิเคราะห์ข้อมูลด้วย AI", use_container_width=True):
        with st.spinner("🤖 AI กำลังอ่านข้อมูลและประมวลผล กรุณารอสักครู่..."):
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

                # 🎯 กล่อง Badge หมวดหมู่
                st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)
                if category == "PERSONAL_DATA":
                    st.markdown('<div class="badge-personal">🆔 คัดแยกสำเร็จ: หมวดหมู่ข้อมูลส่วนตัว / บัตรประชาชนพนักงาน</div>', unsafe_allow_html=True)
                elif category == "COMMISSION":
                    st.markdown('<div class="badge-commission">💰 คัดแยกสำเร็จ: หมวดหมู่รายได้สัมพันธ์ / ค่าคอมมิชชั่นพนักงาน</div>', unsafe_allow_html=True)
                elif category == "OVERTIME":
                    st.markdown('<div class="badge-ot">⏱️ คัดแยกสำเร็จ: หมวดหมู่ข้อมูลเวลาทำงาน / โอทีพนักงาน (OT)</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="badge-generic">📦 คัดแยกสำเร็จ: หมวดหมู่ทั่วไปในระบบ HR</div>', unsafe_allow_html=True)
                st.write("")

                template_path = os.path.join(TEMPLATES_FOLDER, template_name)
                wb_target = load_workbook(template_path)
                ws_target = wb_target.active

                header_map = {str(ws_target.cell(row=1, column=c).value).strip(): c for c in range(1, ws_target.max_column + 1)}

                # 🎯 Logic ข้อมูลตารางแนวนอน อ่านง่าย ไม่ขึ้น None
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
                    st.success(f"🎉 ประมวลผลและนำข้อมูลจัดระเบียบกรอกลงไฟล์ '{template_name}' สำเร็จแล้วค่ะ")

                    st.markdown("#### 👀 ตารางตรวจสอบความถูกต้องก่อนบันทึก (Data Preview)")
                    df_wide = pd.DataFrame(row_preview_list)

                    cols = list(df_wide.columns)
                    if "แถวตารางที่" in cols: cols.insert(0, cols.pop(cols.index("แถวตารางที่")))
                    if "ชื่อ-นามสกุล" in cols: cols.insert(1, cols.pop(cols.index("ชื่อ-นามสกุล")))
                    df_wide = df_wide[cols]

                    st.dataframe(
                        df_wide.style.set_properties(**{
                            'background-color': '#FFFFFF',
                            'color': '#1E293B',
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
                        label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์สมบูรณ์",
                        data=out_stream,
                        file_name=f"HR_SUCCESS_{template_name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ ไม่สามารถบันทึกข้อมูลได้ เนื่องจากโครงสร้างข้อมูลไม่แมตช์กับหัวตารางในเทมเพลต")

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลระบบ: {e}")
st.markdown('</div>', unsafe_allow_html=True)
