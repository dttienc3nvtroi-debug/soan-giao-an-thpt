import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import time

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án 5512 (Sách KNTT Chuẩn)", 
    layout="wide", 
    page_icon="📝"
)

# Giao diện & Font chữ
st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 1250px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    .sidebar-title { color: #0f172a; font-size: 19px !important; font-weight: 700; margin-bottom: 10px; border-bottom: 3px solid #2563eb; padding-bottom: 4px; }
    .step-header { color: #dc2626 !important; font-size: 21px !important; font-weight: 700 !important; margin-top: 15px !important; margin-bottom: 8px !important; padding-left: 8px; border-left: 5px solid #dc2626; }
    div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #1e293b !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CƠ SỞ DỮ LIỆU ĐẦY ĐỦ VÀ ĐÚNG THỨ TỰ SGK KẾT NỐI TRI THỨC
# ==========================================
DATABASE_KNTT = {
    "Toán học": {
        "Lớp 10": {
            "Chương I. Mệnh đề và Tập hợp": [
                "Bài 1. Mệnh đề",
                "Bài 2. Tập hợp và các phép toán trên tập hợp"
            ],
            "Chương II. Bất phương trình và Hệ bất phương trình bậc nhất hai ẩn": [
                "Bài 3. Bất phương trình bậc nhất hai ẩn",
                "Bài 4. Hệ bất phương trình bậc nhất hai ẩn"
            ],
            "Chương III. Hệ thức lượng trong tam giác": [
                "Bài 5. Giá trị lượng giác của một góc từ 0° đến 180°",
                "Bài 6. Hệ thức lượng trong tam giác"
            ],
            "Chương IV. Vectơ": [
                "Bài 7. Các khái niệm mở đầu",
                "Bài 8. Tổng và hiệu của hai vectơ",
                "Bài 9. Tích của một số với một vectơ",
                "Bài 10. Tích vô hướng của hai vectơ"
            ],
            "Chương V. Các số đặc trưng của mẫu số liệu không ghép nhóm": [
                "Bài 11. Số gần đúng và sai số",
                "Bài 12. Số trung bình và trung vị của mẫu số liệu không ghép nhóm",
                "Bài 13. Tứ phân vị và mốt của mẫu số liệu không ghép nhóm",
                "Bài 14. Các số đặc trưng đo độ phân tán"
            ],
            "Chương VI. Hàm số, đồ thị và ứng dụng": [
                "Bài 15. Hàm số và đồ thị",
                "Bài 16. Hàm số bậc hai",
                "Bài 17. Dấu của tam thức bậc hai",
                "Bài 18. Phương trình quy về phương trình bậc hai"
            ],
            "Chương VII. Phương pháp tọa độ trong mặt phẳng": [
                "Bài 19. Phương trình đường thẳng",
                "Bài 20. Vị trí tương đối giữa hai đường thẳng. Góc và khoảng cách",
                "Bài 21. Đường tròn trong mặt phẳng tọa độ",
                "Bài 22. Ba đường conic trong mặt phẳng tọa độ"
            ],
            "Chương VIII. Đại số tổ hợp": [
                "Bài 23. Quy tắc đếm",
                "Bài 24. Hoán vị, chỉnh hợp và tổ hợp",
                "Bài 25. Nhị thức Newton"
            ],
            "Chương IX. Tính xác suất theo định nghĩa cổ điển": [
                "Bài 26. Biến cố và định nghĩa cổ điển của xác suất",
                "Bài 27. Thực hành tính xác suất theo định nghĩa cổ điển"
            ]
        },
        "Lớp 11": {
            "Chương I. Hàm số lượng giác và Phương trình lượng giác": [
                "Bài 1. Giá trị lượng giác của góc lượng giác",
                "Bài 2. Công thức lượng giác",
                "Bài 3. Hàm số lượng giác",
                "Bài 4. Phương trình lượng giác cơ bản"
            ],
            "Chương II. Dãy số. Cấp số cộng và Cấp số nhân": [
                "Bài 5. Dãy số",
                "Bài 6. Cấp số cộng",
                "Bài 7. Cấp số nhân"
            ],
            "Chương III. Các số đặc trưng đo xu thế trung tâm của mẫu số liệu ghép nhóm": [
                "Bài 8. Mẫu số liệu ghép nhóm",
                "Bài 9. Các số đặc trưng đo xu thế trung tâm"
            ],
            "Chương IV. Quan hệ song song trong không gian": [
                "Bài 10. Đường thẳng và mặt phẳng trong không gian",
                "Bài 11. Hai đường thẳng song song",
                "Bài 12. Đường thẳng song song với mặt phẳng",
                "Bài 13. Hai mặt phẳng song song",
                "Bài 14. Phép chiếu song song"
            ],
            "Chương V. Giới hạn. Hàm số liên tục": [
                "Bài 15. Giới hạn của dãy số",
                "Bài 16. Giới hạn của hàm số",
                "Bài 17. Hàm số liên tục"
            ],
            "Chương VI. Hàm số mũ và Hàm số lôgarit": [
                "Bài 18. Lũy thừa với số mũ thực",
                "Bài 19. Lôgarit",
                "Bài 20. Hàm số mũ và hàm số lôgarit",
                "Bài 21. Phương trình, bất phương trình mũ và lôgarit"
            ],
            "Chương VII. Quan hệ vuông góc trong không gian": [
                "Bài 22. Hai đường thẳng vuông góc",
                "Bài 23. Đường thẳng vuông góc với mặt phẳng",
                "Bài 24. Hai mặt phẳng vuông góc",
                "Bài 25. Khoảng cách trong không gian",
                "Bài 26. Góc giữa đường thẳng và mặt phẳng. Góc giữa hai mặt phẳng"
            ],
            "Chương VIII. Các quy tắc tính xác suất": [
                "Bài 27. Biến cố xung khắc và quy tắc cộng xác suất",
                "Bài 28. Biến cố độc lập và quy tắc nhân xác suất"
            ],
            "Chương IX. Đạo hàm": [
                "Bài 29. Định nghĩa và ý nghĩa của đạo hàm",
                "Bài 30. Các quy tắc tính đạo hàm"
            ]
        },
        "Lớp 12": {
            "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số": [
                "Bài 1. Tính đơn điệu và cực trị của hàm số",
                "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số",
                "Bài 3. Đường tiệm cận của đồ thị hàm số",
                "Bài 4. Khảo sát sự biến thiên và vẽ đồ thị của hàm số",
                "Bài 5. Ứng dụng đạo hàm để giải quyết một số vấn đề thực tiễn"
            ],
            "Chương II. Tọa độ của vectơ trong không gian": [
                "Bài 6. Vectơ trong không gian",
                "Bài 7. Hệ trục tọa độ trong không gian",
                "Bài 8. Biểu thức tọa độ của các phép toán vectơ"
            ],
            "Chương III. Các số đặc trưng đo độ phân tán của mẫu số liệu ghép nhóm": [
                "Bài 9. Khoảng biến thiên và khoảng tứ phân vị của mẫu số liệu ghép nhóm",
                "Bài 10. Phương sai và độ lệch chuẩn của mẫu số liệu ghép nhóm"
            ],
            "Chương IV. Nguyên hàm và Tích phân": [
                "Bài 11. Nguyên hàm",
                "Bài 12. Tích phân",
                "Bài 13. Ứng dụng hình học của tích phân"
            ],
            "Chương V. Phương pháp tọa độ trong không gian": [
                "Bài 14. Phương trình mặt phẳng",
                "Bài 15. Phương trình đường thẳng trong không gian",
                "Bài 16. Công thức tính góc và khoảng cách trong không gian",
                "Bài 17. Phương trình mặt cầu"
            ],
            "Chương VI. Xác suất có điều kiện": [
                "Bài 18. Xác suất có điều kiện",
                "Bài 19. Công thức xác suất toàn phần và công thức Bayes"
            ]
        }
    }
}

# ==========================================
# HÀM KẾT NỐI GEMINI ƯU TIÊN MODEL STABLE & TỰ TÌM MODEL RẢNH
# ==========================================
def get_working_model(api_key):
    genai.configure(api_key=api_key)
    try:
        valid_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                if "-tts" not in name.lower() and "audio" not in name.lower():
                    valid_models.append(name)
        
        # Ưu tiên gemini-1.5-flash vì Quota Free lớn hơn hẳn gemini-2.0-flash
        for target in ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"]:
            for vm in valid_models:
                if target == vm:
                    return vm
        return valid_models[0] if valid_models else "gemini-1.5-flash"
    except Exception:
        return "gemini-1.5-flash"

def generate_content_with_retry(model, contents, max_retries=3):
    """Hàm tự động thử lại khi bị dính lỗi Rate Limit 429"""
    for attempt in range(max_retries):
        try:
            return model.generate_content(contents)
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg or "quota" in err_msg:
                if attempt < max_retries - 1:
                    wait_time = 10 * (attempt + 1)
                    st.warning(f"⚠️ Hệ thống Google đang bận (Lỗi 429 Quota). Tự động thử lại sau {wait_time} giây... (Lần {attempt + 1}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    raise Exception("❌ Đã hết lượt gọi miễn phí trong phút này của API Key. Vui lòng đợi khoảng 1-2 phút rồi bấm lại, hoặc đổi một API Key mới!")
            else:
                raise e

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 CẤU HÌNH HỆ THỐNG</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key...")
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN TRƯỜNG & GV</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 15px; background: #eff6ff; padding: 15px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 26px; font-weight: 800; color: #1e3a8a;">HỆ THỐNG SOẠN GIÁO ÁN 5512 TỰ ĐỘNG CHUẨN KNTT</div>
        <div style="font-size: 16px; font-weight: 700; color: #047857; margin-top: 4px;">📘 BỘ SÁCH: KẾT NỐI TRI THỨC VỚI CUỘC SỐNG</div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 4px;">Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN, LỚP, CHƯƠNG VÀ BÀI
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN, LỚP, CHƯƠNG & BÀI HỌC (ĐÚNG THỨ TỰ SGK)</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    subject_list = list(DATABASE_KNTT.keys()) + ["Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"]
    subject = st.selectbox("Môn học:", subject_list)

with c2:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

with c3:
    duration = st.number_input("Số tiết dạy:", value=2, min_value=1, max_value=10)

available_chapters = DATABASE_KNTT.get(subject, {}).get(grade, {})

col_ch, col_les = st.columns([1, 1])

if available_chapters:
    with col_ch:
        selected_chapter = st.selectbox("Chọn Chương / Chủ đề (Đúng thứ tự SGK):", list(available_chapters.keys()))
    with col_les:
        lessons_in_chap = available_chapters[selected_chapter]
        selected_lesson = st.selectbox("Chọn Bài dạy cụ thể:", lessons_in_chap)
else:
    with col_ch:
        selected_chapter = st.text_input("Tên Chương / Chủ đề:", placeholder="Nhập tên chương...")
    with col_les:
        selected_lesson = st.text_input("Tên Bài dạy cụ thể:", placeholder="Nhập tên bài...")

# ==========================================
# BƯỚC 2: YÊU CẦU CẦN ĐẠT
# ==========================================
st.markdown('<div class="step-header">🎯 BƯỚC 2: YÊU CẦU CẦN ĐẠT (SGV KẾT NỐI TRI THỨC)</div>', unsafe_allow_html=True)

if 'current_lesson_key' not in st.session_state or st.session_state['current_lesson_key'] != selected_lesson:
    st.session_state['current_lesson_key'] = selected_lesson
    st.session_state['auto_ycd'] = ""

col_load_ycd, col_upload = st.columns([1, 1])
with col_load_ycd:
    if st.button("⚡ Tải nhanh Yêu cầu cần đạt chuẩn SGV bài này", type="secondary", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng dán Gemini API Key ở menu bên trái!")
        else:
            try:
                active_model = get_working_model(api_key.strip())
                model = genai.GenerativeModel(active_model)
                
                prompt_ycd = f"""Bạn là Chuyên gia Giáo dục môn {subject}. Hãy trích xuất NGUYÊN VÃN toàn bộ Yêu cầu cần đạt (về kiến thức, năng lực) chuẩn SGV bộ sách Kết nối tri thức với cuộc sống - {grade} cho bài: "{selected_lesson}" (Thuộc {selected_chapter}). Viết ngắn gọn dạng gạch đầu dòng."""
                
                with st.spinner("⚡ AI đang tải YCĐ từ SGV Kết nối tri thức..."):
                    res = generate_content_with_retry(model, [prompt_ycd])
                    st.session_state['auto_ycd'] = res.text
                    st.success("✅ Đã tải xong YCĐ!")
            except Exception as e:
                st.error(f"{str(e)}")

with col_upload:
    uploaded_sgv_file = st.file_uploader("Hoặc Tải file ảnh/PDF trang SGV (Nếu có):", type=["pdf", "png", "jpg", "jpeg"])

ycd_content = st.text_area("Mô tả Yêu cầu cần đạt (Có thể sửa thêm):", value=st.session_state.get('auto_ycd', ''), height=140, placeholder="Nhấn nút 'Tải nhanh Yêu cầu cần đạt' ở trên để AI tự điền...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: YẾU TỐ TÍCH HỢP</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn tích hợp:",
    [
        "Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)", 
        "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện & Tự học"
    ],
    default=["Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)", "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)"]
)

# ==========================================
# XUẤT FILE WORD 5512 CỰC ĐẦY ĐỦ
# ==========================================
def generate_doc(content_text):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    p_left = table.cell(0, 0).paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    p_left.add_run(f"TRƯỜNG: {school_name.upper()}\n").bold = True
    p_left.add_run(f"TỔ: {dept_name.upper()}").bold = True

    p_right = table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    p_right.add_run("Họ và tên giáo viên:\n").bold = True
    p_right.add_run(teacher_name).bold = True

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"KẾ HOẠCH BÀI DẠY: {selected_lesson.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn: {subject} ({grade}) - Sách: Kết nối tri thức với cuộc sống\nThời lượng: {duration} tiết").italic = True

    lines = content_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str or line_str.startswith("---") or line_str.startswith("# "):
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        clean_text = line_str.replace("**", "").replace("*", "")

        if clean_text.startswith(("I. ", "II. ", "III. ", "IV. ")):
            p.paragraph_format.space_before = Pt(10)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(13)
        elif clean_text.startswith(("HOẠT ĐỘNG ", "Tiết ", "1. ", "2. ", "3. ", "4. ")):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith(("a)", "b)", "c)", "d)")):
            p.paragraph_format.left_indent = Inches(0.15)
            p.add_run(clean_text).bold = True
        elif "Bước 1:" in clean_text or "Bước 2:" in clean_text or "Bước 3:" in clean_text or "Bước 4:" in clean_text:
            p.paragraph_format.left_indent = Inches(0.3)
            p.add_run(clean_text).bold = True
        else:
            p.add_run(clean_text)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẤM TẠO GIÁO ÁN WORD CHUẨN 5512 (ĐẦY ĐỦ TỪ A-Z)", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở bên trái!")
    elif not selected_lesson:
        st.error("⚠️ Vui lòng chọn Tên Bài dạy!")
    else:
        try:
            active_model = get_working_model(api_key.strip())
            
            model = genai.GenerativeModel(
                model_name=active_model,
                generation_config=genai.GenerationConfig(
                    temperature=0.2,
                    max_output_tokens=8192
                )
            )
            
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt_5512 = f"""
            Đóng vai Giáo viên Giỏi môn {subject}. Hãy viết Kế hoạch bài dạy CHUẨN CÔNG VĂN 5512 RẤT CHI TIẾT VÀ ĐẦY ĐỦ, KHÔNG ĐƯỢC TẮT/TÓM TẮT BẤT KỲ MỤC NÀO.
            
            THÔNG TIN BÀI DẠY:
            - Môn: {subject} ({grade}) - Bộ sách: Kết nối tri thức với cuộc sống
            - Bài dạy: {selected_lesson} (Thuộc {selected_chapter})
            - Thời lượng: {duration} tiết
            - Yếu tố tích hợp: {integration_str}
            - Yêu cầu cần đạt SGV: {ycd_content}

            CẤU TRÚC BẮT BUỘC 5512 (VIẾT ĐẦY ĐỦ):
            I. MỤC TIÊU
               1. Kiến thức
               2. Năng lực (Năng lực chung + Năng lực đặc thù)
               3. Phẩm chất
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Phải chia đầy đủ cho {duration} tiết dạy):
               - Hoạt động 1: Mở đầu / Khởi động
               - Hoạt động 2: Hình thành kiến thức mới (Chi tiết từng đơn vị kiến thức)
               - Hoạt động 3: Luyện tập (Có bài tập cụ thể và lời giải chi tiết)
               - Hoạt động 4: Vận dụng (Yêu cầu thực tế/bài tập mở rộng)

            MỖI HOẠT ĐỘNG BẮT BUỘC CÓ ĐỦ 4 MỤC:
            a) Mục tiêu
            b) Nội dung (Đưa ra câu hỏi, bài tập cụ thể trong SGK Kết nối tri thức)
            c) Sản phẩm (Lời giải, đáp án chi tiết từng câu hỏi)
            d) Tổ chức thực hiện (Phải viết chi tiết đủ 4 bước: Bước 1 Chuyển giao; Bước 2 Thực hiện; Bước 3 Báo cáo thảo luận; Bước 4 Kết luận, nhận định)
            """

            contents = [prompt_5512]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                contents.append({"mime_type": uploaded_sgv_file.type, "data": bytes_data})

            with st.spinner(f"⚡ Đang dùng model '{active_model}' tạo Giáo án Word 5512 chi tiết..."):
                res = generate_content_with_retry(model, contents)
                doc_file = generate_doc(res.text)
                
                st.success("🎉 Đã tạo xong Giáo án hoàn chỉnh!")
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{selected_lesson.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown(res.text)

        except Exception as e:
            st.error(f"{str(e)}")
