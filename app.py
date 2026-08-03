import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import json
import re
import time

# ==========================================
# CẤU HÌNH TRANG STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn SGV)", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# CẤU HÌNH GIAO DIỆN & FONT CHỮ
# ==========================================
st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1300px;
    }
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Roboto, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    .sidebar-title {
        color: #0f172a;
        font-size: 20px !important;
        font-weight: 700;
        margin-bottom: 12px;
        border-bottom: 3px solid #2563eb;
        padding-bottom: 6px;
    }
    .step-header {
        color: #dc2626 !important;
        font-size: 22px !important;
        font-weight: 700 !important;
        margin-top: 20px !important;
        margin-bottom: 12px !important;
        padding-left: 10px;
        border-left: 5px solid #dc2626;
    }
    div[data-testid="stWidgetLabel"] p, .custom-label {
        font-size: 19px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 4px !important;
    }
    .stSelectbox div[data-baseweb="select"] *,
    .stSelectbox [data-testid="stMarkdownContainer"] p,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="input"] input, .stTextInput input, .stNumberInput input {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="textarea"] textarea, .stTextArea textarea {
        font-size: 17px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
    .stButton button p {
        font-size: 19px !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# HÀM XỬ LÝ TỆP ĐÍNH KÈM
# ==========================================
def process_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    file_type = uploaded_file.type
    bytes_data = uploaded_file.getvalue()
    if "pdf" in file_type:
        mime = "application/pdf"
    elif "png" in file_type:
        mime = "image/png"
    elif "jpeg" in file_type or "jpg" in file_type:
        mime = "image/jpeg"
    else:
        mime = file_type if file_type else "application/octet-stream"
    return {"mime_type": mime, "data": bytes_data}

# ==========================================
# HÀM GỌI GEMINI API THÔNG MINH (CHỐNG LỖI 429 METRIC QUOTA)
# ==========================================
def call_gemini(api_key, preferred_model, contents):
    genai.configure(api_key=api_key)
    
    # Danh sách ưu tiên chuyển đổi khi bị nghẽn API
    fallback_models = [
        preferred_model.replace("models/", "").strip(),
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-2.0-flash"
    ]
    
    # Loại bỏ trùng lặp giữ nguyên thứ tự
    seen = set()
    models_to_try = [m for m in fallback_models if not (m in seen or seen.add(m))]

    generation_config = genai.types.GenerationConfig(
        temperature=0.1,
        top_p=0.8,
        top_k=40
    )

    last_exception = None
    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name=m_name, generation_config=generation_config)
            response = model.generate_content(contents)
            return response
        except Exception as e:
            last_exception = e
            err_msg = str(e)
            # Nếu dính lỗi 429 (Quota Exceeded / Rate Limit)
            if "429" in err_msg or "Quota exceeded" in err_msg or "ResourceHasBeenExhausted" in err_msg:
                time.sleep(2) # Tạm dừng 2s rồi thử model khác
                continue
            else:
                raise e
                
    # Nếu thử toàn bộ danh sách vẫn lỗi 429
    raise Exception(
        "⚠️ Hệ thống AI Google đang bị quá tải lượt gọi Free (Lỗi 429). "
        "Thầy vui lòng chờ khoảng 30 - 45 giây rồi bấm thử lại, "
        "hoặc tạo thêm 1 API Key mới tại https://aistudio.google.com/ app để sử dụng mượt mà hơn."
    ) from last_exception

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP & CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key vào đây...")
    model_name = st.selectbox("Mô hình AI ưu tiên:", ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"], index=0)
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ CHÍNH CỦA ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 20px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 18px; border-radius: 12px; border: 1px solid #bfdbfe;">
        <div style="font-size: 28px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 100% SGV (5512)
        </div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 6px;">
            📝 Tác giả: DƯƠNG TẤN TIẾN — GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)

col_sub, col_grd, col_book = st.columns(3)
with col_sub:
    subject = st.selectbox("Môn học/Hoạt động GD:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"])
with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
with col_book:
    book_series = st.selectbox("Bộ sách Giáo khoa:", ["Kết nối tri thức với cuộc sống", "Cánh diều", "Chân trời sáng tạo"])

# ==========================================
# BƯỚC 2: TRA CỨU TẠO BÀI HỌC VÀ MỤC TIÊU CẦN ĐẠT
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: XÁC ĐỊNH BÀI HỌC VÀ YÊU CẦU CẦN ĐẠT (CHUẨN SGV)</div>', unsafe_allow_html=True)

input_mode = st.radio(
    "Phương thức lấy thông tin bài học:",
    ["📂 Đọc từ Tệp đính kèm (SGV / Cấu trúc chương trình PDF/Ảnh)", "🤖 Tra cứu tự động bằng AI Gemini", "✍️ Nhập thủ công (Khuyên dùng khi cần tên bài chính xác 100%)"],
    horizontal=True
)

chapter_title = ""
lesson_title = ""
requirements = ""
duration = 3

if "📂 Đọc từ Tệp" in input_mode:
    uploaded_file = st.file_uploader("Tải lên File PDF hoặc Ảnh trang SGV/Nội dung bài dạy:", type=["pdf", "png", "jpg", "jpeg"])
    clean_api_key = api_key.strip() if api_key else ""
    
    if uploaded_file and clean_api_key:
        if st.button("⚡ Phân tích Tệp & Trích xuất Tên bài + Mục tiêu chuẩn SGV", type="primary"):
            with st.spinner("🔍 AI đang đọc chính xác nguyên văn SGV từ file..."):
                try:
                    file_part = process_uploaded_file(uploaded_file)
                    prompt_extract = """
                    Hãy đọc chính xác nội dung trong file/ảnh SGV này và trích xuất dữ liệu dưới dạng JSON thuần túy (không dùng markdown code block, không giải thích gì thêm):
                    YÊU CẦU ĐẶC BIỆT: Tên Chương và Tên Bài học PHẢI GHI ĐẦY ĐỦ 100% NGUYÊN VĂN THEO FILE, KHÔNG ĐƯỢC CẮT BỚT HOẶC TÓM TẮT DÙ CHỈ 1 TỪ.
                    {
                        "chapter": "Tên Chương/Chủ đề đầy đủ",
                        "lesson": "Tên bài học nguyên văn đầy đủ",
                        "duration": "Số tiết thực hiện (Số nguyên)",
                        "requirements": "Liệt kê chính xác, đầy đủ từng gạch đầu dòng các Yêu cầu cần đạt nguyên văn theo SGV"
                    }
                    """
                    res = call_gemini(clean_api_key, model_name, [file_part, prompt_extract])
                    raw_text = res.text.strip()
                    
                    match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                    if match:
                        json_str = match.group(0)
                        data = json.loads(json_str)
                        st.session_state['extracted_data'] = data
                        st.success("✅ Trích xuất thành công! Thầy có thể kiểm tra và chỉnh sửa lại bên dưới.")
                    else:
                        st.error("Không thể phân tích dữ liệu dạng JSON từ kết quả.")
                except Exception as e:
                    st.error(f"❌ {e}")

    if 'extracted_data' in st.session_state:
        ext = st.session_state['extracted_data']
        chapter_title = st.text_input("Chương / Chủ đề (Đầy đủ):", value=ext.get('chapter', ''))
        lesson_title = st.text_input("Tên Bài dạy đầy đủ nguyên văn SGK (Thầy có thể sửa lại trực tiếp ở đây):", value=ext.get('lesson', ''))
        try:
            duration = int(ext.get('duration', 3))
        except:
            duration = 3
        duration = st.number_input("Số tiết thực hiện:", value=duration, min_value=1, max_value=20)
        requirements = st.text_area("Yêu cầu cần đạt (Trích xuất từ SGV):", value=ext.get('requirements', ''), height=150)
    else:
        chapter_title = st.text_input("Chương / Chủ đề:", value="")
        lesson_title = st.text_input("Tên Bài dạy đầy đủ (Chuẩn SGV):", value="")
        duration = st.number_input("Số tiết thực hiện:", value=3, min_value=1, max_value=20)
        requirements = st.text_area("Yêu cầu cần đạt (Trích xuất từ SGV):", value="", height=120)

elif "🤖 Tra cứu tự động" in input_mode:
    clean_api_key = api_key.strip() if api_key else ""
    
    if st.button("🔍 Tra cứu ngay Chương trình & Yêu cầu cần đạt SGV"):
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập API Key ở menu bên trái!")
        else:
            with st.spinner("⚡ AI đang tìm kiếm chương trình chính xác..."):
                try:
                    prompt_lookup = f"""
                    Bạn là Chuyên gia SGK. Hãy tra cứu và liệt kê chính xác TẤT CẢ các Bài học và Yêu cầu cần đạt chuẩn SGV môn {subject} {grade} ({book_series}).
                    YÊU CẦU BẮT BUỘC: Tên bài học ("lesson") và Tên chương ("chapter") PHẢI VIẾT NGUYÊN VĂN DÀI ĐẦY ĐỦ 100% THEO CẤU TRÚC SGK BỘ GIÁO DỤC, KHÔNG ĐƯỢC TÓM TẮT.
                    
                    Trả về duy nhất định dạng JSON thuần túy dạng mảng các object (không kèm markdown):
                    [
                        {{
                            "chapter": "Tên Chương đầy đủ...",
                            "lesson": "Tên Bài học nguyên văn đầy đủ...",
                            "duration": 3,
                            "requirements": "Yêu cầu cần đạt chuẩn SGV..."
                        }}
                    ]
                    """
                    res = call_gemini(clean_api_key, model_name, [prompt_lookup])
                    raw_text = res.text.strip()
                    
                    match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    if match:
                        data = json.loads(match.group(0))
                        st.session_state['ai_lessons'] = data
                        st.success(f"✅ Đã tìm thấy {len(data)} bài học chuẩn SGV!")
                    else:
                        st.error("Lỗi cấu trúc dữ liệu trả về.")
                except Exception as e:
                    st.error(f"❌ {e}")
                    
    if 'ai_lessons' in st.session_state:
        lessons_list = st.session_state['ai_lessons']
        options = [f"{item.get('chapter', '')} ➔ {item.get('lesson', '')}" for item in lessons_list]
        selected_idx = st.selectbox("Chọn Bài học từ danh sách SGV:", range(len(options)), format_func=lambda i: options[i])
        
        sel = lessons_list[selected_idx]
        chapter_title = st.text_input("Chương / Chủ đề (Đầy đủ):", value=sel.get('chapter', ''))
        lesson_title = st.text_input("Tên Bài dạy đầy đủ nguyên văn SGK (Thầy có thể bổ sung/sửa lại ở đây):", value=sel.get('lesson', ''))
        try:
            dur_val = int(sel.get('duration', 3))
        except:
            dur_val = 3
        duration = st.number_input("Số tiết thực hiện:", value=dur_val, min_value=1, max_value=20)
        requirements = st.text_area("Yêu cầu cần đạt chuẩn SGV:", value=sel.get('requirements', ''), height=150)
    else:
        chapter_title = st.text_input("Chương / Chủ đề:", value="")
        lesson_title = st.text_input("Tên Bài dạy đầy đủ (Chuẩn SGV):", value="")
        duration = st.number_input("Số tiết thực hiện:", value=3, min_value=1, max_value=20)
        requirements = st.text_area("Yêu cầu cần đạt chuẩn SGV:", value="", height=120)

else:
    chapter_title = st.text_input("Chương / Chủ đề (Nhập nguyên văn SGK):", value="")
    lesson_title = st.text_input("Tên Bài dạy đầy đủ nguyên văn SGK (Nhập chính xác 100%):", value="")
    duration = st.number_input("Số tiết thực hiện:", value=3, min_value=1, max_value=20)
    requirements = st.text_area("Yêu cầu cần đạt chuẩn SGV (Dán từ SGV hoặc để trống AI tự soạn):", value="", height=150)

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ & CÔNG NGHỆ 4.0
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ & YẾU TỐ CHUYỂN ĐỔI SỐ</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn các yếu tố tích hợp hiện đại vào Kế hoạch bài dạy:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)",
        "Tích hợp Công cụ AI trong dạy và học (Gemini, ChatGPT, Canva AI...)",
        "Giáo dục STEM / STEAM",
        "Phát triển Tư duy phản biện & Giải quyết vấn đề thực tậi",
        "Tích hợp Giáo dục Đạo đức, Căn cước công dân & Pháp luật"
    ],
    default=[
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)",
        "Tích hợp Công cụ AI trong dạy và học (Gemini, ChatGPT, Canva AI...)"
    ]
)

# ==========================================
# HÀM TẠO FILE WORD CHUẨN ĐỊNH DẠNG CÔNG VĂN 5512
# ==========================================
def generate_doc(content_text, locked_chapter_title, locked_lesson_title):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    # 1. BẢNG HEADER TRƯỜNG & TỔ CHUYÊN MÔN
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    run_left = p_left.add_run(f"TRƯỜNG: {school_name.upper()}\nTỔ: {dept_name.upper()}")
    run_left.bold = True

    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    run_right = p_right.add_run(f"Họ và tên giáo viên:\n{teacher_name}")
    run_right.bold = True

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # 2. TIÊU ĐỀ BÀI DẠY (KHÓA NGUYÊN VĂN TỪ GIAO DIỆN)
    if locked_chapter_title:
        p_chap = doc.add_paragraph()
        p_chap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_chap.paragraph_format.space_before = Pt(14)
        p_chap.paragraph_format.space_after = Pt(2)
        r_chap = p_chap.add_run(locked_chapter_title.upper())
        r_chap.bold = True
        r_chap.font.size = Pt(13)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {locked_lesson_title.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)")
    r_sub.italic = True

    # 3. NỘI DUNG GIÁO ÁN
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
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(14)
        elif clean_text.startswith(("TIẾT ", "HOẠT ĐỘNG ", "Nội dung ", "Khối kiến thức ")):
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(13)
        elif clean_text.startswith(("1. ", "2. ", "3. ", "4. ")):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith(("a)", "b)", "c)", "d)")):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.15)
            run = p.add_run(clean_text)
            run.bold = True
        elif "Bước 1:" in clean_text or "Bước 2:" in clean_text or "Bước 3:" in clean_text or "Bước 4:" in clean_text:
            p.paragraph_format.left_indent = Inches(0.3)
            run = p.add_run(clean_text)
            run.bold = True
        else:
            p.add_run(clean_text)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# BƯỚC 4: TẠO GIÁO ÁN VÀ XUẤT FILE WORD
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    clean_api_key = api_key.strip() if api_key else ""
    
    if not clean_api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập hoặc tra cứu Tên bài dạy ở Bước 2!")
    else:
        try:
            integration_str = ", ".join(integrations) if integrations else "Không"
            
            prompt_main = f"""
            Bạn là Chuyên gia Giáo dục hàng đầu của Bộ Giáo dục và Đào tạo Việt Nam.
            Hãy biên soạn Kế hoạch bài dạy (Giáo án) CHI TIẾT KỸ LƯỠNG, CHUẨN 100% CÔNG VĂN 5512/BGDĐT.

            THÔNG TIN BÀI DẠY:
            - Môn học: {subject} | {grade} | Bộ sách: {book_series}
            - Chương/Chủ đề: {chapter_title}
            - Tên Bài dạy chính xác nguyên văn: {lesson_title}
            - Thời lượng thực hiện: {duration} tiết
            - YÊU CẦU CẦN ĐẠT SGV / MỤC TIÊU CẦN BÁM SÁT:
            {requirements if requirements else 'Tự đề xuất bám sát chương trình GDPT 2018 chuẩn SGV.'}
            - CÁC YẾU TỐ TÍCH HỢP CẦN CÓ: {integration_str}

            YÊU CẦU CẤU TRÚC KẾ HOẠCH BÀI DẠY (CÔNG VĂN 5512):

            I. MỤC TIÊU
            1. Về kiến thức: Liệt kê chi tiết kiến thức học sinh thu nhận được.
            2. Về năng lực:
               - Năng lực chung: Tự chủ và tự học, Giao tiếp và hợp tác, Giải quyết vấn đề và sáng tạo.
               - Năng lực đặc thù môn học: Chi tiết theo từng nội dung.
               - Năng lực Số / Ứng dụng CNTT & AI (Nếu có chọn tích hợp): Nêu rõ học sinh sử dụng công cụ nào (Padlet, Kahoot, Geogebra, AI...).
            3. Về phẩm chất: Yêu nước, Nhân ái, Chăm chỉ, Trung thực, Trách nhiệm.

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            1. Giáo viên: KHBD, SGK, SGV, thiết bị CNTT, phần mềm, phiếu học tập...
            2. Học sinh: SGK, vở ghi, thiết bị thông minh (nếu có)...

            III. TIẾN TRÌNH DẠY HỌC
            (Phân bổ hợp lý chi tiết cho {duration} tiết dạy)

            MỖI HOẠT ĐỘNG DẠY HỌC PHẢI BẮT BUỘC ĐỦ 4 MỤC CHUẨN 5512:
            a) Mục tiêu
            b) Nội dung
            c) Sản phẩm
            d) Tổ chức thực hiện:
               - Bước 1: Chuyển giao nhiệm vụ (GV giao nhiệm vụ cụ thể, rõ ràng)
               - Bước 2: Thực hiện nhiệm vụ (HS làm việc cá nhân/nhóm; GV quan sát, hỗ trợ)
               - Bước 3: Báo cáo, thảo luận (Đại diện báo cáo; lớp nhận xét, phản biện)
               - Bước 4: Kết luận, nhận định (GV chốt kiến thức, đánh giá)

            CÁC HOẠT ĐỘNG CẦN CÓ:
            - HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (Tạo tình huống / Kết nối)
            - HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI / KHÁM PHÁ (Chi tiết theo từng đơn vị kiến thức)
            - HOẠT ĐỘNG 3: LUYỆN TẬP (Hệ thống bài tập phân hóa)
            - HOẠT ĐỘNG 4: VẬN DỤNG (Bài tập thực tiễn)

            Hãy trình bày rõ ràng, văn phong sư phạm chuẩn mực, chi tiết không tóm tắt.
            """

            with st.spinner("⚡ AI đang soạn thảo Giáo án chi tiết chuẩn 5512... Vui lòng chờ trong giây lát..."):
                response = call_gemini(clean_api_key, model_name, [prompt_main])
                
                st.success("🎉 Đã hoàn thành biên soạn Giáo án chuẩn 5512!")
                
                doc_file = generate_doc(response.text, locked_chapter_title=chapter_title, locked_lesson_title=lesson_title)
                safe_file_name = re.sub(r'[\\/*?:"<>|]', "", lesson_title).replace(" ", "_")
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX) CHUẨN 5512",
                    data=doc_file,
                    file_name=f"KHBD_5512_{safe_file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            st.error(f"❌ {e}")
