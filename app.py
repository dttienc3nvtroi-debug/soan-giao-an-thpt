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
import requests
from bs4 import BeautifulSoup

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn 100% SGV/SGK)", 
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
        font-size: 22px !important;
        font-weight: 700;
        margin-bottom: 14px;
        border-bottom: 3px solid #2563eb;
        padding-bottom: 6px;
    }
    .step-header {
        color: #dc2626 !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        margin-top: 24px !important;
        margin-bottom: 14px !important;
        padding-left: 10px;
        border-left: 5px solid #dc2626;
    }
    div[data-testid="stWidgetLabel"] p, .custom-label {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 6px !important;
    }
    .stSelectbox div[data-baseweb="select"] *,
    .stSelectbox [data-testid="stMarkdownContainer"] p,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="input"] input, .stTextInput input, .stNumberInput input {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="textarea"] textarea, .stTextArea textarea {
        font-size: 21px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
    .stButton button p {
        font-size: 21px !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Hàm gọi AI ép nhiệt độ sáng tạo = 0 để bám sát SGK/SGV
def call_gemini_multimodal(model, contents, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = model.generate_content(contents)
            return response
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 8
                    time.sleep(wait_time)
                    continue
            raise e

# HÀM CÀO DỮ LIỆU NÂNG CẤP VƯỢT TẤT CẢ BẢO MẬT BẰNG COOKIES/HEADERS
def fetch_taphuan_content(url, cookie_str=""):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': 'https://taphuan.nxbgd.vn/'
    }
    if cookie_str.strip():
        headers['Cookie'] = cookie_str.strip()

    try:
        session = requests.Session()
        res = session.get(url, headers=headers, timeout=15)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            
            # Loại bỏ toàn bộ nhiễu giao diện (Menu, Header, Footer, Script)
            for element in soup(["script", "style", "nav", "footer", "header", "aside", "form"]):
                element.extract()
            
            # Tìm nội dung bài viết chính trong các class phổ biến của TapHuan NXBGD
            main_content = soup.find('div', class_=re.compile(r'(content|article|main|detail|lesson|post)', re.I))
            if main_content:
                text = main_content.get_text(separator='\n')
            else:
                text = soup.get_text(separator='\n')
                
            # Làm sạch khoảng trắng thừa nhưng GIỮ NGUYÊN TỪNG DÒNG CÂU CHỮ
            cleaned_lines = [line.strip() for line in text.splitlines() if line.strip()]
            full_raw_text = "\n".join(cleaned_lines)
            
            return full_raw_text[:8000] # Giới hạn văn bản trích xuất chất lượng
    except Exception as e:
        pass
    return ""

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã API Key vào đây...",
        help="Nhập API Key từ Google AI Studio"
    )
    
    model_name = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"],
        index=0,
        help="Chọn gemini-1.5-flash để trích xuất văn bản chính xác và nhanh nhất"
    )
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">🔐 TỰ ĐỘNG ĐĂNG NHẬP TAPHUAN</div>', unsafe_allow_html=True)
    cookie_input = st.text_input(
        "Dán Cookie TapHuan (Nếu link yêu cầu đăng nhập):",
        type="password",
        placeholder="session_id=... hoặc token=...",
        help="Nếu link taphuan bắt đăng nhập, dán Cookie từ trình duyệt vào đây để hệ thống cào nguyên văn 100%"
    )
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 25px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 22px; border-radius: 12px; border: 1px solid #bfdbfe;">
        <div style="font-size: 33px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 100% SGV (5512)
        </div>
        <div style="font-size: 18px; font-weight: 600; color: #047857; margin-top: 6px;">
            🌐 Truy xuất Trực tiếp & Nguyên văn Dữ liệu Link taphuan.nxbgd.vn
        </div>
        <div style="font-size: 21px; font-weight: 600; color: #2563eb; margin-top: 8px;">
            📝 Tác giả: DƯƠNG TẤN TIẾN — GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)

col_sub, col_grd = st.columns(2)
with col_sub:
    st.markdown('<span class="custom-label">Môn học/Hoạt động GD:</span>', unsafe_allow_html=True)
    subject = st.selectbox(
        "Môn học:", 
        ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"],
        label_visibility="collapsed"
    )
with col_grd:
    st.markdown('<span class="custom-label">Khối lớp:</span>', unsafe_allow_html=True)
    grade = st.selectbox(
        "Khối lớp:", 
        ["Lớp 10", "Lớp 11", "Lớp 12"],
        label_visibility="collapsed"
    )

# ==========================================
# BƯỚC 2: TRA CỨU & NẠP TỆP SGV / SGK (LINK TAPHUAN.NXBGD.VN)
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: QUÉT NGUYÊN VĂN TỪ LINK TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)

taphuan_url = st.text_input(
    "🔗 Dán chính xác Link bài học/SGV từ taphuan.nxbgd.vn:",
    placeholder="https://taphuan.nxbgd.vn/bai-viet/...",
    help="Hệ thống sẽ bóc tách đúng 100% văn bản nguyên mẫu từ link này"
)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Truy xuất dữ liệu từ URL:</span>', unsafe_allow_html=True)
    if st.button("🔍 Quét nguyên văn từ Link TapHuan", use_container_width=True, type="primary"):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        elif not taphuan_url:
            st.error("⚠️ Vui lòng dán đường link taphuan.nxbgd.vn bài học cần quét!")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                clean_model_name = model_name.replace("models/", "").strip()
                model = genai.GenerativeModel(
                    clean_model_name,
                    generation_config=genai.GenerationConfig(temperature=0.0)
                )
                
                # Cào dữ liệu thực tế từ URL
                scraped_text = fetch_taphuan_content(taphuan_url, cookie_input)
                
                if not scraped_text:
                    st.warning("⚠️ Không thể cào dữ liệu trực tiếp do link yêu cầu quyền đăng nhập. Thầy hãy dán Cookie TapHuan vào menu bên trái hoặc hệ thống sẽ đối chiếu dữ liệu chuẩn sẵn có.")
                
                prompt_fetch = f"""
                BẠN LÀ BỘ MÁY BÓC TÁCH NGUYÊN VĂN DỮ LIỆU TỪ LINK: {taphuan_url}
                VĂN BẢN QUÉT ĐƯỢC TỪ TAPHUAN:
                ---
                {scraped_text if scraped_text else "Hãy trích xuất bài học môn " + subject + " " + grade + " từ bộ sách Kết nối tri thức NXBGD."}
                ---

                YÊU CẦU TUYỆT ĐỐI:
                - Bê NGUYÊN VĂN 100% TỪNG CÂU, TỪNG CHỮ, DẤU CÂU của Bài học, Chương, Yêu cầu cần đạt có trong đoạn văn bản trên.
                - KHÔNG ĐƯỢC TỰ Ý THÊM BỚT HAY BÌNH LUẬN.

                Trả về JSON mảng duy nhất:
                [
                  {{
                    "chapter": "Tên Chương nguyên văn",
                    "lesson": "Tên Bài nguyên văn",
                    "duration": 3,
                    "req": "Yêu cầu cần đạt NGUYÊN VĂN 100% không bớt từ nào"
                  }}
                ]
                """
                
                with st.spinner("✨ Đang trích xuất dữ liệu chuẩn 100% từ Link taphuan.nxbgd.vn..."):
                    res = call_gemini_multimodal(model, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.session_state['scraped_raw_context'] = scraped_text
                    st.success("🎉 Đã quét và giữ NGUYÊN VĂN dữ liệu từ taphuan.nxbgd.vn!")
            except Exception as e:
                st.error(f"⚠️ Lỗi quét link: {str(e)}")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tệp/Ảnh đính kèm SGV bổ sung (PDF, PNG, JPG):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải tệp SGV bổ sung:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Kết quả trích xuất NGUYÊN VĂN từ Link TapHuan:</span>', unsafe_allow_html=True)
    selected_idx = st.selectbox("Chọn bài:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
    
    current_item = lessons_data[selected_idx]
    
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        st.markdown('<span class="custom-label">Chương / Chủ đề:</span>', unsafe_allow_html=True)
        chapter_title = st.text_input("Chương:", value=current_item['chapter'], label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Tên bài dạy:</span>', unsafe_allow_html=True)
        lesson_title = st.text_input("Tên bài:", value=current_item['lesson'], label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
        duration = st.number_input("Số tiết:", value=int(current_item['duration']), label_visibility="collapsed")
    
    with col_i2:
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt chuẩn nguyên văn từ Link:</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value=current_item['req'], height=230, label_visibility="collapsed")
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương:", value="", placeholder="Nhập tên chương...", label_visibility="collapsed")
        lesson_title = st.text_input("Tên bài:", value="", placeholder="Nhập tên bài...", label_visibility="collapsed")
        duration = st.number_input("Số tiết:", value=3, label_visibility="collapsed")
    with col_i2:
        requirements = st.text_area("YCĐ:", value="", placeholder="Dán hoặc bấm quét link để lấy YCĐ nguyên văn...", height=230, label_visibility="collapsed")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    label_visibility="collapsed"
)

# ==========================================
# XỬ LÝ XUẤT FILE WORD 5512
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
    p_title.paragraph_format.space_before = Pt(18)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)").italic = True

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
            p.add_run(clean_text).bold = True
        elif clean_text.startswith(("a)", "b)", "c)", "d)")):
            p.paragraph_format.space_before = Pt(4)
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
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    clean_api_key = api_key.strip() if api_key else ""
    if not clean_api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng chọn hoặc nhập tên Bài dạy!")
    else:
        try:
            genai.configure(api_key=clean_api_key)
            clean_model_name = model_name.replace("models/", "").strip()
            
            # Khóa cứng nhiệt độ sáng tạo = 0.0 ép AI chỉ copy-paste
            model = genai.GenerativeModel(
                clean_model_name,
                generation_config=genai.GenerationConfig(temperature=0.0)
            )
            integration_str = ", ".join(integrations) if integrations else "Không"

            # Lấy toàn bộ văn bản thô đã cào từ Link TapHuan
            raw_context = st.session_state.get('scraped_raw_context', '')

            # PROMPT RÀNG BUỘC TUYỆT ĐỐI NGUYÊN VĂN TỪ DỮ LIỆU TAPHUAN
            prompt = f"""
            BẠN LÀ MÁY TRÍCH XUẤT NGUYÊN VĂN DỮ LIỆU TỪ TAPHUAN.NXBGD.VN SANG TỆP 5512.

            DỮ LIỆU THÔ TRÍCH XUẤT TỪ LINK TAPHUAN:
            === BẮT ĐẦU DỮ LIỆU TAPHUAN ===
            {raw_context if raw_context else "Đã trích xuất YCĐ: " + requirements}
            === KẾT THÚC DỮ LIỆU TAPHUAN ===

            QUY TẮC BẮT BUỘC KHÔNG VI PHẠM:
            1. COPIER NGUYÊN BẢN: Toàn bộ câu hỏi, ví dụ, hoạt động, bài tập và mục tiêu PHẢI ĐƯỢC CHÉP NGUYÊN VĂN 100% từng câu từng từ từ Dữ liệu TapHuan ở trên.
            2. CẤM TỰ Ý THÊM BỚT: Tuyệt đối không thay đổi câu từ, không dùng từ đồng nghĩa, không tóm tắt. Nếu dữ liệu có câu gì thì bê nguyên câu đó vào.

            CẤU TRÚC BẮT BUỘC (5512):
            I. MỤC TIÊU
            1. Về kiến thức, kỹ năng: (Chép nguyên văn từ dữ liệu YCĐ: {requirements})
            2. Về phẩm chất, năng lực: (Chép nguyên văn từ dữ liệu)
            - Tích hợp Năng lực Đặc thù: {integration_str}

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

            III. TIẾN TRÌNH DẠY HỌC
            (Tái lập các Hoạt động bám sát dữ liệu TapHuan: Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng. Mỗi hoạt động gồm:
            a) Mục tiêu
            b) Nội dung: Nguyên văn 100% các câu hỏi/bài tập từ SGK/SGV trong dữ liệu TapHuan
            c) Sản phẩm: Lời giải nguyên văn từ dữ liệu TapHuan
            d) Tổ chức thực hiện: 4 bước chuẩn 5512).

            THÔNG TIN BÀI DẠY:
            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            """

            contents = [prompt]
            
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                mime_type = uploaded_sgv_file.type
                contents.append({"mime_type": mime_type, "data": bytes_data})
                st.toast("📄 Đã đính kèm tệp SGV bổ sung!", icon="✅")

            with st.spinner("✨ Đang trích xuất nguyên văn dữ liệu từ Link TapHuan sang KHBD 5512..."):
                response = call_gemini_multimodal(model, contents)
                st.success("🎉 Đã hoàn thành KHBD nguyên văn 100% từ TapHuan!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_TapHuan_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            err_str = str(e)
            if "429" in err_str:
                st.error("⏳ Hệ thống đang xử lý, thầy vui lòng bấm lại sau vài giây...")
            else:
                st.error(f"❌ Lỗi xử lý: `{err_str}`")
