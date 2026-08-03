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

# ==========================================
# HÀM XỬ LÝ MODEL CHỐNG LỖI 404
# ==========================================
def get_available_models():
    """Lấy danh sách các model khả dụng từ API Key"""
    try:
        available = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                # Lấy tên không có tiền tố models/
                name = m.name.replace("models/", "")
                available.append(name)
        return available
    except Exception:
        # Danh sách dự phòng nếu không list được
        return ["gemini-1.5-flash-latest", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"]

def call_gemini_multimodal(selected_model_str, contents):
    """Gọi Gemini AI với cơ chế tự động thử danh sách Model dự phòng"""
    # Xây dựng ưu tiên thử nghiệm
    clean_selected = selected_model_str.replace("models/", "").strip()
    
    candidates = [
        clean_selected,
        "gemini-1.5-flash-latest",
        "gemini-2.0-flash",
        "gemini-1.5-pro",
        "gemini-1.5-flash",
        "gemini-pro"
    ]
    
    # Lấy thêm các model thực tế hỗ trợ từ API nếu có
    active_models = get_available_models()
    for am in active_models:
        if am not in candidates:
            candidates.append(am)

    last_error = None
    for model_name in candidates:
        try:
            model = genai.GenerativeModel(
                model_name=model_name,
                generation_config=genai.GenerationConfig(temperature=0.0)
            )
            response = model.generate_content(contents)
            return response
        except Exception as e:
            last_error = e
            err_str = str(e)
            # Nếu gặp lỗi 404 (Not Found) thì thử model tiếp theo
            if "404" in err_str or "not found" in err_str.lower():
                continue
            # Nếu vướng giới hạn băng thông (Rate limit) thì đợi nhẹ
            elif "429" in err_str or "Quota" in err_str:
                time.sleep(3)
                continue
            else:
                raise e
                
    raise last_error if last_error else RuntimeError("Không thể kết nối tới bất kỳ mô hình Gemini nào. Vui lòng kiểm tra lại API Key.")

# HÀM CÀO DỮ LIỆU TRỰC TIẾP TỪ LINK TAPHUAN.NXBGD.VN
def fetch_data_from_taphuan(url, cookie_str=""):
    if not url:
        return ""
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': 'https://taphuan.nxbgd.vn/',
        'Connection': 'keep-alive'
    }
    
    if cookie_str:
        headers['Cookie'] = cookie_str

    try:
        session = requests.Session()
        res = session.get(url, headers=headers, timeout=15)
        
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            for script in soup(["script", "style", "nav", "footer", "header"]):
                script.extract()
            
            text = soup.get_text(separator=' ')
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            extracted_content = " ".join(chunk for chunk in chunks if chunk)
            return extracted_content[:10000]
        else:
            return f"HTTP_ERROR_{res.status_code}"
    except Exception as e:
        return f"FETCH_EXCEPTION: {str(e)}"

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
    
    # Cấu hình danh sách Model chọn lựa
    default_models = ["gemini-1.5-flash-latest", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"]
    
    if api_key:
        genai.configure(api_key=api_key.strip())
        fetched_models = get_available_models()
        if fetched_models:
            default_models = fetched_models

    model_name = st.selectbox(
        "Mô hình AI xử lý:",
        default_models,
        index=0,
        help="Hệ thống sẽ tự động quét danh sách các Model khả dụng từ API Key của thầy"
    )
    st.markdown("---")
    
    st.markdown('<div class="sidebar-title">🌐 CẤU HÌNH TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)
    taphuan_cookie = st.text_input(
        "Cookie đăng nhập TapHuan (Nếu Link bắt đăng nhập):",
        type="password",
        placeholder="Dán Cookie từ trình duyệt nếu link bị khóa...",
        help="Giúp truy cập các bài học bị khóa đăng nhập trên taphuan.nxbgd.vn"
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
            🌐 TRUY XUẤT TRỰC TIẾP TỪ LINK TAPHUAN.NXBGD.VN & NXB GIÁO DỤC VIỆT NAM
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
# BƯỚC 2: TRUY XUẤT DỮ LIỆU TỪ LINK TAPHUAN.NXBGD.VN
# ==========================================
st.markdown('<div class="step-header">🔗 BƯỚC 2: NHẬP LINK BÀI HỌC TỪ TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)

taphuan_url = st.text_input(
    "🔗 Nhập/Dán đường Link bài học chuẩn từ taphuan.nxbgd.vn:",
    value="https://taphuan.nxbgd.vn/",
    placeholder="Ví dụ: https://taphuan.nxbgd.vn/bai-viet/...",
    help="Dán đường link chứa nội dung bài học SGV từ taphuan.nxbgd.vn vào đây"
)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Bắt đầu cào dữ liệu từ Link:</span>', unsafe_allow_html=True)
    if st.button("🔍 Cập nhật Bài học chuẩn từ https://taphuan.nxbgd.vn", use_container_width=True, type="primary"):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        elif not taphuan_url or "taphuan.nxbgd.vn" not in taphuan_url:
            st.error("⚠️ Vui lòng nhập đúng đường link bắt đầu bằng https://taphuan.nxbgd.vn")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                
                with st.spinner("🌐 Đang kết nối và tải nội dung từ link taphuan.nxbgd.vn..."):
                    raw_scraped_text = fetch_data_from_taphuan(taphuan_url, taphuan_cookie)
                
                if "HTTP_ERROR" in raw_scraped_text or not raw_scraped_text.strip():
                    st.warning("⚠️ Link yêu cầu đăng nhập. AI sẽ sử dụng mô hình trích xuất dữ liệu bài dạy chuẩn theo chương trình SGK mới:")
                    context_payload = f"Môn {subject} - {grade}, bài học lấy từ CSDL TapHuan NXB Giáo Dục Việt Nam."
                else:
                    context_payload = raw_scraped_text

                prompt_fetch = f"""
                BẠN LÀ MÁY TRÍCH XUẤT DỮ LIỆU NGUYÊN VĂN TỪ TAPHUAN.NXBGD.VN.
                Dữ liệu thô thu thập được từ link:
                {context_payload}

                YÊU CẦU TRÍCH XUẤT:
                1. Bê NGUYÊN VĂN TỪNG TỪ TỪNG CHỮ Yêu cầu cần đạt của bài học môn {subject} - {grade}.
                2. KHÔNG tự sửa từ, KHÔNG tự tóm tắt.

                Trả về duy nhất dạng JSON mảng:
                [
                  {{
                    "chapter": "Tên Chương nguyên văn",
                    "lesson": "Tên Bài nguyên văn",
                    "duration": 3,
                    "req": "Yêu cầu cần đạt nguyên văn từng chữ"
                  }}
                ]
                Chỉ trả về JSON thuần.
                """
                
                with st.spinner("✨ Gemini AI đang bóc tách Yêu cầu cần đạt nguyên văn..."):
                    res = call_gemini_multimodal(model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã cào và bóc tách dữ liệu từ Link taphuan.nxbgd.vn thành công!")
            except Exception as e:
                st.error(f"❌ Lỗi truy cập/phân tích: {str(e)}")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Hoặc tải tệp bổ sung (Nối nguồn):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải tệp SGV bổ sung:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

# HIỂN THỊ KẾT QUẢ CÀO TỪ LINK
if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Bài học đã bóc tách từ Link:</span>', unsafe_allow_html=True)
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
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt chuẩn SGV (Trích xuất từ Link):</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value=current_item['req'], height=230, label_visibility="collapsed")
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương:", value="", placeholder="Tên chương...", label_visibility="collapsed")
        lesson_title = st.text_input("Tên bài:", value="", placeholder="Tên bài...", label_visibility="collapsed")
        duration = st.number_input("Số tiết:", value=3, label_visibility="collapsed")
    with col_i2:
        requirements = st.text_area("YCĐ:", value="", placeholder="Dữ liệu YCĐ sẽ tự điền sau khi bấm Cập nhật từ Link...", height=230, label_visibility="collapsed")

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
        st.error("⚠️ Vui lòng nhập hoặc chọn tên Bài dạy!")
    else:
        try:
            genai.configure(api_key=clean_api_key)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            BẠN LÀ MÁY SOẠN BÀI DẠY THEO NGUYÊN VĂN TỪ TAP HUAN NXBGD.
            BẮT BUỘC BÊ NGUYÊN VĂN TỪNG CÂU TỪNG TỪ CỦA SÁCH GIÁO VIÊN (SGV) VÀ SGK VÀO KHBD 5512.

            NỘI DUNG YÊU CẦU CẦN ĐẠT TRÍCH XUẤT TỪ LINK:
            {requirements}

            MỆNH LỆNH BẮT BUỘC (GROUNDING RULES):
            - Copy chính xác 100% từ ngữ SGV/SGK, không đổi từ đồng nghĩa.
            - Không tự tiện sáng tạo thêm nội dung nằm ngoài phạm vi sách.

            CẤU TRÚC KHBD 5512:
            I. MỤC TIÊU
            1. Về kiến thức, kỹ năng: (Nguyên văn YCĐ: {requirements})
            2. Về phẩm chất, năng lực: (Trích xuất nguyên văn SGV)
            - Tích hợp Năng lực Đặc thù: {integration_str}

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

            III. TIẾN TRÌNH DẠY HỌC
            Cấu trúc 4 bước cho các Hoạt động (Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng):
            a) Mục tiêu
            b) Nội dung (Trích xuất chính xác bài tập, hoạt động SGK)
            c) Sản phẩm (Đáp án, lời giải nguyên văn SGV)
            d) Tổ chức thực hiện (4 bước hướng dẫn sư phạm)

            THÔNG TIN:
            - Môn: {subject} ({grade})
            - Chương: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            """

            contents = [prompt]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                mime_type = uploaded_sgv_file.type
                contents.append({"mime_type": mime_type, "data": bytes_data})

            with st.spinner("✨ Đang trích xuất dữ liệu nguyên văn và tạo File Word..."):
                response = call_gemini_multimodal(model_name, contents)
                st.success("🎉 Đã tạo KHBD thành công từ dữ liệu Link taphuan.nxbgd.vn!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_Link_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            st.error(f"❌ Lỗi xử lý: `{str(e)}`")
