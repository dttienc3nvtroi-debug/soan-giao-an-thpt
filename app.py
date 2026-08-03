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
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
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
    section[data-testid="sidebar"] {
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
# HÀM XỬ LÝ MODEL LỌC SẠCH TTS & LỖI MODALITY
# ==========================================
def get_available_models():
    """Ưu tiên các dòng Flash để phản hồi tốc độ cao"""
    try:
        available = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                if "-tts" in name.lower() or "audio" in name.lower():
                    continue
                available.append(name)
        # Sắp xếp ưu tiên các dòng Flash lên đầu
        flash_models = [m for m in available if "flash" in m.lower()]
        other_models = [m for m in available if "flash" not in m.lower()]
        return flash_models + other_models
    except Exception:
        return ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-1.5-pro"]

def call_gemini_fast(selected_model_str, contents):
    """Gọi Gemini AI tối ưu tốc độ phản hồi"""
    clean_selected = selected_model_str.replace("models/", "").strip()
    
    try:
        model = genai.GenerativeModel(
            model_name=clean_selected,
            generation_config=genai.GenerationConfig(
                temperature=0.0,
                max_output_tokens=2048 # Giới hạn độ dài để xử lý siêu nhanh
            )
        )
        return model.generate_content(contents)
    except Exception:
        # Fallback nhanh sang flash
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=genai.GenerationConfig(temperature=0.0)
        )
        return model.generate_content(contents)

# HÀM CÀO DỮ LIỆU SIÊU TỐC TỪ LINK
def fetch_data_from_taphuan_fast(url, cookie_str=""):
    if not url:
        return ""
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9',
        'Referer': 'https://taphuan.nxbgd.vn/'
    }
    if cookie_str:
        headers['Cookie'] = cookie_str

    try:
        # Giới hạn Timeout 5 giây để tránh treo
        res = requests.get(url, headers=headers, timeout=5)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            for tag in soup(["script", "style", "nav", "footer", "header"]):
                tag.extract()
            text = soup.get_text(separator=' ')
            clean_text = " ".join(text.split())
            return clean_text[:4000] # Lấy 4000 ký tự đầu là đủ thông tin
        return f"HTTP_ERROR_{res.status_code}"
    except Exception as e:
        return f"FETCH_EXCEPTION: {str(e)}"

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã API Key vào đây..."
    )
    
    default_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    
    if api_key:
        genai.configure(api_key=api_key.strip())
        fetched_models = get_available_models()
        if fetched_models:
            default_models = fetched_models

    model_name = st.selectbox(
        "Mô hình AI xử lý (Khuyên dùng Flash để xử lý nhanh):",
        default_models,
        index=0
    )
    st.markdown("---")
    
    st.markdown('<div class="sidebar-title">🌐 CẤU HÌNH TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)
    taphuan_cookie = st.text_input(
        "Cookie TapHuan (Nếu có):",
        type="password"
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
            🌐 TRUY XUẤT TRỰC TIẾP TỪ LINK TAPHUAN.NXBGD.VN
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
    subject = st.selectbox(
        "Môn học:", 
        ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"]
    )
with col_grd:
    grade = st.selectbox(
        "Khối lớp:", 
        ["Lớp 10", "Lớp 11", "Lớp 12"]
    )

# ==========================================
# BƯỚC 2: TRUY XUẤT DỮ LIỆU
# ==========================================
st.markdown('<div class="step-header">🔗 BƯỚC 2: NHẬP LINK BÀI HỌC TỪ TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)

taphuan_url = st.text_input(
    "🔗 Nhập đường Link bài học:",
    value="https://taphuan.nxbgd.vn/",
    placeholder="https://taphuan.nxbgd.vn/..."
)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    if st.button("⚡ Phân tích siêu tốc từ Link", use_container_width=True, type="primary"):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                
                with st.spinner("⚡ Đang cào dữ liệu nhanh..."):
                    raw_scraped_text = fetch_data_from_taphuan_fast(taphuan_url, taphuan_cookie)
                
                context_payload = raw_scraped_text if "HTTP_ERROR" not in raw_scraped_text and raw_scraped_text.strip() else f"Môn {subject} - {grade}"

                prompt_fetch = f"""
                Trích xuất Yêu cầu cần đạt cho môn {subject} - {grade} từ văn bản sau:
                {context_payload}

                Trả về ĐÚNG 1 MẢNG JSON duy nhất, không thêm lời thoại:
                [
                  {{
                    "chapter": "Tên Chương",
                    "lesson": "Tên Bài",
                    "duration": 3,
                    "req": "Yêu cầu cần đạt nguyên văn"
                  }}
                ]
                """
                
                with st.spinner("⚡ AI đang xử lý..."):
                    res = call_gemini_fast(model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    
                    # Trích xuất JSON siêu tốc
                    match = re.search(r'\[\s*\{.*\}\s*\]', raw_text, re.DOTALL)
                    clean_json_str = match.group(0) if match else raw_text

                    try:
                        st.session_state['fetched_lessons'] = json.loads(clean_json_str)
                    except Exception:
                        st.session_state['fetched_lessons'] = [{
                            "chapter": f"Bài học môn {subject}",
                            "lesson": f"Nội dung bài học {grade}",
                            "duration": 3,
                            "req": raw_text
                        }]
                    
                    st.success("⚡ Hoàn thành bóc tách siêu tốc!")
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")

with col_file_upload:
    uploaded_sgv_file = st.file_uploader(
        "Tải tệp SGV bổ sung (Nếu có):", 
        type=["pdf", "png", "jpg", "jpeg"]
    )

# HIỂN THỊ KẾT QUẢ
if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item.get('chapter', '')} - {item.get('lesson', '')}" for item in lessons_data]
    
    selected_idx = st.selectbox("Bài học trích xuất:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x])
    current_item = lessons_data[selected_idx]
    
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương:", value=current_item.get('chapter', ''))
        lesson_title = st.text_input("Tên bài:", value=current_item.get('lesson', ''))
        duration = st.number_input("Số tiết:", value=int(current_item.get('duration', 3)))
    
    with col_i2:
        requirements = st.text_area("Yêu cầu cần đạt:", value=current_item.get('req', ''), height=200)
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương:", value="")
        lesson_title = st.text_input("Tên bài:", value="")
        duration = st.number_input("Số tiết:", value=3)
    with col_i2:
        requirements = st.text_area("Yêu cầu cần đạt:", value="", height=200)

# ==========================================
# BƯỚC 3: TÍCH HỢP
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
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)"]
)

# ==========================================
# XUẤT FILE WORD 5512
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
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập hoặc chọn tên Bài dạy!")
    else:
        try:
            genai.configure(api_key=clean_api_key)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Soạn KHBD 5512 môn {subject} - {grade}.
            Bài dạy: {lesson_title} ({chapter_title}). Thời lượng: {duration} tiết.
            
            Yêu cầu cần đạt nguyên văn:
            {requirements}

            CẤU TRÚC 5512:
            I. MỤC TIÊU (Kiến thức, Phẩm chất, Năng lực, Tích hợp: {integration_str})
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Các Hoạt động Mở đầu, Hình thành KT, Luyện tập, Vận dụng chuẩn 4 bước: a.Mục tiêu, b.Nội dung, c.Sản phẩm, d.Tổ chức thực hiện)
            """

            contents = [prompt]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                mime_type = uploaded_sgv_file.type
                contents.append({"mime_type": mime_type, "data": bytes_data})

            with st.spinner("⚡ Đang tạo KHBD..."):
                response = call_gemini_fast(model_name, contents)
                st.success("🎉 Tạo KHBD thành công!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            st.error(f"❌ Lỗi xử lý: `{str(e)}`")
