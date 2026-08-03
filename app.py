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
from PIL import Image

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

# Hàm gọi API có cơ chế tự động thử lại (Retry) khi dính Quota 429
def call_gemini_with_retry(model, prompt, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = model.generate_content(prompt)
            return response
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 8  # Chờ 8s, 16s...
                    time.sleep(wait_time)
                    continue
            raise e

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
    
    # Chuẩn hóa các dòng model Gemini chính thức của Google
    model_name = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"],
        index=0,
        help="Khuyên dùng gemini-3.6-flash để tốc độ nhanh và hạn ngạch cao nhất"
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
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG (CHUẨN 5512)
        </div>
        <div style="font-size: 21px; font-weight: 600; color: #2563eb; margin-top: 10px;">
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
# BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Tải danh mục bài học từ taphuan.nxbgd.vn:</span>', unsafe_allow_html=True)
    if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                clean_model_name = model_name.replace("models/", "").strip()
                model = genai.GenerativeModel(clean_model_name)
                
                prompt_fetch = f"""
                Hãy đóng vai Cơ sở dữ liệu chính thức của NXB Giáo dục Việt Nam (taphuan.nxbgd.vn).
                Liệt kê ĐẦY ĐỦ tất cả các Bài học thuộc môn {subject} - {grade} (Bộ sách Kết nối tri thức/GDPT 2018).
                
                Trả về duy nhất dạng JSON mảng:
                [
                  {{
                    "chapter": "Tên Chương 1",
                    "lesson": "Tên Bài 1",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt chuẩn của bài"
                  }}
                ]
                Chỉ trả về mã JSON mảng [ ... ], không viết lời chào.
                """
                
                with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade}..."):
                    res = call_gemini_with_retry(model, prompt_fetch)
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã tải xong danh mục bài học chuẩn!")
            except Exception as e:
                st.warning("⚠️ API bận hoặc dính hạn ngạch. Hệ thống nạp bài học mẫu chuẩn để thầy dùng ngay:")
                st.session_state['fetched_lessons'] = [
                    {
                        "chapter": "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số",
                        "lesson": "Bài 1. Tính đơn điệu và cực trị của hàm số",
                        "duration": 3,
                        "req": "Nhận biết được tính đơn điệu, điểm cực trị của hàm số thông qua bảng biến thiên hoặc đồ thị."
                    },
                    {
                        "chapter": "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số",
                        "lesson": "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số",
                        "duration": 3,
                        "req": "Xác định được giá trị lớn nhất, giá trị nhỏ nhất của hàm số trên một đoạn."
                    }
                ]

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Hoặc tải lên File SGV (Sách Giáo Viên):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Tải lên File SGV:", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

req_default_value = st.session_state.get('req_from_sgv_file', '')

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
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
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
        final_req = req_default_value if req_default_value else current_item['req']
        requirements = st.text_area("YCĐ:", value=final_req, height=230, label_visibility="collapsed")
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        st.markdown('<span class="custom-label">Chương / Chủ đề:</span>', unsafe_allow_html=True)
        chapter_title = st.text_input("Chương:", value="", placeholder="Nhập tên chương...", label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Tên bài dạy:</span>', unsafe_allow_html=True)
        lesson_title = st.text_input("Tên bài:", value="", placeholder="Nhập tên bài...", label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
        duration = st.number_input("Số tiết:", value=2, label_visibility="collapsed")
        
    with col_i2:
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value=req_default_value, placeholder="Nhập yêu cầu cần đạt...", height=230, label_visibility="collapsed")

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
        "Phát triển Tư duy phản biện", 
        "Học tập qua Dự án (PBL)"
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
            model = genai.GenerativeModel(clean_model_name)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Soạn Kế hoạch bài dạy chuẩn 100% CÔNG VĂN 5512/BGDĐT:
            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - Yêu cầu cần đạt: {requirements}
            - YẾU TỐ TÍCH HỢP: {integration_str}

            QUY ĐỊNH:
            - Không dùng ký tự kẻ ngang (---).
            - Trình bày chính xác cấu trúc:
              I. MỤC TIÊU
              1. Kiến thức:
              2. Năng lực: (2.1. Năng lực chuyên môn, 2.2. Năng lực chung, 2.3. Năng lực Số / AI)
              3. Phẩm chất:
              II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
              III. TIẾN TRÌNH DẠY HỌC
              Mỗi hoạt động đúng 4 mục: a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện (Bước 1 -> Bước 4).
            """

            with st.spinner("✨ AI đang tạo Kế hoạch bài dạy 5512 (Đang ưu tiên xử lý...):"):
                response = call_gemini_with_retry(model, prompt)
                st.success("🎉 Tạo giáo án thành công!")
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
            err_str = str(e)
            if "429" in err_str:
                st.error("⏳ API hiện tại bị quá tải hạn ngạch trong ngắn hạn. Thầy vui lòng đợi 25 giây rồi bấm tạo lại, hoặc chọn mô hình gemini-1.5-flash ở thanh bên trái!")
            else:
                st.error(f"❌ Lỗi khi sinh giáo án: `{err_str}`")
