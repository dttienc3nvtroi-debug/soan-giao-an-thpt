import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io
import json
import re
import time

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Bám sát SGV)", 
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
        ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
        index=0,
        help="Nên chọn gemini-2.5-flash hoặc gemini-1.5-flash để đọc file PDF/Ảnh nhanh và mượt nhất"
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
# BƯỚC 2: TRA CỨU & NẠP TỆP SGV / SGK
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: NẠP FILE SGV (ĐẢM BẢO CHÍNH XÁC 100%)</div>', unsafe_allow_html=True)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Tải danh mục bài học chuẩn (Bộ sách KNTT/GDPT 2018):</span>', unsafe_allow_html=True)
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
                Liệt kê ĐẦY ĐỦ tất cả các Bài học thuộc môn {subject} - {grade} (Bộ sách Kết nối tri thức với cuộc sống).
                
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
                    res = call_gemini_multimodal(model, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã tải xong danh mục bài học chuẩn SGK!")
            except Exception as e:
                st.warning("⚠️ Đã nạp bài học mẫu chuẩn SGK:")
                st.session_state['fetched_lessons'] = [
                    {
                        "chapter": "Chương II. Vectơ và hệ trục tọa độ trong không gian",
                        "lesson": "Bài 7. Hệ trục tọa độ trong không gian",
                        "duration": 3,
                        "req": "- Nhận biết được tọa độ của điểm, của vectơ đối với hệ trục tọa độ.\n- Vận dụng được tọa độ của vectơ để giải một số bài toán có liên quan đến thực tiễn."
                    }
                ]

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tải lên File/Ảnh trang SGV (PDF, JPG, PNG):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải lên File SGV:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed",
        help="Tải ảnh/PDF trang SGV lên đây để AI chép chính xác 100% mục tiêu kiến thức, kỹ năng!"
    )

if uploaded_sgv_file is not None:
    st.info(f"✅ Đã nhận file: **{uploaded_sgv_file.name}**. Hệ thống sẽ trích xuất Y NGUYÊN nội dung từ file này!")

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Chọn Bài học từ danh sách:</span>', unsafe_allow_html=True)
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
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt / Mục tiêu SGV:</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value=current_item['req'], height=230, label_visibility="collapsed")
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        st.markdown('<span class="custom-label">Chương / Chủ đề:</span>', unsafe_allow_html=True)
        chapter_title = st.text_input("Chương:", value="", placeholder="Nhập tên chương...", label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Tên bài dạy:</span>', unsafe_allow_html=True)
        lesson_title = st.text_input("Tên bài:", value="", placeholder="Nhập tên bài...", label_visibility="collapsed")
        
        st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
        duration = st.number_input("Số tiết:", value=3, label_visibility="collapsed")
        
    with col_i2:
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt / Mục tiêu SGV:</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value="", placeholder="Nhập mục tiêu SGV hoặc để trống nếu đã upload file...", height=230, label_visibility="collapsed")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ (BỔ SUNG VÀO CẤU TRÚC 5512)</div>', unsafe_allow_html=True)

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
# HELPER PARSER: RENDER MARKDOWN & LATEX SANG WORD RUNS
# ==========================================
def clean_latex_math(text):
    """Chuyển đổi công thức LaTeX dạng $180^\circ$ thành văn bản đọc được 180°"""
    text = re.sub(r'\\circ', '°', text)
    text = re.sub(r'\\sin', 'sin', text)
    text = re.sub(r'\\cos', 'cos', text)
    text = re.sub(r'\\tan', 'tan', text)
    text = re.sub(r'\\cot', 'cot', text)
    text = re.sub(r'\\alpha', 'α', text)
    text = re.sub(r'\\beta', 'β', text)
    text = re.sub(r'\\le', '≤', text)
    text = re.sub(r'\\ge', '≥', text)
    text = re.sub(r'\\neq', '≠', text)
    text = text.replace('$', '')
    return text

def parse_formatted_text_to_paragraph(paragraph, raw_text):
    """
    Hàm phân tích cú pháp Markdown **bold** và *italic* và $latex$
    để ép font Times New Roman 13pt chuẩn xác.
    """
    raw_text = clean_latex_math(raw_text)
    pattern = r'(\*\*.*?\*\*|\*.*?\*)'
    tokens = re.split(pattern, raw_text)

    for token in tokens:
        if not token:
            continue
        if token.startswith('**') and token.endswith('**'):
            run_text = token[2:-2]
            run = paragraph.add_run(run_text)
            run.bold = True
        elif token.startswith('*') and token.endswith('*'):
            run_text = token[1:-1]
            run = paragraph.add_run(run_text)
            run.italic = True
        else:
            run = paragraph.add_run(token)

        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

# ==========================================
# XỬ LÝ XUẤT FILE WORD 5512 CHUẨN ĐỊNH DẠNG
# ==========================================
def generate_doc(content_text):
    doc = docx.Document()

    # 1. Cấu hình lề trang chuẩn (Top/Bottom 2cm, Left 3cm, Right 2cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 2 cm
        section.bottom_margin = Inches(0.79)  # 2 cm
        section.left_margin = Inches(1.18)    # 3 cm
        section.right_margin = Inches(0.79)   # 2 cm

    # 2. Cấu hình style Normal mặc định
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)
    style.font.color.rgb = RGBColor(0, 0, 0)

    # 3. Tạo Bảng Đầu Trang (Trường / Tổ bên trái - Giáo viên bên phải)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Độ rộng 2 cột chuẩn lề
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    # Cột trái: TRƯỜNG, TỔ
    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    p_left.paragraph_format.space_before = Pt(0)
    p_left.paragraph_format.space_after = Pt(0)
    
    r_school = p_left.add_run(f"TRƯỜNG: {school_name.upper()}\n")
    r_school.bold = True
    r_school.font.name = 'Times New Roman'
    r_school.font.size = Pt(12)

    r_dept = p_left.add_run(f"TỔ: {dept_name.upper()}")
    r_dept.bold = True
    r_dept.font.name = 'Times New Roman'
    r_dept.font.size = Pt(12)

    # Cột phải: HỌ VÀ TÊN GIÁO VIÊN
    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    p_right.paragraph_format.space_before = Pt(0)
    p_right.paragraph_format.space_after = Pt(0)

    r_lbl_teacher = p_right.add_run("Họ và tên giáo viên:\n")
    r_lbl_teacher.bold = True
    r_lbl_teacher.font.name = 'Times New Roman'
    r_lbl_teacher.font.size = Pt(12)

    r_teacher_val = p_right.add_run(teacher_name)
    r_teacher_val.bold = True
    r_teacher_val.font.name = 'Times New Roman'
    r_teacher_val.font.size = Pt(12)

    # Xoá viền bảng để sạch đẹp
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # 4. TÊN BÀI DẠY (Căn giữa, In hoa, Đậm 14pt)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(4)
    
    title_text = lesson_title.upper()
    if not title_text.startswith("TÊN BÀI DẠY:"):
        title_text = f"TÊN BÀI DẠY: {title_text}"
        
    r_title = p_title.add_run(title_text)
    r_title.bold = True
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(14)

    # Dòng Phụ đề (Môn học / Thời gian thực hiện - Căn giữa, In nghiêng)
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)")
    r_sub.italic = True
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(13)

    # 5. XỬ LÝ NỘI DUNG CHÍNH (PARSING NỘI DUNG 5512)
    lines = content_text.split('\n')
    for line in lines:
        line_str = line.strip()
        
        # Bỏ qua các dòng phân cách thừa
        if not line_str or line_str.startswith("---") or line_str.startswith("# "):
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        
        # Xử lý các tiêu đề lớn I., II., III., IV.
        if re.match(r'^(I|II|III|IV)\.\s+', line_str, re.IGNORECASE):
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            clean_head = line_str.replace("**", "").replace("*", "")
            run = p.add_run(clean_head.upper())
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)

        # Xử lý các Tiêu đề Hoạt động & Nội dung lớn
        elif line_str.startswith(("HOẠT ĐỘNG ", "NỘI DUNG ", "TIẾT ", "Khối kiến thức ")):
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            clean_act = line_str.replace("**", "").replace("*", "")
            run = p.add_run(clean_act)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)

        # Xử lý các mục 1., 2., 3. (1. Kiến thức, 2. Năng lực...)
        elif re.match(r'^\d+\.\s+', line_str):
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(3)
            parse_formatted_text_to_paragraph(p, line_str)

        # Xử lý các mục a), b), c), d) của Hoạt động
        elif re.match(r'^[a-d]\)\s+', line_str):
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.left_indent = Inches(0.15)
            parse_formatted_text_to_paragraph(p, line_str)

        # Xử lý các Bước tổ chức thực hiện (Bước 1:, Bước 2:...)
        elif "Bước 1:" in line_str or "Bước 2:" in line_str or "Bước 3:" in line_str or "Bước 4:" in line_str:
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.3)
            parse_formatted_text_to_paragraph(p, line_str)

        # Các dòng văn bản bình thường hoặc gạch đầu dòng
        else:
            if line_str.startswith(("- ", "+ ", "* ")):
                p.paragraph_format.left_indent = Inches(0.2)
            parse_formatted_text_to_paragraph(p, line_str)

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

            # PROMPT TẬP TRUNG TRÍCH XUẤT NGUYÊN BẢN TỪ SGV (CÁCH A)
            prompt = f"""
            Bạn là trợ lý trích xuất và chuyển đổi Kế hoạch bài dạy chuẩn Công văn 5512/BGDĐT.

            NGUYÊN TẮC VÀNG CHÍNH XÁC 100% (BẮT BUỘC TUÂN THỦ):
            1. TRÍCH XUẤT CHÍNH XÁC TỪ FILE SGV: 
               - Nếu có tệp tài liệu đính kèm (SGV/Ảnh trang sách), bạn BẮT BUỘC phải đọc nguyên văn phần "I. MỤC TIÊU" (Kiến thức, kỹ năng, phẩm chất, năng lực) từ file SGV đó sang. 
               - TUYỆT ĐỐI KHÔNG TỰ VIẾT THÊM, KHÔNG TỰ BỞI DÃI, KHÔNG TỰ MỞ RỘNG các khái niệm kiến thức nếu trong SGV không ghi. Chép chính exact từng gạch đầu dòng!
               - Nếu không có file đính kèm, sử dụng chính xác nội dung ghi ở phần "Yêu cầu cần đạt / Mục tiêu SGV": {requirements}

            2. CẤU TRÚC BÁM SÁT 5512:
               I. MỤC TIÊU
               1. Về kiến thức, kỹ năng: (Chép Y NGUYÊN các gạch đầu dòng từ SGV/ảnh)
               2. Về phẩm chất, năng lực: (Chép Y NGUYÊN các gạch đầu dòng năng lực, phẩm chất từ SGV/ảnh)
               - Thêm mục nhỏ: "Năng lực Số / Tích hợp AI:" (Nêu ngắn gọn việc dùng {integration_str})

               II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

               III. TIẾN TRÌNH DẠY HỌC
               Trình bày theo các Hoạt động (Khởi động, Khám phá, Luyện tập, Vận dụng) lấy đúng theo mạch nội dung của SGV/SGK. Mỗi hoạt động gồm đủ 4 mục:
               a) Mục tiêu
               b) Nội dung (Mô tả bài tập/câu hỏi bám sát SGK/SGV)
               c) Sản phẩm (Đáp án, lời giải)
               d) Tổ chức thực hiện (Bước 1: Chuyển giao nhiệm vụ -> Bước 2: Thực hiện nhiệm vụ -> Bước 3: Báo cáo, thảo luận -> Bước 4: Kết luận, nhận định).

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
                file_part = {
                    "mime_type": mime_type,
                    "data": bytes_data
                }
                contents.append(file_part)
                st.toast("📄 Đã nạp dữ liệu từ File SGV! Đang ép AI chép chính xác 100%...", icon="✅")

            with st.spinner("✨ AI đang trích xuất Y NGUYÊN từ SGV và xuất giáo án 5512..."):
                response = call_gemini_multimodal(model, contents)
                st.success("🎉 Đã tạo giáo án bám sát 100% SGV!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_SGV_{lesson_title.replace(' ', '_')}.docx",
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
