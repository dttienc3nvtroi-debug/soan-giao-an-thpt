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

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn KNTT 2026)", 
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
# CƠ SỞ DỮ LIỆU MẶC ĐỊNH (KẾT NỐI TRI THỨC VỚI CUỘC SỐNG)
# ==========================================
BUILTIN_LESSONS = {
    "Toán học_Lớp 10": [
        {"chapter": "Chương I: Mệnh đề và Tập hợp", "lesson": "Bài 1: Mệnh đề", "duration": 2, "req": "- Thiết lập và phát biểu được mệnh đề, mệnh đề phủ định, mệnh đề kéo theo, mệnh đề tương đương.\n- Xác định tính đúng/sai của mệnh đề.\n- Sử dụng đúng ký hiệu ∀, ∃."},
        {"chapter": "Chương I: Mệnh đề và Tập hợp", "lesson": "Bài 2: Tập hợp và các phép toán trên tập hợp", "duration": 3, "req": "- Thực hiện thành thạo các phép toán hợp, giao, hiệu của hai tập hợp."}
    ],
    "Toán học_Lớp 11": [
        {"chapter": "Chương I: Hàm số lượng giác và Phương trình lượng giác", "lesson": "Bài 1: Góc lượng giác", "duration": 2, "req": "- Nhận biết khái niệm góc lượng giác, đường tròn lượng giác.\n- Đổi đơn vị từ độ sang radian và ngược lại."}
    ],
    "Toán học_Lớp 12": [
        {"chapter": "Chương I: Ứng dụng đạo hàm để khảo sát hàm số", "lesson": "Bài 1: Tính đơn điệu và cực trị của hàm số", "duration": 3, "req": "- Xét tính đơn điệu và tìm cực trị của hàm số dựa vào dấu của đạo hàm bậc nhất."}
    ],
    "Tiếng Anh_Lớp 10": [
        {"chapter": "Unit 1: Family Life", "lesson": "Getting Started - Household Chores", "duration": 1, "req": "- Use words and phrases related to household chores and family life.\n- Identify and practice /br/, /kr/, and /tr/."}
    ],
    "Tiếng Anh_Lớp 11": [
        {"chapter": "Unit 1: A Long and Healthy Life", "lesson": "Getting Started - Healthy habits", "duration": 1, "req": "- Use words/phrases related to health and fitness.\n- Pronounce strong and weak forms of auxiliary verbs."}
    ],
    "Tiếng Anh_Lớp 12": [
        {"chapter": "Unit 1: Life Stories", "lesson": "Getting Started - Historical figures", "duration": 1, "req": "- Use vocabulary related to achievements and life stories of famous people."}
    ],
    "Ngữ văn_Lớp 10": [
        {"chapter": "Bài 1: Sức hấp dẫn của truyện kể", "lesson": "Văn bản 1: Truyện về các vị thần sáng tạo thế giới (Thần thoại VN)", "duration": 2, "req": "- Phân tích các yếu tố không gian, thời gian, cốt truyện, nhân vật thần thoại."}
    ],
    "Ngữ văn_Lớp 11": [
        {"chapter": "Bài 1: Khai phá thế giới kì ảo", "lesson": "Văn bản 1: Truyện thơ dân gian", "duration": 2, "req": "- Phân tích đặc điểm truyện thơ, nhân vật, cốt truyện và thông điệp."}
    ],
    "Ngữ văn_Lớp 12": [
        {"chapter": "Bài 1: Khả năng lớn lao của tiểu thuyết", "lesson": "Văn bản 1: Nhìn về vốn văn hóa dân tộc", "duration": 2, "req": "- Phân tích tư tưởng, nhận thức luận và nghệ thuật lập luận trong văn bản nghị luận."}
    ]
}

# ==========================================
# CƠ CHẾ KHẮC PHỤC LỖI 404 & GỌI MODEL TỰ ĐỘNG
# ==========================================
def get_available_models(api_key):
    """Lấy danh sách các model thực sự hỗ trợ generateContent từ API Key"""
    if not api_key:
        return []
    try:
        genai.configure(api_key=api_key.strip())
        valid_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                valid_models.append(m.name)
        return valid_models
    except Exception:
        return []

def call_gemini_safe(api_key, target_model_name, contents, max_retries=3):
    genai.configure(api_key=api_key.strip())
    
    # Chuẩn hóa tên model
    clean_name = target_model_name.replace("models/", "").strip() if target_model_name else "gemini-1.5-flash"
    
    # Lập danh sách ưu tiên fallback nếu model chọn bị lỗi 404
    fallback_candidates = [
        clean_name,
        f"models/{clean_name}",
        "gemini-1.5-flash-latest",
        "gemini-1.5-pro-latest",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ]
    
    last_error = None
    for model_code in fallback_candidates:
        try:
            model = genai.GenerativeModel(model_code)
            return model.generate_content(contents)
        except Exception as e:
            last_error = e
            if "404" in str(e) or "not found" in str(e):
                continue  # Thử model tiếp theo trong danh sách fallback
            elif "429" in str(e) or "Quota" in str(e):
                time.sleep(3)
                continue
            else:
                raise e
    
    raise last_error

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key vào đây...")
    
    # Tự động lấy danh sách model thực tế từ API Key để tránh lỗi 404
    available_models = get_available_models(api_key) if api_key else []
    
    if available_models:
        model_name = st.selectbox("Mô hình AI khả dụng:", available_models, index=0)
    else:
        model_name = st.selectbox("Mô hình AI mặc định:", ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"], index=0)

    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ CHUYÊN MÔN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 25px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 22px; border-radius: 12px; border: 1px solid #bfdbfe;">
        <div style="font-size: 33px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN KNTT 2026 (5512)
        </div>
        <div style="font-size: 21px; font-weight: 600; color: #2563eb; margin-top: 10px;">
            📝 Tác giả: DƯƠNG TẤN TIẾN — GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP (BỘ KẾT NỐI TRI THỨC VỚI CUỘC SỐNG)</div>', unsafe_allow_html=True)
col_sub, col_grd = st.columns(2)

ALL_12_SUBJECTS = [
    "Toán học", "Ngữ văn", "Tiếng Anh", 
    "Vật lý", "Hóa học", "Sinh học", 
    "Lịch sử", "Địa lý", "GDKT&PL", 
    "Tin học", "Công nghệ", "HĐTN, Hướng nghiệp"
]

with col_sub:
    st.markdown('<span class="custom-label">Môn học/Hoạt động GD:</span>', unsafe_allow_html=True)
    subject = st.selectbox("Môn học:", ALL_12_SUBJECTS, label_visibility="collapsed")
with col_grd:
    st.markdown('<span class="custom-label">Khối lớp:</span>', unsafe_allow_html=True)
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"], label_visibility="collapsed")

# ==========================================
# BƯỚC 2: QUẢN LÝ DANH MỤC BÀI HỌC VÀ SGV
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: DANH MỤC BÀI HỌC VÀ FILE SGV (TAPHUAN.NXBGD.VN)</div>', unsafe_allow_html=True)

key_sub_grade = f"{subject}_{grade}"

default_fallback = [
    {
        "chapter": f"Chương I / Chủ đề 1 môn {subject} ({grade})",
        "lesson": f"Bài 1: Bài học mở đầu môn {subject} {grade}",
        "duration": 2,
        "req": f"- Nắm vững các kiến thức trọng tâm môn {subject} {grade} theo bộ sách Kết nối tri thức với cuộc sống.\n- Vận dụng giải quyết các bài tập thực hành theo YCĐ của Bộ GD&ĐT."
    }
]

default_lessons = BUILTIN_LESSONS.get(key_sub_grade, default_fallback)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown(f'<span class="custom-label">🌐 Tải trọn bộ danh mục bài học chuẩn SGV 2026:</span>', unsafe_allow_html=True)
    if st.button(f"🔄 Đồng bộ đầy đủ tất cả bài học {subject} - {grade}", use_container_width=True):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                prompt_fetch = f"""
                Bạn là CSDL chuẩn Bộ sách "Kết nối tri thức với cuộc sống" (NXB Giáo dục Việt Nam).
                Hãy xuất danh sách TẤT CẢ các Bài học chính thức thuộc SGK môn {subject} - {grade}.
                Trả về duy nhất định dạng JSON Mảng đối tượng không chứa markdown:
                [
                  {{
                    "chapter": "Tên Chương / Chủ đề / Unit",
                    "lesson": "Tên Bài học / Lesson",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt chi tiết của bài theo chuẩn SGV"
                  }}
                ]
                """
                with st.spinner(f"✨ Đang kết nối CSDL NXB Giáo Dục cho môn {subject} {grade}..."):
                    res = call_gemini_safe(clean_api_key, model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    parsed_lessons = json.loads(clean_json)
                    if isinstance(parsed_lessons, list) and len(parsed_lessons) > 0:
                        st.session_state[f'fetched_{key_sub_grade}'] = parsed_lessons
                        st.success(f"🎉 Đã nạp thành công {len(parsed_lessons)} bài học môn {subject} {grade}!")
            except Exception as e:
                st.error(f"❌ Không thể đồng bộ: {str(e)}")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tải lên trang SGV (PDF/Ảnh chụp taphuan.nxbgd.vn):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Tải lên File SGV:", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

if uploaded_sgv_file is not None:
    st.info(f"✅ Đã nhận file: **{uploaded_sgv_file.name}**. AI sẽ trích xuất Y NGUYÊN mục tiêu từ file này!")

active_lessons = st.session_state.get(f'fetched_{key_sub_grade}', default_lessons)

lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in active_lessons]
st.markdown(f'<span class="custom-label">👉 Chọn Bài học môn {subject} ({grade}):</span>', unsafe_allow_html=True)
selected_idx = st.selectbox("Chọn bài:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
current_item = active_lessons[selected_idx]

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

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ (BỔ SUNG VÀO CẤU TRÚC 5512)</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Geogebra, Padlet, Kahoot, Quizizz...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Geogebra, Padlet, Kahoot, Quizizz...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    label_visibility="collapsed"
)

# ==========================================
# HELPER PARSER & XUẤT WORD CHUẨN 5512
# ==========================================
def clean_latex_math(text):
    text = re.sub(r'\\circ', '°', text)
    text = text.replace('$', '')
    return text

def parse_formatted_text_to_paragraph(paragraph, raw_text):
    raw_text = clean_latex_math(raw_text)
    pattern = r'(\*\*.*?\*\*|\*.*?\*)'
    tokens = re.split(pattern, raw_text)
    for token in tokens:
        if not token:
            continue
        if token.startswith('**') and token.endswith('**'):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith('*') and token.endswith('*'):
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        else:
            run = paragraph.add_run(token)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

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
    table.autofit = False
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    p_left = table.cell(0, 0).paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    p_left.add_run(f"TRƯỜNG: {school_name.upper()}\nTỔ: {dept_name.upper()}").bold = True

    p_right = table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    p_right.add_run(f"Họ và tên giáo viên:\n{teacher_name}").bold = True

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)")
    r_sub.italic = True

    for line in content_text.split('\n'):
        line_str = line.strip()
        if not line_str or line_str.startswith("---") or line_str.startswith("# "):
            continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        
        if re.match(r'^(I|II|III|IV)\.\s+', line_str, re.IGNORECASE):
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(line_str.replace("**", "").replace("*", "").upper())
            run.bold = True
            run.font.size = Pt(14)
        elif line_str.startswith(("HOẠT ĐỘNG ", "NỘI DUNG ", "TIẾT ")):
            p.paragraph_format.space_before = Pt(10)
            run = p.add_run(line_str.replace("**", "").replace("*", ""))
            run.bold = True
        else:
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
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Bạn là trợ lý soạn Kế hoạch bài dạy chuẩn Công văn 5512/BGDĐT từ tài liệu SGK Kết nối tri thức với cuộc sống (taphuan.nxbgd.vn).

            NGUYÊN TẮC TUÂN THỦ:
            1. Nếu có tệp SGV/Ảnh kèm theo, trích xuất CHÍNH XÁC Y NGUYÊN phần Mục tiêu từ SGV sang.
            2. Nếu không có tệp đính kèm, sử dụng đúng Yêu cầu cần đạt: {requirements}

            CẤU TRÚC BÁM SÁT 5512:
            I. MỤC TIÊU (Về kiến thức, kỹ năng, năng lực, phẩm chất, Năng lực số: {integration_str})
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Khởi động, Khám phá, Luyện tập, Vận dụng). Mỗi hoạt động đủ 4 bước: a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện.

            THÔNG TIN:
            - Môn: {subject} ({grade})
            - Bài dạy: {lesson_title} ({duration} tiết)
            """

            contents = [prompt]
            if uploaded_sgv_file is not None:
                contents.append({"mime_type": uploaded_sgv_file.type, "data": uploaded_sgv_file.getvalue()})

            with st.spinner(f"✨ AI đang trích xuất SGV môn {subject} và tạo file Word..."):
                response = call_gemini_safe(clean_api_key, model_name, contents)
                st.success("🎉 Đã tạo giáo án bám sát SGV!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{subject}_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            st.error(f"❌ Lỗi khi sinh giáo án: `{str(e)}`")
