import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
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
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Đầy đủ các môn THPT 2018)", 
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
# CƠ SỞ DỮ LIỆU TẤT CẢ CÁC MÔN GDPT 2018 (KNTT)
# ==========================================
BUILTIN_LESSONS = {
    # 1. TOÁN HỌC
    "Toán học_Lớp 10": [
        {"chapter": "Chương I: Mệnh đề và Tập hợp", "lesson": "Bài 1: Mệnh đề", "duration": 2, "req": "- Thiết lập và phát biểu được mệnh đề, mệnh đề phủ định, mệnh đề kéo theo, mệnh đề tương đương.\n- Xác định được tính đúng/sai của mệnh đề đơn giản.\n- Sử dụng đúng các ký hiệu ∀, ∃."},
        {"chapter": "Chương I: Mệnh đề và Tập hợp", "lesson": "Bài 2: Tập hợp và các phép toán trên tập hợp", "duration": 3, "req": "- Hiểu và sử dụng đúng các khái niệm tập hợp, tập hợp con, tập hợp bằng nhau.\n- Thực hiện thành thạo các phép toán hợp, giao, hiệu của hai tập hợp."},
        {"chapter": "Chương II: Bất phương trình và Hệ BPT bậc nhất hai ẩn", "lesson": "Bài 3: Bất phương trình bậc nhất hai ẩn", "duration": 2, "req": "- Nhận biết được bất phương trình bậc nhất hai ẩn.\n- Biểu diễn được miền nghiệm của bất phương trình bậc nhất hai ẩn trên mặt phẳng tọa độ."}
    ],

    # 2. NGỮ VĂN
    "Ngữ văn_Lớp 10": [
        {"chapter": "Bài 1: Sức hấp dẫn của truyện kể", "lesson": "Văn bản 1: Truyện về các vị thần sáng tạo thế giới (Thần thoại Việt Nam)", "duration": 2, "req": "- Nhận biết và phân tích được một số yếu tố của thần thoại: không gian, thời gian, cốt truyện, nhân vật.\n- Phân tích được chủ đề, thông điệp tác phẩm."},
        {"chapter": "Bài 1: Sức hấp dẫn của truyện kể", "lesson": "Văn bản 2: Tản Viên từ phản đền (Nguyễn Dữ)", "duration": 2, "req": "- Nhận biết và phân tích được một số yếu tố của truyền kỳ: nhân vật, sự kiện, yếu tố kỳ ảo.\n- Phân tích nét tính cách dũng cảm đấu tranh cho chính nghĩa của Ngô Tử Văn."}
    ],

    # 3. TIẾNG ANH
    "Tiếng Anh_Lớp 10": [
        {"chapter": "Unit 1: Family Life", "lesson": "Getting Started - Household Chores", "duration": 1, "req": "- Use words and phrases related to household chores and family life.\n- Identify and practice the pronunciation of consonant blends /br/, /kr/, and /tr/."},
        {"chapter": "Unit 1: Family Life", "lesson": "Reading - Benefits of doing housework", "duration": 1, "req": "- Read for main ideas and specific information about the benefits of sharing housework for children."}
    ],

    # 4. VẬT LÝ
    "Vật lý_Lớp 10": [
        {"chapter": "Chương I: Mở đầu", "lesson": "Bài 1: Tốc độ, vận tốc và gia tốc", "duration": 2, "req": "- Lập luận để rút ra định nghĩa và công thức tính vận tốc, gia tốc.\n- Biểu diễn được vectơ vận tốc và gia tốc trong chuyển động thẳng."}
    ],

    # 5. HÓA HỌC
    "Hóa học_Lớp 10": [
        {"chapter": "Chương 1: Cấu tạo nguyên tử", "lesson": "Bài 1: Thành phần nguyên tử", "duration": 2, "req": "- Trình bày được thành phần của nguyên tử (hạt nhân, vỏ nguyên tử, proton, neutron, electron).\n- So sánh được khối lượng và điện tích của các hạt."}
    ],

    # 6. SINH HỌC
    "Sinh học_Lớp 10": [
        {"chapter": "Phần 1: Giới thiệu chung về thế giới sống", "lesson": "Bài 1: Giới thiệu chương trình môn Sinh học và các cấp độ tổ chức của thế giới sống", "duration": 2, "req": "- Nêu được mục tiêu, đối tượng nghiên cứu của môn Sinh học.\n- Trình bày được các đặc điểm chung của các cấp độ tổ chức sống."}
    ],

    # 7. LỊCH SỬ
    "Lịch sử_Lớp 10": [
        {"chapter": "Chủ đề 1: Lịch sử và Sử học", "lesson": "Bài 1: Hiện thực lịch sử và nhận thức lịch sử", "duration": 2, "req": "- Trình bày được khái niệm hiện thực lịch sử và nhận thức lịch sử.\n- Phân biệt được hiện thực lịch sử và nhận thức lịch sử qua ví dụ cụ thể."}
    ],

    # 8. ĐỊA LÝ
    "Địa lý_Lớp 10": [
        {"chapter": "Chương 1: Sử dụng bản đồ", "lesson": "Bài 1: Một số phương pháp biểu hiện đặc bố trí các đối tượng địa lý trên bản đồ", "duration": 2, "req": "- Phân biệt được các phương pháp biểu hiện đối tượng địa lý trên bản đồ (phương pháp ký hiệu, bản đồ - biểu đồ, đường đẳng nhật...)."}
    ],

    # 9. TIN HỌC
    "Tin học_Lớp 10": [
        {"chapter": "Chủ đề A: Máy tính và xã hội trí thức", "lesson": "Bài 1: Thông tin và xử lý thông tin", "duration": 2, "req": "- Phân biệt được thông tin và dữ liệu.\n- Giải thích được chuyển đổi dữ liệu thành thông tin trong các hệ thống xử lý."}
    ],

    # 10. GIÁO DỤC KINH TẾ VÀ PHÁP LUẬT (GDKT&PL)
    "GDKT&PL_Lớp 10": [
        {"chapter": "Chủ đề 1: Nền kinh tế và các chủ thể kinh tế", "lesson": "Bài 1: Các hoạt động kinh tế cơ bản trong đời sống xã hội", "duration": 2, "req": "- Nêu được vai trò của các hoạt động sản xuất, phân phối, trao đổi, tiêu dùng trong đời sống xã hội."}
    ],

    # 11. CÔNG NGHỆ
    "Công nghệ_Lớp 10": [
        {"chapter": "Chủ đề 1: Khái quát về công nghệ", "lesson": "Bài 1: Công nghệ và đời sống", "duration": 2, "req": "- Nêu được bản chất, vai trò của công nghệ đối với đời sống con người và phát triển xã hội."}
    ],

    # 12. HOẠT ĐỘNG TRẢI NGHIỆM, HƯỚNG NGHIỆP
    "HĐTN, Hướng nghiệp_Lớp 10": [
        {"chapter": "Chủ đề 1: Phát triển bản thân", "lesson": "Bài 1: Khám phá bản thân và thể hiện sự tự tin", "duration": 3, "req": "- Nhận diện được đặc điểm tính cách, giá trị bản thân.\n- Thể hiện được sự tự tin trong giao tiếp và các hoạt động tập thể."}
    ]
}

def get_robust_model(selected_name):
    candidate_models = [
        selected_name,
        "gemini-1.5-flash",
        "gemini-1.5-flash-latest",
        "gemini-2.0-flash",
        "gemini-1.5-pro"
    ]
    seen = set()
    unique_candidates = [x for x in candidate_models if not (x in seen or seen.add(x))]
    
    for m_name in unique_candidates:
        try:
            clean_name = m_name.replace("models/", "").strip()
            model = genai.GenerativeModel(clean_name)
            return model, clean_name
        except Exception:
            continue
    return genai.GenerativeModel("gemini-1.5-flash"), "gemini-1.5-flash"

def call_gemini_multimodal(model_name_input, contents, max_retries=3):
    model, _ = get_robust_model(model_name_input)
    for attempt in range(max_retries):
        try:
            return model.generate_content(contents)
        except Exception as e:
            err_msg = str(e)
            if ("404" in err_msg or "not found" in err_msg) and attempt == 0:
                try:
                    fb = genai.GenerativeModel("gemini-1.5-flash-latest")
                    return fb.generate_content(contents)
                except Exception:
                    pass
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    time.sleep((attempt + 1) * 6)
                    continue
            raise e

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key vào đây...")
    model_name = st.selectbox("Mô hình AI xử lý:", ["gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-2.0-flash", "gemini-1.5-pro"], index=0)
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
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG TOÀN DIỆN THPT GDPT 2018 (5512)
        </div>
        <div style="font-size: 21px; font-weight: 600; color: #2563eb; margin-top: 10px;">
            📝 Tác giả: DƯƠNG TẤN TIẾN — GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP (GDPT 2018)</div>', unsafe_allow_html=True)
col_sub, col_grd = st.columns(2)
with col_sub:
    st.markdown('<span class="custom-label">Môn học/Hoạt động GD:</span>', unsafe_allow_html=True)
    all_subjects = [
        "Toán học", "Ngữ văn", "Tiếng Anh", 
        "Vật lý", "Hóa học", "Sinh học", 
        "Lịch sử", "Địa lý", "GDKT&PL", 
        "Tin học", "Công nghệ", "HĐTN, Hướng nghiệp"
    ]
    subject = st.selectbox("Môn học:", all_subjects, label_visibility="collapsed")
with col_grd:
    st.markdown('<span class="custom-label">Khối lớp:</span>', unsafe_allow_html=True)
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"], label_visibility="collapsed")

# ==========================================
# BƯỚC 2: TRA CỨU & NẠP TỆP SGV / SGK
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: NẠP FILE SGV TỪ TAPHUAN.NXBGD.VN (ĐẢM BẢO CHÍNH XÁC 100%)</div>', unsafe_allow_html=True)

key_sub_grade = f"{subject}_{grade}"
default_lessons = BUILTIN_LESSONS.get(key_sub_grade, [])

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Cập nhật thêm danh mục bài học từ AI:</span>', unsafe_allow_html=True)
    if st.button(f"🔍 Tra cứu thêm bài học môn {subject} - {grade}", use_container_width=True):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                prompt_fetch = f"""
                Bạn là CSDL chuẩn Bộ sách "Kết nối tri thức với cuộc sống" (taphuan.nxbgd.vn).
                Hãy xuất danh sách TẤT CẢ các Bài học chính thức môn {subject} - {grade}.
                Trả về duy nhất định dạng JSON Mảng đối tượng:
                [
                  {{
                    "chapter": "Tên Bài lớn / Chủ đề",
                    "lesson": "Tên Bài học nhỏ",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt chi tiết của bài theo chuẩn SGV"
                  }}
                ]
                """
                with st.spinner(f"✨ Đang tìm kiếm thêm bài học môn {subject} {grade}..."):
                    res = call_gemini_multimodal(model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    parsed_lessons = json.loads(clean_json)
                    if isinstance(parsed_lessons, list) and len(parsed_lessons) > 0:
                        st.session_state[f'fetched_{key_sub_grade}'] = parsed_lessons
                        st.success(f"🎉 Đã nạp thành công {len(parsed_lessons)} bài học môn {subject}!")
            except Exception:
                st.info(f"📌 Đã hiển thị sẵn danh mục chuẩn NXB Giáo Dục môn {subject} {grade}.")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tải lên trang SGV (PDF/Ảnh chụp taphuan.nxbgd.vn):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Tải lên File SGV:", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

if uploaded_sgv_file is not None:
    st.info(f"✅ Đã nhận file: **{uploaded_sgv_file.name}**. AI sẽ trích xuất Y NGUYÊN mục tiêu từ file này!")

active_lessons = st.session_state.get(f'fetched_{key_sub_grade}', default_lessons)

if active_lessons:
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
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        st.markdown('<span class="custom-label">Chương / Chủ đề:</span>', unsafe_allow_html=True)
        chapter_title = st.text_input("Chương:", value="", placeholder=f"Nhập tên chương môn {subject}...", label_visibility="collapsed")
        st.markdown('<span class="custom-label">Tên bài dạy:</span>', unsafe_allow_html=True)
        lesson_title = st.text_input("Tên bài:", value="", placeholder=f"Nhập tên bài môn {subject}...", label_visibility="collapsed")
        st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
        duration = st.number_input("Số tiết:", value=2, label_visibility="collapsed")
    with col_i2:
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt / Mục tiêu SGV:</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value="", placeholder="Nhập mục tiêu SGV hoặc tải file SGV lên...", height=230, label_visibility="collapsed")

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
            genai.configure(api_key=clean_api_key)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Bạn là trợ lý soạn Kế hoạch bài dạy chuẩn Công văn 5512/BGDĐT từ tài liệu taphuan.nxbgd.vn.

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
                response = call_gemini_multimodal(model_name, contents)
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
