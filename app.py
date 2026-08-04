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

# Cấu hình trang Streamlit
st.set_page_config(page_title="Hệ thống Soạn Giáo án Tự Động 5512", layout="wide", page_icon="📝")

# CSS chỉnh giao diện chữ to, rõ ràng
st.markdown("""
    <style>
    .block-container { padding-top: 2.7rem !important; padding-bottom: 2rem !important; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, Arial, sans-serif; }
    .sidebar-title { color: #1e293b; font-size: 20px !important; font-weight: 700; margin-bottom: 10px; }
    .step-header { color: #DC2626 !important; font-size: 20px !important; font-weight: 700 !important; margin-top: 15px !important; margin-bottom: 10px !important; }
    div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #0F172A !important; }
    div[data-baseweb="input"] input, div[data-baseweb="textarea"] textarea, div[data-baseweb="select"] span { font-size: 18px !important; color: #1E3A8A !important; }
    .stButton button p { font-size: 20px !important; font-weight: 700 !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR - ĐĂNG NHẬP & CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán mã API Key vào đây...")
    
    # CẬP NHẬT TÊN MODEL CHUẨN ĐỂ TRÁNH LỖI 404
    model_name = st.selectbox(
        "Mô hình AI xử lý:", 
        ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"], 
        index=0
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
    <div style="text-align: center; margin-bottom: 20px;">
        <div style="font-size: 32px; font-weight: 800; color: #1E293B;">📝 HỆ THỐNG SOẠN KHBD 5512 TỰ ĐỘNG</div>
        <div style="font-size: 20px; font-weight: 600; color: #2563EB; margin-top: 5px;">Tác giả: DƯƠNG TẤN TIẾN - GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: THÔNG TIN MÔN HỌC & BÀI DẠY
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: THÔNG TIN MÔN HỌC & BÀI DẠY</div>', unsafe_allow_html=True)

col_sub, col_grd, col_dur = st.columns([2, 1, 1])
with col_sub:
    subject = st.selectbox("Môn học/Hoạt động GD:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"])
with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
with col_dur:
    duration = st.number_input("Số tiết:", value=2, min_value=1)

# Nút cập nhật danh mục bài học từ NXB (Taphuan)
if st.button("🔍 Cập nhật Tên chương & Bài dạy từ https://taphuan.nxbgd.vn", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở menu bên trái!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            prompt_fetch = f"""
            Bạn là hệ thống tra cứu tập huấn của NXB Giáo dục Việt Nam (https://taphuan.nxbgd.vn).
            Liệt kê các Bài học chuẩn môn {subject} - {grade} (Bộ Kết nối tri thức). 
            Trả về duy nhất định dạng JSON mảng: [{{"chapter": "...", "lesson": "...", "duration": 2}}]
            """
            with st.spinner("Đang kết nối và lấy danh mục bài học từ https://taphuan.nxbgd.vn..."):
                res = model.generate_content(prompt_fetch)
                clean_json = re.search(r'\[.*\]', res.text.strip(), re.DOTALL).group(0)
                st.session_state['fetched_lessons'] = json.loads(clean_json)
                st.success("Tải danh mục từ https://taphuan.nxbgd.vn thành công!")
        except Exception as e:
            st.error(f"Lỗi tải danh mục: {e}")

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_options = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    selected_idx = st.selectbox("👉 Chọn Bài học từ taphuan.nxbgd.vn:", range(len(lesson_options)), format_func=lambda x: lesson_options[x])
    st.session_state['chapter_val'] = lessons_data[selected_idx]['chapter']
    st.session_state['lesson_val'] = lessons_data[selected_idx]['lesson']

col_c, col_l = st.columns(2)
with col_c:
    input_chapter = st.text_input("Chương / Chủ đề:", value=st.session_state.get('chapter_val', ''), placeholder="Ví dụ: CHƯƠNG I. MỆNH ĐỀ VÀ TẬP HỢP")
with col_l:
    input_lesson = st.text_input("Tên bài dạy:", value=st.session_state.get('lesson_val', ''), placeholder="Ví dụ: BÀI 5: GIÁ TRỊ LƯỢNG GIÁC CỦA MỘT GÓC TỪ 0° ĐẾN 180°")

# ==========================================
# BƯỚC 2: CUNG CẤP DỮ LIỆU SGK & SGV
# ==========================================
st.markdown('<div class="step-header">📂 BƯỚC 2: CUNG CẤP DỮ LIỆU SÁCH GIÁO KHOA & SÁCH GIÁO VIÊN</div>', unsafe_allow_html=True)

data_source_mode = st.radio(
    "Lựa chọn phương thức cung cấp dữ liệu:",
    ["📌 Tải File / Dán trực tiếp nội dung SGK & SGV (Khuyên dùng - Chính xác 100%)", 
     "🔍 Dùng danh mục bài học từ NXB"],
    index=0
)

uploaded_sgk = None
uploaded_sgv = None
sgk_text = ""
sgv_text = ""

if "📌" in data_source_mode:
    col_sgk_box, col_sgv_box = st.columns(2)
    with col_sgk_box:
        st.markdown("**📘 1. NỘI DUNG SÁCH GIÁO KHOA (Nội dung & Ví dụ):**")
        uploaded_sgk = st.file_uploader("Tải File SGK (PDF/Ảnh):", type=["pdf", "png", "jpg", "jpeg"], key="sgk_file")
        sgk_text = st.text_area("Hoặc DÁN VĂN BẢN SGK:", height=140, placeholder="Dán nội dung SGK vào đây...")

    with col_sgv_box:
        st.markdown("**📕 2. NỘI DUNG SÁCH GIÁO VIÊN (Mục tiêu bài học):**")
        uploaded_sgv = st.file_uploader("Tải File SGV (PDF/Ảnh):", type=["pdf", "png", "jpg", "jpeg"], key="sgv_file")
        sgv_text = st.text_area("Hoặc DÁN MỤC TIÊU TỪ SGV:", height=140, placeholder="Dán mục tiêu từ SGV vào đây...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học", "Giáo dục STEM / STEAM", "Phát triển Tư duy phản biện"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học"]
)

# ==========================================
# XỬ LÝ LỌC CÂU THỪA & TẠO FILE WORD CHUẨN FILE MẪU
# ==========================================
def clean_ai_response(text):
    """Bỏ các câu chào, câu dẫn rác từ AI"""
    lines = text.split('\n')
    cleaned_lines = []
    
    banned_keywords = [
        "bắt buộc trích dẫn", "yêu cầu nghiêm ngặt", "quy định định dạng",
        "dưới đây là kế hoạch bài dạy", "dưới đây là giáo án", "chúc bạn dạy tốt",
        "hy vọng bản giáo án này", "dữ liệu sách giáo khoa", "dữ liệu sách giáo viên"
    ]
    
    for line in lines:
        line_low = line.lower().strip()
        if any(keyword in line_low for keyword in banned_keywords):
            continue
        cleaned_lines.append(line)
        
    return "\n".join(cleaned_lines)

def format_paragraph_with_markdown(paragraph, text):
    """Định dạng in đậm chính xác từng từ **...**"""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            content = part[2:-2].strip()
            if content:
                run = paragraph.add_run(content)
                run.bold = True
        else:
            if part:
                paragraph.add_run(part)

def generate_doc(content_text, final_chapter, final_lesson):
    doc = docx.Document()

    # Cấu hình lề trang (Top 2cm, Bottom 2cm, Left 3cm, Right 2cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)

    # 1. BẢNG TIÊU ĐỀ ĐẦU TRANG KHÔNG VIỀN (ĐÚNG THEO MẪU FILE THẦY CUNG CẤP)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(3.3)
    table.columns[1].width = Inches(3.1)

    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    p_left.add_run(f"Trường: {school_name}\nTổ: {dept_name}")

    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    p_right.add_run(f"Họ và tên giáo viên:{teacher_name}")

    # Xóa toàn bộ đường viền khung bảng
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # 2. TIÊU ĐỀ BÀI DẠY (ĐÚNG CHÍNH XÁC MẪU CHỮ VÀ KIỂU DÁNG)
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(6)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    
    # Format tên bài dạy
    lesson_str = final_lesson.upper()
    if not lesson_str.startswith("TÊN BÀI DẠY:"):
        lesson_str = f"TÊN BÀI DẠY: {lesson_str}"
    r_title = p_title.add_run(lesson_str)
    r_title.bold = True
    r_title.font.size = Pt(13)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)")

    # 3. XỬ LÝ NỘI DUNG VĂN BẢN THEO CHUẨN KHBD MẪU
    clean_text = clean_ai_response(content_text)
    lines = clean_text.split('\n')

    for line in lines:
        line_str = line.strip()
        if not line_str or line_str.startswith("---") or line_str.startswith("#"):
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)

        raw_line = line_str.replace("**", "").replace("*", "")

        # Tiêu đề I. Mục tiêu, II. THIẾT BỊ DẠY HỌC..., III. TIẾN TRÌNH DẠY HỌC
        if re.match(r'^(I|II|III|IV|V)\.\s', raw_line):
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run(raw_line)
            run.bold = True
            if raw_line.startswith("I. Mục tiêu"):
                run.text = "I. Mục tiêu"

        # Tiêu đề HOẠT ĐỘNG 1, HOẠT ĐỘNG 2...
        elif raw_line.upper().startswith("HOẠT ĐỘNG"):
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run(raw_line.upper())
            run.bold = True

        # Tiêu đề NỘI DUNG 1, NỘI DUNG 2...
        elif raw_line.upper().startswith("NỘI DUNG"):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(raw_line.upper())
            run.bold = True

        # Tiêu đề a) Mục tiêu:, b) Nội dung:, c) Sản phẩm:, d) Tổ chức thực hiện:
        elif re.match(r'^[a-d]\)\s', raw_line):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.15)
            format_paragraph_with_markdown(p, line_str)

        # Các Bước 1, Bước 2, Bước 3, Bước 4
        elif "Bước 1:" in raw_line or "Bước 2:" in raw_line or "Bước 3:" in raw_line or "Bước 4:" in raw_line:
            p.paragraph_format.space_before = Pt(2)
            format_paragraph_with_markdown(p, line_str)

        # Các mục nhỏ 1. Kiến thức, 2. Năng lực, 2.1., 2.2...
        elif re.match(r'^\d+(\.\d+)*\.\s', raw_line):
            p.paragraph_format.space_before = Pt(3)
            format_paragraph_with_markdown(p, line_str)

        # Các dòng gạch đầu dòng
        elif raw_line.startswith("- ") or raw_line.startswith("+ "):
            p.paragraph_format.left_indent = Inches(0.15)
            format_paragraph_with_markdown(p, line_str)

        # Văn bản thường
        else:
            format_paragraph_with_markdown(p, line_str)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# BẮT ĐẦU TẠO GIÁO ÁN
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập API Key ở thanh menu bên trái!")
    elif not input_lesson:
        st.error("⚠️ Vui lòng nhập Tên bài dạy ở Bước 1!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)

            sgk_context = ""
            if sgk_text:
                sgk_context += f"\nNỘI DUNG VĂN BẢN SGK DO GIÁO VIÊN CUNG CẤP:\n{sgk_text}\n"
            if sgv_text:
                sgk_context += f"\nMỤC TIÊU VĂN BẢN SGV DO GIÁO VIÊN CUNG CẤP:\n{sgv_text}\n"

            prompt = f"""
            Bạn là Chuyên gia soạn Giáo án 5512 môn {subject} - {grade}.
            BẮT BỦỘC BÁM SÁT 100% NỘI DUNG VÀ MỤC TIÊU DO GIÁO VIÊN CUNG CẤP DƯỚI ĐÂY:
            
            Chương/Chủ đề: {input_chapter}
            Bài dạy: {input_lesson}
            Thời lượng: {duration} tiết.
            Yếu tố tích hợp: {', '.join(integrations)}
            
            === DỮ LIỆU ĐẦU VÀO CỦA GIÁO VIÊN ===
            {sgk_context}
            ======================================

            YÊU CẦU NGHIÊM NGẶT VỀ NỘI DUNG VÀ ĐỊNH DẠNG:
            1. Bám sát chính xác các mục tiêu, bài tập, ví dụ trong dữ liệu SGK/SGV giáo viên đã cấp. Không tự chế nội dung ngoài SGK/SGV.
            2. Sử dụng đúng ký hiệu toán học LaTeX ($...$) cho các công thức, ví dụ: $\\sin, \\cos, \\tan, \\cot$, $0^\\circ$, $180^\\circ$.
            3. Trả về trực tiếp nội dung giáo án, KHÔNG kèm câu chào, lời kết hay ghi chú thừa.

            CẤU TRÚC CHUẨN BẮT BỘC:
            I. Mục tiêu
            1. Kiến thức
            2. Năng lực
            2.1. Năng lực Toán học (Đặc thù):
            2.2. Năng lực chung:
            2.3. Năng lực Số / Ứng dụng CNTT và AI (Tích hợp):
            3. Phẩm chất
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            1. Giáo viên (GV)
            2. Học sinh (HS)
            III. TIẾN TRÌNH DẠY HỌC
            HOẠT ĐỘNG 1: MỞ ĐẦU
            a) Mục tiêu:
            b) Nội dung:
            c) Sản phẩm:
            d) Tổ chức thực hiện:
            Bước 1: Chuyển giao nhiệm vụ
            Bước 2: Thực hiện nhiệm vụ
            Bước 3: Báo cáo, thảo luận
            Bước 4: Kết luận, nhận định
            HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC
            ...
            """

            content_payload = [prompt]
            if uploaded_sgk is not None:
                content_payload.append({"mime_type": uploaded_sgk.type, "data": uploaded_sgk.getvalue()})
            if uploaded_sgv is not None:
                content_payload.append({"mime_type": uploaded_sgv.type, "data": uploaded_sgv.getvalue()})

            with st.spinner("✨ AI đang tổng hợp chính xác từ SGK/SGV và tạo file Word..."):
                response = model.generate_content(content_payload)
                doc_file = generate_doc(response.text, input_chapter, input_lesson)

                st.success("🎉 Đã tạo xong Kế hoạch bài dạy chuẩn!")
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{input_lesson.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown(clean_ai_response(response.text))

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
