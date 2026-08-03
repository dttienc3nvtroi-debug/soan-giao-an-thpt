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

# HÀM XỬ LÝ TỆP ĐÍNH KÈM CHUẨN GEMINI API
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

# HÀM GỌI GEMINI API
def call_gemini(api_key, model_choice, contents):
    genai.configure(api_key=api_key)
    clean_model_name = model_choice.replace("models/", "").strip()
    generation_config = genai.types.GenerationConfig(
        temperature=0.1,
        top_p=0.8,
        top_k=40
    )
    try:
        model = genai.GenerativeModel(model_name=clean_model_name, generation_config=generation_config)
        response = model.generate_content(contents)
        return response
    except Exception as e:
        if clean_model_name != "gemini-2.0-flash":
            model = genai.GenerativeModel(model_name="gemini-2.0-flash", generation_config=generation_config)
            return model.generate_content(contents)
        raise e

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key vào đây...")
    model_name = st.selectbox("Mô hình AI ưu tiên:", ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"], index=0)
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
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
# BƯỚC 2: NHẬP VÀ XÁC NHẬN TÊN BÀI HỌC (CHUẨN 100%)
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TÊN CHƯƠNG & TÊN BÀI DẠY (CHÍNH XÁC NGUYÊN VĂN SGK)</div>', unsafe_allow_html=True)

input_mode = st.radio(
    "Hình thức nhập bài học:",
    ["✍️ Tự nhập / Sửa tên bài trực tiếp (Khuyên dùng - Chuẩn 100%)", "🤖 Tra cứu tự động từ AI", "📂 Đọc từ File/Ảnh SGV"],
    horizontal=True
)

chapter_title = ""
lesson_title = ""
requirements = ""
duration = 3

if "✍️ Tự nhập" in input_mode:
    col_c1, col_c2 = st.columns([1, 1])
    with col_c1:
        chapter_title = st.text_input("Tên Chương / Chủ đề (Nguyên văn SGK):", value="Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị của hàm số")
        lesson_title = st.text_input("Tên Bài dạy (Nguyên văn SGK - Đầy đủ 100%):", value="Bài 1. Tính đơn điệu và cực trị của hàm số")
    with col_c2:
        duration = st.number_input("Số tiết thực hiện:", value=3, min_value=1, max_value=20)
        requirements = st.text_area("Mục tiêu / Yêu cầu cần đạt (Nếu có - Hoặc để trống AI tự tạo):", value="", height=100)

elif "🤖 Tra cứu tự động" in input_mode:
    clean_api_key = api_key.strip() if api_key else ""
    if st.button("🔍 Bấm để AI liệt kê danh sách Bài học"):
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập API Key ở menu bên trái!")
        else:
            with st.spinner("⚡ Đang tra cứu danh mục bài học đầy đủ..."):
                try:
                    p = f"Liệt kê danh sách tất cả các Bài học môn {subject} {grade} ({book_series}). Trả về duy nhất mảng JSON dạng: [{{\"chapter\": \"Chương...\", \"lesson\": \"Bài...\"}}]"
                    res = call_gemini(clean_api_key, model_name, [p])
                    raw = res.text.strip()
                    m = re.search(r'\[.*\]', raw, re.DOTALL)
                    clean_j = m.group(0) if m else raw
                    st.session_state['ai_fetched'] = json.loads(clean_j)
                    st.success("✅ Đã lấy xong danh sách!")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
    
    if 'ai_fetched' in st.session_state:
        items = st.session_state['ai_fetched']
        opts = [f"{x.get('chapter','')} ➔ {x.get('lesson','')}" for x in items]
        idx = st.selectbox("Chọn bài học:", range(len(opts)), format_func=lambda i: opts[i])
        chapter_title = items[idx].get('chapter', '')
        lesson_title = items[idx].get('lesson', '')
        duration = st.number_input("Số tiết thực hiện:", value=3)
        requirements = st.text_area("Yêu cầu cần đạt:", value="", height=80)

else: # Đọc từ file
    uploaded_file = st.file_uploader("Tải lên File/Ảnh trang SGV:", type=["pdf", "png", "jpg", "jpeg"])
    chapter_title = st.text_input("Tên Chương (Kiểm tra lại sau khi AI đọc):", value="")
    lesson_title = st.text_input("Tên Bài dạy đầy đủ (Kiểm tra lại sau khi AI đọc):", value="")
    duration = st.number_input("Số tiết:", value=3)
    requirements = st.text_area("Mục tiêu trích xuất:", value="", height=80)

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ (CẤU TRÚC 5512)</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", "Giáo dục STEM / STEAM", "Phát triển Tư duy phản biện"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)

# ==========================================
# HÀM XUẤT FILE WORD CHUẨN 5512
# ==========================================
def generate_doc(content_text, full_chapter, full_lesson):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)

    # Bảng Header
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

    # Tên Chương - KHÓA CỐ ĐỊNH TỪ Ô NHẬP LIỆU
    if full_chapter:
        p_chap = doc.add_paragraph()
        p_chap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_chap.paragraph_format.space_before = Pt(14)
        p_chap.paragraph_format.space_after = Pt(2)
        r_chap = p_chap.add_run(f"{full_chapter.upper()}")
        r_chap.bold = True
        r_chap.font.size = Pt(13)

    # Tên Bài - KHÓA CỐ ĐỊNH TỪ Ô NHẬP LIỆU (100% KHÔNG BỊ NÓI TÓM TẮT)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {full_lesson.upper()}")
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

# ==========================================
# NÚT BẮT ĐẦU TẠO GIÁO ÁN
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    clean_api_key = api_key.strip() if api_key else ""
    if not clean_api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập hoặc chọn Tên bài dạy!")
    else:
        try:
            integration_str = ", ".join(integrations) if integrations else "Không"
            
            prompt = f"""
            Bạn là Chuyên gia Giáo dục. Hãy biên soạn Kế hoạch bài dạy (Giáo án) theo chuẩn Công văn 5512/BGDĐT.

            MÔN HỌC: {subject} - {grade} ({book_series})
            CHƯƠNG / CHỦ ĐỀ: {chapter_title}
            TÊN BÀI DẠY: {lesson_title}
            SỐ TIẾT: {duration}
            YÊU CẦU CẦN ĐẠT SGV / MỤC TIÊU: {requirements if requirements else 'Tự đề xuất chuẩn SGV'}
            YẾU TỐ TÍCH HỢP: {integration_str}

            YÊU CẦU BẮT BUỘC VỀ CẤU TRÚC GIÁO ÁN:
            I. MỤC TIÊU
            1. Về kiến thức
            2. Về năng lực (Năng lực chung, Năng lực đặc thù và tích hợp: {integration_str})
            3. Về phẩm chất

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            - GV: ...
            - HS: ...

            III. TIẾN TRÌNH DẠY HỌC
            Chi tiết các Hoạt động (Khởi động, Khám phá, Luyện tập, Vận dụng).
            Mỗi Hoạt động PHẢI đủ 4 phần:
            a) Mục tiêu
            b) Nội dung
            c) Sản phẩm
            d) Tổ chức thực hiện (Bước 1: Chuyển giao, Bước 2: Thực hiện, Bước 3: Báo cáo, Bước 4: Kết luận).
            """

            with st.spinner("⚡ AI đang biên soạn nội dung giáo án..."):
                response = call_gemini(clean_api_key, model_name, [prompt])
                st.success("🎉 Đã tạo xong giáo án chuẩn 5512!")
                
                # Ép buộc lấy đúng 100% Tên bài dạy từ ô nhập liệu của Thầy
                doc_file = generate_doc(response.text, full_chapter=chapter_title, full_lesson=lesson_title)
                
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
            st.error(f"❌ Lỗi xử lý: `{e}`. Thầy vui lòng kiểm tra lại API Key!")
