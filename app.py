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
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# CẤU HÌNH GIAO DIỆN & TĂNG KÍCH THƯỚC TRIỆT ĐỂ CHO SELECTBOX
# ==========================================
st.markdown("""
    <style>
    /* 1. Khoảng cách tổng thể trang */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1300px;
    }

    /* 2. Font chữ mặc định hệ thống */
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Roboto, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
    }
    
    /* 3. Tùy chỉnh thanh Sidebar đăng nhập */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }

    .sidebar-title {
        color: #0f172a;
        font-size: 22px !important;
        font-weight: 700;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
        border-bottom: 3px solid #2563eb;
        padding-bottom: 6px;
    }
    
    /* 4. Định dạng Tiêu đề các BƯỚC */
    .step-header {
        color: #dc2626 !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        margin-top: 24px !important;
        margin-bottom: 14px !important;
        padding-left: 10px;
        border-left: 5px solid #dc2626;
    }

    /* 5. Định dạng Nhãn (Labels) cho các Input/Select */
    div[data-testid="stWidgetLabel"] p, .custom-label {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 6px !important;
    }

    /* ========================================================
       6. ĐẶC TRỊ TĂNG FONT CHO 3 Ô SELECTBOX VÀ MULTISELECT
       ======================================================== */
    /* Ép tất cả văn bản bên trong các ô Selectbox (Môn học, Khối lớp, Bài học chuẩn) */
    .stSelectbox div[data-baseweb="select"] *,
    .stSelectbox [data-testid="stMarkdownContainer"] p,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
        line-height: 1.4 !important;
    }

    /* Tăng font danh sách sổ xuống khi click vào Selectbox */
    ul[role="listbox"] li,
    ul[role="listbox"] li * {
        font-size: 21px !important;
        font-weight: 600 !important;
    }

    /* Thẻ Tag Tích hợp năng lực (Multiselect ở Bước 3) */
    span[data-baseweb="tag"] * {
        font-size: 18px !important;
        font-weight: 700 !important;
    }

    /* B. Chữ bên trong ô Input (Chương, Tên bài, Số tiết,...) */
    div[data-baseweb="input"] input,
    .stTextInput input,
    .stNumberInput input {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
        padding-top: 8px !important;
        padding-bottom: 8px !important;
    }

    /* C. Chữ bên trong ô Textarea (Yêu cầu cần đạt) */
    div[data-baseweb="textarea"] textarea,
    .stTextArea textarea {
        font-size: 21px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
        line-height: 1.4 !important;
    }

    /* D. Nút tăng giảm số tiết (+ / -) */
    .stNumberInput button {
        height: 100% !important;
    }
    .stNumberInput button * {
        font-size: 19px !important;
    }

    /* 7. Thiết kế Nút Bấm */
    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease-in-out;
    }
    
    .stButton button p {
        font-size: 21px !important;
        font-weight: 700 !important;
    }
    
    /* Chữ ghi chú caption & thông báo */
    .stCaption, .stAlert p {
        font-size: 18px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP & CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown("""
        <div class="sidebar-title">
            🔑 ĐĂNG NHẬP & CẤU HÌNH
        </div>
    """, unsafe_allow_html=True)
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã API Key vào đây...",
        help="Nhập API Key để kích hoạt trợ lý AI"
    )
    
   model_name = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"],
        index=0,
        help="Khuyên dùng gemini-1.5-flash để tốc độ xử lý nhanh và ổn định nhất"
    )
    
    st.markdown("---")
    
    st.markdown("""
        <div class="sidebar-title">
            👤 THÔNG TIN GIÁO VIÊN
        </div>
    """, unsafe_allow_html=True)
    
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")
    
    st.markdown("---")
    st.caption("🟢 **Trạng thái:** Hệ thống sẵn sàng")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 25px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 22px; border-radius: 12px; border: 1px solid #bfdbfe;">
        <div style="font-size: 33px; font-weight: 800; color: #1e3a8a; line-height: 1.3;">
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
        "Môn học/Hoạt động GD:", 
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
# BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV (ĐÃ ĐIỀU CHỈNH GIAO DIỆN)
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

# Khung 1: Tra cứu tự động hoặc Tải File SGV (Bố cục 2 cột cân đối)
col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Tải danh mục bài học từ taphuan.nxbgd.vn:</span>', unsafe_allow_html=True)
    if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                
                prompt_fetch = f"""
                Hãy đóng vai Cơ sở dữ liệu chính thức của NXB Giáo dục Việt Nam (taphuan.nxbgd.vn).
                Liệt kê ĐẦY ĐỦ, ĐÚNG THỨ TỰ tất cả các Bài học thuộc môn {subject} - {grade} (Bộ sách Kết nối tri thức/GDPT 2018).
                
                LƯU Ý ĐẶC BIỆT: Bắt buộc phải bao gồm đầy đủ cả các bài đánh số VÀ các bài "Bài tập cuối chương...", "Ôn tập chương..." ở cuối mỗi chương.
                
                Trả về duy nhất dạng JSON theo cấu trúc mảng:
                [
                  {{
                    "chapter": "Tên Chương 1",
                    "lesson": "Tên Bài 1 hoặc Bài tập cuối chương...",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt chuẩn của bài"
                  }}
                ]
                Chỉ trả về mã JSON nguyên bản trong mảng [ ... ], không thêm bất kỳ văn bản giải thích nào khác.
                """
                
                with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade}..."):
                    res = model.generate_content(prompt_fetch)
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã tải xong danh mục bài học chuẩn đầy đủ!")
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg or "Quota exceeded" in err_msg:
                    st.warning("⚠️ API Key của thầy đã chạm trần lượt gọi miễn phí (Quota 429). Hệ thống tự động chuyển sang dữ liệu bài học mẫu chuẩn!")
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
                        },
                        {
                            "chapter": "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số",
                            "lesson": "Bài tập cuối chương I",
                            "duration": 2,
                            "req": "Củng cố kiến thức về khảo sát và vẽ đồ thị hàm số, giải các bài toán thực tế liên quan."
                        }
                    ]
                else:
                    st.error(f"Lỗi khi tải danh mục bài học: {e}")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Hoặc tải lên File SGV (Sách Giáo Viên):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải lên File SGV:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

if uploaded_sgv_file is not None:
    if st.button("🔍 Đọc Yêu cầu cần đạt từ File SGV", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                with st.spinner("✨ AI đang trích xuất Yêu cầu cần đạt từ File SGV..."):
                    if uploaded_sgv_file.type in ["image/png", "image/jpeg", "image/jpg"]:
                        img = Image.open(uploaded_sgv_file)
                        res = model.generate_content(["Trích xuất chính xác Yêu cầu cần đạt / Mục tiêu bài học trong file ảnh SGV này:", img])
                        st.session_state['req_from_sgv_file'] = res.text
                    else:
                        pdf_bytes = uploaded_sgv_file.read()
                        res = model.generate_content([
                            "Trích xuất chính xác Yêu cầu cần đạt / Mục tiêu bài học trong file PDF SGV này:",
                            {"mime_type": "application/pdf", "data": pdf_bytes}
                        ])
                        st.session_state['req_from_sgv_file'] = res.text
                st.success("🎉 Đã trích xuất Yêu cầu cần đạt từ File SGV thành công!")
            except Exception as e:
                st.error(f"Lỗi khi xử lý File SGV: {e}")

st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# Lấy giá trị YCĐ ưu tiên từ file SGV vừa đọc (nếu có)
req_default_value = st.session_state.get('req_from_sgv_file', '')

# Khung 2: Hiển thị Selectbox và các Ô nhập liệu thông tin bài học
if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
    selected_idx = st.selectbox(
        "👉 Chọn Bài học chuẩn từ danh sách vừa tải:", 
        range(len(lesson_titles)), 
        format_func=lambda x: lesson_titles[x],
        label_visibility="collapsed"
    )
    
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
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt (Quy định chuẩn NXB Giáo dục / SGV):</span>', unsafe_allow_html=True)
        final_req = req_default_value if req_default_value else current_item['req']
        requirements = st.text_area("YCĐ:", value=final_req, height=230, label_visibility="collapsed")

else:
    st.info("💡 Thầy có thể bấm nút **'🔍 Cập nhật danh sách Chương & Bài học...'** để AI tự động tải danh mục bài chuẩn hoặc **Tải file SGV** ở trên!")
    
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
        requirements = st.text_area("YCĐ:", value=req_default_value, placeholder="Nhập yêu cầu cần đạt hoặc đọc tự động từ File SGV...", height=230, label_visibility="collapsed")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)

st.markdown('<span class="custom-label">Lựa chọn yếu tố tích hợp:</span>', unsafe_allow_html=True)
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
    table.autofit = False
    
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    run_school = p_left.add_run(f"TRƯỜNG: {school_name.upper()}\n")
    run_school.bold = True
    run_dept = p_left.add_run(f"TỔ: {dept_name.upper()}")
    run_dept.bold = True

    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    run_teacher_label = p_right.add_run("Họ và tên giáo viên:\n")
    run_teacher_label.bold = True
    run_teacher_val = p_right.add_run(teacher_name)
    run_teacher_val.bold = True

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
    r_sub = p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)")
    r_sub.italic = True

    lines = content_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str or line_str.startswith("---") or line_str.startswith("# "):
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)

        clean_text = line_str.replace("**", "").replace("*", "")

        if clean_text.startswith("I. ") or clean_text.startswith("II. ") or clean_text.startswith("III. ") or clean_text.startswith("IV. "):
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(14)
        elif clean_text.startswith("TIẾT ") or clean_text.startswith("HOẠT ĐỘNG ") or clean_text.startswith("Nội dung ") or clean_text.startswith("Khối kiến thức "):
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(13)
        elif clean_text.startswith("1. ") or clean_text.startswith("2. ") or clean_text.startswith("3. ") or clean_text.startswith("4. "):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith("2.1") or clean_text.startswith("2.2") or clean_text.startswith("2.3"):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.15)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith("a)") or clean_text.startswith("b)") or clean_text.startswith("c)") or clean_text.startswith("d)"):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.15)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith("Bước 1:") or clean_text.startswith("Bước 2:") or clean_text.startswith("Bước 3:") or clean_text.startswith("Bước 4:") or clean_text.startswith("- Bước 1:") or clean_text.startswith("- Bước 2:") or clean_text.startswith("- Bước 3:") or clean_text.startswith("- Bước 4:"):
            p.paragraph_format.left_indent = Inches(0.3)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith("- ") or clean_text.startswith("+ "):
            p.paragraph_format.left_indent = Inches(0.3)
            p.add_run(clean_text)
        else:
            p.add_run(clean_text)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng chọn hoặc nhập tên Bài dạy!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)

            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Hãy soạn Kế hoạch bài dạy chuẩn 100% CÔNG VĂN 5512/BGDĐT:
            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - Yêu cầu cần đạt: {requirements}
            - YẾU TỐ TÍCH HỢP: {integration_str}

            QUY ĐỊNH ĐỊNH DẠNG VĂN BẢN TRẢ VỀ (RẤT QUAN TRỌNG):
            - Xuất nội dung trực tiếp, không có lời chào hỏi, không dùng ký tự kẻ ngang (---).
            - Trình bày chính xác theo cấu trúc mục tiêu và tiến trình chuẩn Công văn 5512:
              I. MỤC TIÊU (In hoa hoàn toàn)
              1. Kiến thức:
              2. Năng lực: (2.1. Năng lực toán học/chuyên môn, 2.2. Năng lực chung, 2.3. Năng lực Số / Ứng dụng CNTT và AI)
              3. Phẩm chất:
              II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (In hoa hoàn toàn)
              1. Giáo viên:
              2. Học sinh:
              III. TIẾN TRÌNH DẠY HỌC (In hoa hoàn toàn)
              Phân chia các TIẾT, HOẠT ĐỘNG, Nội dung kiến thức cụ thể.
              Mỗi hoạt động trình bày đúng 4 mục: a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện.
              Trong phần d) Tổ chức thực hiện trình bày rõ 4 bước: - Bước 1: Chuyển giao nhiệm vụ, - Bước 2: Thực hiện nhiệm vụ, - Bước 3: Báo cáo, thảo luận, - Bước 4: Kết luận, nhận định.
            """

            with st.spinner("✨ AI đang tạo Kế hoạch bài dạy 5512..."):
                response = model.generate_content(prompt)
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
                
                preview_header = f"""
<table style="width:100%; border:none; margin-bottom: 20px;">
  <tr>
    <td style="text-align:left; border:none;"><b>TRƯỜNG: {school_name.upper()}</b><br><b>TỔ: {dept_name.upper()}</b></td>
    <td style="text-align:right; border:none;"><b>Họ và tên giáo viên:</b><br><b>{teacher_name}</b></td>
  </tr>
</table>

<h3 style="text-align: center; margin-top: 10px;">TÊN BÀI DẠY: {lesson_title.upper()}</h3>
<p style="text-align: center; font-style: italic;">Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}<br>Thời gian thực hiện: ({duration} tiết)</p>

---
                """
                st.markdown(preview_header, unsafe_allow_html=True)
                st.markdown(response.text)

        except Exception as e:
            st.error(f"Đã có lỗi xảy ra: {e}")
