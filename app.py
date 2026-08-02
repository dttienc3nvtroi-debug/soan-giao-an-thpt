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
# CẤU HÌNH GIAO DIỆN & TĂNG KÍCH THƯỚC CSS
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

    /* 6. ĐẶC TRỊ TĂNG FONT CHO CÁC Ô SELECTBOX VÀ INPUT */
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

    ul[role="listbox"] li,
    ul[role="listbox"] li * {
        font-size: 21px !important;
        font-weight: 600 !important;
    }

    span[data-baseweb="tag"] * {
        font-size: 18px !important;
        font-weight: 700 !important;
    }

    div[data-baseweb="input"] input,
    .stTextInput input,
    .stNumberInput input {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
        padding-top: 8px !important;
        padding-bottom: 8px !important;
    }

    div[data-baseweb="textarea"] textarea,
    .stTextArea textarea {
        font-size: 21px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
        line-height: 1.4 !important;
    }

    .stNumberInput button * {
        font-size: 19px !important;
    }

    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
    }
    
    .stButton button p {
        font-size: 22px !important;
        font-weight: 700 !important;
    }
    
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
        ["gemini-gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"],
        index=0,
        help="Khuyên dùng gemini-3.6-flash cho tốc độ soạn thảo nhanh và tối ưu nhất"
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
# BƯỚC 2: TRA CỨU & NHẬP NỘI DUNG SGV / CHỌN FILE SGV
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & NHẬP NỘI DUNG SGV (TAPHUAN.NXBGD.VN)</div>', unsafe_allow_html=True)

# --- BỔ SUNG CHỌN FILE SGV ---
st.markdown('<span class="custom-label">📂 Chọn File SGV (Tải từ taphuan.nxbgd.vn - PDF hoặc Ảnh):</span>', unsafe_allow_html=True)
uploaded_sgv_file = st.file_uploader(
    "Tải lên File SGV:", 
    type=["pdf", "png", "jpg", "jpeg"], 
    label_visibility="collapsed"
)

if uploaded_sgv_file is not None:
    if st.button("🔍 Đọc Yêu cầu cần đạt từ File SGV"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh bên trái!")
        else:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                with st.spinner("AI đang đọc và trích xuất YCĐ chính xác từ File SGV..."):
                    if uploaded_sgv_file.type in ["image/png", "image/jpeg", "image/jpg"]:
                        img = Image.open(uploaded_sgv_file)
                        res = model.generate_content(["Hãy trích xuất NGUYÊN VĂN toàn bộ phần Yêu cầu cần đạt / Mục tiêu bài học trong file ảnh SGV này:", img])
                        st.session_state['req_from_file'] = res.text
                    else:
                        # Với file PDF SGV
                        pdf_bytes = uploaded_sgv_file.read()
                        res = model.generate_content([
                            "Hãy trích xuất NGUYÊN VĂN toàn bộ phần Yêu cầu cần đạt / Mục tiêu bài học trong tài liệu SGV này:",
                            {"mime_type": "application/pdf", "data": pdf_bytes}
                        ])
                        st.session_state['req_from_file'] = res.text
                st.success("🎉 Đã đọc thành công YCĐ từ File SGV!")
            except Exception as e:
                st.error(f"Lỗi khi đọc file SGV: {e}")

st.markdown("<br>", unsafe_allow_html=True)

col_i1, col_i2 = st.columns([1, 2])

with col_i1:
    st.markdown('<span class="custom-label">Chương / Chủ đề:</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương:", value="Chương I", placeholder="Nhập tên chương...", label_visibility="collapsed")
    
    st.markdown('<span class="custom-label">Tên bài dạy:</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài:", value="", placeholder="Nhập tên bài học chuẩn SGV...", label_visibility="collapsed")
    
    st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết:", value=2, label_visibility="collapsed")

with col_i2:
    st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt (Trích từ SGV - taphuan.nxbgd.vn):</span>', unsafe_allow_html=True)
    
    # Lấy nội dung tự động nếu vừa trích xuất từ file SGV
    default_req_val = st.session_state.get('req_from_file', '')
    
    requirements = st.text_area(
        "YCĐ:", 
        value=default_req_val, 
        placeholder="Nội dung YCĐ sẽ tự động điền khi đọc File SGV ở trên, hoặc thầy có thể copy/paste trực tiếp từ taphuan.nxbgd.vn vào đây...", 
        height=210, 
        label_visibility="collapsed"
    )

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
    elif not requirements:
        st.error("⚠️ Vui lòng nhập hoặc tải file SGV để trích xuất Yêu cầu cần đạt!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)

            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Bạn là trợ lý biên soạn giáo án chuyên nghiệp. Hãy sử dụng NGUYÊN VĂN dữ liệu Yêu cầu cần đạt từ Sách Giáo Viên (taphuan.nxbgd.vn) dưới đây để lập Kế hoạch bài dạy chuẩn Công văn 5512:

            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - YÊU CẦU CẦN ĐẠT CHUẨN SGV: 
            {requirements}
            
            - YẾU TỐ TÍCH HỢP: {integration_str}

            YÊU CẦU NỘI DUNG VÀ CẤU TRÚC:
            I. MỤC TIÊU:
               1. Kiến thức: Nêu chính xác các nội dung kiến thức dựa theo YCĐ chuẩn SGV.
               2. Năng lực:
                  - 2.1. Năng lực đặc thù ({subject}): Bám sát YCĐ từ SGV.
                  - 2.2. Năng lực chung: Tự chủ & tự học, giao tiếp & hợp tác, giải quyết vấn đề sáng tạo.
                  - 2.3. Năng lực Số / Ứng dụng CNTT và AI: Chi tiết các công cụ sử dụng.
               3. Phẩm chất: Yêu nước, nhân ái, chăm chỉ, trung thực, trách nhiệm.
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Chia rõ từng tiết, mỗi hoạt động đủ 4 bước: Chuyển giao -> Thực hiện -> Báo cáo -> Kết luận).
            
            Quy định: Trả về văn bản Markdown rõ ràng, không kèm lời chào.
            """

            with st.spinner("✨ AI đang tạo Kế hoạch bài dạy chuẩn SGV 5512..."):
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
