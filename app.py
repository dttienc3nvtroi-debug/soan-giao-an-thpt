import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, Cm
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
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn SGK & SGV)", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# CẤU HÌNH GIAO DIỆN & FONT CHỮ TRỰC QUAN
# ==========================================
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.5rem !important;
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
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 4px !important;
    }
    .stSelectbox div[data-baseweb="select"] *,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="input"] input, .stTextInput input, .stNumberInput input {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="textarea"] textarea, .stTextArea textarea {
        font-size: 17px !important;
        font-weight: 500 !important;
        color: #0f172a !important;
    }
    .stButton button p {
        font-size: 19px !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Hàm gọi AI với Temperature = 0.0 để CHỐNG SÁNG TẠO / SỬA CÂU CHỮ
def call_gemini_strict(model_name, api_key, contents, max_retries=3):
    genai.configure(api_key=api_key)
    # Cấu hình khóa độ sáng tạo = 0.0
    generation_config = genai.GenerationConfig(
        temperature=0.0,
        top_p=0.1,
        top_k=1
    )
    model = genai.GenerativeModel(model_name=model_name, generation_config=generation_config)
    
    for attempt in range(max_retries):
        try:
            response = model.generate_content(contents)
            return response
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    time.sleep((attempt + 1) * 6)
                    continue
            raise e

# ==========================================
# THANH BÊN (SIDEBAR) CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã API Key vào đây...",
        help="Nhập API Key từ Google AI Studio"
    )
    
    model_name_choice = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"],
        index=0
    )
    st.markdown("---")
    
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN PHÒNG / TRƯỜNG</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 20px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 18px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 28px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 100% SGK & SGV (5512)
        </div>
        <div style="font-size: 16px; font-weight: 600; color: #047857; margin-top: 4px;">
            📌 Giữ nguyên văn 100% dữ liệu SGK/SGV taphuan.nxbgd.vn — Định dạng Word 5512 Chuẩn
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
# BƯỚC 2: NẠP NGUYÊN VĂN NỘI DUNG SGK & SGV
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: NẠP DỮ LIỆU NGUYÊN VĂN SGK & SGV (taphuan.nxbgd.vn)</div>', unsafe_allow_html=True)

col_i1, col_i2 = st.columns([1, 2], gap="large")

with col_i1:
    st.markdown('<span class="custom-label">Chương / Chủ đề (Theo SGK):</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương:", value="Chương II. Vectơ và hệ trục tọa độ trong không gian", label_visibility="collapsed")
    
    st.markdown('<span class="custom-label">Tên bài dạy (Theo SGK):</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài:", value="Bài 7. Hệ trục tọa độ trong không gian", label_visibility="collapsed")
    
    st.markdown('<span class="custom-label">Số tiết thực hiện:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết:", value=3, label_visibility="collapsed")

with col_i2:
    st.markdown('<span class="custom-label">📌 Nội dung Trích xuất Nguyên văn từ SGV & SGK (Paste từ taphuan.nxbgd.vn hoặc File):</span>', unsafe_allow_html=True)
    sgv_raw_text = st.text_area(
        "Dữ liệu SGV/SGK:", 
        value="- Nhận biết được tọa độ của điểm, của vectơ đối với hệ trục tọa độ.\n- Vận dụng được tọa độ của vectơ để giải một số bài toán có liên quan đến thực tiễn.", 
        height=180, 
        placeholder="Dán chính xác từng câu chữ SGV/SGK vào đây để AI giữ nguyên 100%...",
        label_visibility="collapsed"
    )
    
    uploaded_sgv_file = st.file_uploader(
        "Hoặc tải tệp/ảnh SGV bổ sung (PDF, PNG, JPG):", 
        type=["pdf", "png", "jpg", "jpeg"]
    )

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ (5512)</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)"],
    label_visibility="collapsed"
)

# ==========================================
# THƯ VIỆN ĐỊNH DẠNG WORD CHUẨN 5512 (KHÔNG THAY ĐỔI)
# ==========================================
def generate_docx_5512_standard(content_text):
    doc = docx.Document()
    
    # Set Căn lề chuẩn Khung 5512: Trên 2cm, Dưới 2cm, Trái 3cm, Phải 2cm
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

    # Style mặc định: Times New Roman, 13pt
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)

    # Bảng Header Trường & GV
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    p_left = table.cell(0, 0).paragraphs[0]
    p_left.paragraph_format.line_spacing = 1.15
    p_left.paragraph_format.space_after = Pt(0)
    p_left.add_run(f"TRƯỜNG: {school_name.upper()}\n").bold = True
    p_left.add_run(f"TỔ: {dept_name.upper()}").bold = True

    p_right = table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.line_spacing = 1.15
    p_right.paragraph_format.space_after = Pt(0)
    p_right.add_run("Họ và tên giáo viên:\n").bold = True
    p_right.add_run(teacher_name).bold = True

    # Ẩn đường viền bảng Header
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # Tiêu đề Giáo án
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)").italic = True

    # Duyệt từng dòng văn bản AI trả về và định dạng chuẩn
    lines = content_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str or line_str.startswith("---") or line_str.startswith("#"):
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        clean_text = line_str.replace("**", "").replace("*", "")

        # Định dạng La Mã (I., II., III., IV.)
        if re.match(r'^(I|II|III|IV|V)\.\s', clean_text):
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(14)
        # Định dạng các mục lớn (1., 2., 3.)
        elif re.match(r'^\d+\.\s', clean_text):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(clean_text)
            run.bold = True
        # Định dạng các mục a), b), c), d)
        elif re.match(r'^[a-d]\)\s', clean_text):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.left_indent = Inches(0.2)
            run = p.add_run(clean_text)
            run.bold = True
        # Định dạng Tiến trình các Bước
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
# TIẾN HÀNH TẠO GIÁO ÁN
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 XUẤT GIÁO ÁN WORD CHUẨN 100% SGK & SGV (5512)", type="primary", use_container_width=True):
    clean_api_key = api_key.strip() if api_key else ""
    if not clean_api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập Tên Bài Dạy!")
    else:
        try:
            clean_model_name = model_name_choice.replace("models/", "").strip()
            integration_str = ", ".join(integrations) if integrations else "Không"

            # PROMPT KHÓA CỨNG AI (NÓI KHÔNG VỚI TỰ Ý SỬA TỪ NGUYÊN VĂN)
            prompt = f"""
            Bạn là hệ thống trích xuất dữ liệu chính xác tuyệt đối. Nhiệm vụ của bạn là lập Kế hoạch bài dạy (KHBD) chuẩn Công văn 5512.

            QUY TẮC BẮT BUỘC KHÔNG ĐƯỢC VI PHẠM:
            1. KHÔNG ĐƯỢC TỰ Ý THAY ĐỔI, KHÔNG ĐƯỢC BỔ SUNG, KHÔNG TỰ VIẾT LẠI CÂU CHỮ TRONG SGK VÀ SGV.
            2. Trích xuất CHÍNH XÁC NGUYÊN VĂN 100% dữ liệu SGV/SGK dưới đây vào đúng các mục tương ứng:
               [DỮ LIỆU SGV & SGK CUNG CẤP]:
               {sgv_raw_text}

            3. ĐỊNH DẠNG ĐÚNG KHUNG 5512:
               I. MỤC TIÊU
               1. Về kiến thức, kỹ năng: (Chép lại CHÍNH XÁC NGUYÊN VĂN từng câu chữ từ dữ liệu SGV trên)
               2. Về phẩm chất, năng lực: (Chép lại CHÍNH XÁC NGUYÊN VĂN từng câu chữ từ dữ liệu SGV trên)
               - Năng lực Số / CNTT: (Tích hợp: {integration_str})

               II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
               - Giáo viên: SGK, SGV môn {subject} {grade}, máy tính, tivi.
               - Học sinh: SGK, vở ghi.

               III. TIẾN TRÌNH DẠY HỌC
               Tổ chức đúng 4 Hoạt động chuẩn SGK (Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng). Mỗi hoạt động bắt buộc có đủ 4 mục:
               a) Mục tiêu
               b) Nội dung (Bám sát nội dung câu hỏi/bài tập từ SGK/SGV)
               c) Sản phẩm (Lời giải chuẩn xác)
               d) Tổ chức thực hiện:
                  - Bước 1: Chuyển giao nhiệm vụ
                  - Bước 2: Thực hiện nhiệm vụ
                  - Bước 3: Báo cáo, thảo luận
                  - Bước 4: Kết luận, nhận định

            THÔNG TIN CHUNG:
            Môn: {subject} ({grade}) | Bài: {lesson_title} | Thời lượng: {duration} tiết.
            """

            contents = [prompt]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                mime_type = uploaded_sgv_file.type
                contents.append({"mime_type": mime_type, "data": bytes_data})

            with st.spinner("✨ Đang trích xuất chính xác nguyên văn SGK/SGV và đóng gói File Word..."):
                response = call_gemini_strict(clean_model_name, clean_api_key, contents)
                st.success("🎉 Đã xuất thành công Giáo án chuẩn Word 5512!")
                
                # Tạo file Word
                doc_file = generate_docx_5512_standard(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN CHUẨN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_Chuẩn_SGV_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.subheader("📋 Xem trước nội dung xuất:")
                st.markdown(response.text)

        except Exception as e:
            err_str = str(e)
            if "429" in err_str:
                st.error("⏳ Hệ thống bận. Thầy bấm lại sau 5 giây hoặc đổi mô hình gemini-1.5-flash.")
            else:
                st.error(f"❌ Lỗi xử lý: `{err_str}`")
