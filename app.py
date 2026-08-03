import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import time

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.markdown("""
    <style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; max-width: 1300px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    .sidebar-title { color: #0f172a; font-size: 20px !important; font-weight: 700; margin-bottom: 12px; border-bottom: 3px solid #2563eb; padding-bottom: 4px; }
    .step-header { color: #dc2626 !important; font-size: 22px !important; font-weight: 700 !important; margin-top: 18px !important; margin-bottom: 10px !important; padding-left: 10px; border-left: 5px solid #dc2626; }
    div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #1e293b !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# HÀM XỬ LÝ MODEL GEMINI CỰC NHANH
# ==========================================
def get_working_model(api_key):
    genai.configure(api_key=api_key)
    try:
        models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        # Ưu tiên flash
        for m in models:
            if "flash" in m.lower() and "-tts" not in m.lower():
                return m
        return models[0] if models else "gemini-1.5-flash"
    except Exception:
        return "gemini-1.5-flash"

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán mã API Key vào đây...")
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 20px; background: #eff6ff; padding: 18px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 28px; font-weight: 800; color: #1e3a8a;">HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 5512</div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 5px;">Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: BÀI HỌC & CHƯƠNG TRÌNH
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: THÔNG TIN BÀI DẠY</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    subject = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"])
with c2:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
with c3:
    book_series = st.selectbox("Bộ sách:", ["Kết nối tri thức với cuộc sống", "Cánh diều", "Chân trời sáng tạo"])

c_l1, c_l2, c_l3 = st.columns([1.5, 2, 0.8])
with c_l1:
    chapter_title = st.text_input("Tên Chương/Chủ đề:", placeholder="Ví dụ: Chương I. Mệnh đề và Tập hợp")
with c_l2:
    lesson_title = st.text_input("Tên Bài dạy cụ thể:", placeholder="Ví dụ: Bài 1. Mệnh đề")
with c_l3:
    duration = st.number_input("Số tiết:", value=2, min_value=1)

# ==========================================
# BƯỚC 2: YÊU CẦU CẦN ĐẠT (AI TỰ TẠO HOẶC NHẬP)
# ==========================================
st.markdown('<div class="step-header">🎯 BƯỚC 2: YÊU CẦU CẦN ĐẠT (YCĐ)</div>', unsafe_allow_html=True)

col_btn, col_file = st.columns([1, 1])

with col_btn:
    if st.button("✨ Tự động tra cứu YCĐ chuẩn SGV từ AI (Cực nhanh)", type="secondary", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng nhập API Key ở menu bên trái!")
        elif not lesson_title:
            st.error("⚠️ Vui lòng nhập Tên Bài dạy trước!")
        else:
            try:
                model_name = get_working_model(api_key.strip())
                model = genai.GenerativeModel(model_name)
                
                prompt_ycd = f"""Bạn là Chuyên gia Giáo dục môn {subject}. Hãy trích xuất NGUYÊN VĂN Yêu cầu cần đạt theo đúng SGK {book_series} - {grade} cho bài: "{lesson_title}" (Thuộc {chapter_title}). Trả về ngắn gọn dạng các gạch đầu dòng."""
                
                with st.spinner("⚡ AI đang lấy YCĐ từ CSDL SGK..."):
                    res = model.generate_content(prompt_ycd)
                    st.session_state['auto_ycd'] = res.text
                    st.success("✅ Đã tải thành công YCĐ chuẩn SGK!")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

with col_file:
    uploaded_sgv_file = st.file_uploader("Hoặc Tải ảnh/PDF trang SGV bài dạy này:", type=["pdf", "png", "jpg", "jpeg"])

ycd_default = st.session_state.get('auto_ycd', '')
requirements = st.text_area("Nội dung Yêu cầu cần đạt (Có thể chỉnh sửa):", value=ycd_default, height=150, placeholder="Nhấn nút Tra cứu tự động ở trên hoặc tự điền YCĐ vào đây...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC & CÔNG NGHỆ</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp vào KHBD:",
    [
        "Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)", 
        "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện & Tự học"
    ],
    default=["Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)", "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)"]
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

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"KẾ HOẠCH BÀI DẠY: {lesson_title.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn: {subject} ({grade}) - Sách: {book_series}\nThời lượng: {duration} tiết").italic = True

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
            p.paragraph_format.space_before = Pt(10)
            run = p.add_run(clean_text)
            run.bold = True
            run.font.size = Pt(13)
        elif clean_text.startswith(("HOẠT ĐỘNG ", "Tiết ", "1. ", "2. ", "3. ", "4. ")):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(clean_text)
            run.bold = True
        elif clean_text.startswith(("a)", "b)", "c)", "d)")):
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
if st.button("🚀 XUẤT GIÁO ÁN WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập Tên bài dạy!")
    else:
        try:
            model_name = get_working_model(api_key.strip())
            model = genai.GenerativeModel(model_name)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt_5512 = f"""
            Hãy đóng vai Giáo viên Giỏi môn {subject}. Viết Kế hoạch bài dạy (Giáo án) CHUẨN CÔNG VĂN 5512 đầy đủ và chi tiết.
            
            THÔNG TIN BÀI DẠY:
            - Môn: {subject} ({grade}) - Bộ sách: {book_series}
            - Bài dạy: {lesson_title} (Thuộc {chapter_title})
            - Thời lượng: {duration} tiết
            - Tích hợp: {integration_str}
            - Yêu cầu cần đạt: {requirements}

            CẤU TRÚC BẮT BUỘC (ĐÚNG MẪU 5512):
            I. MỤC TIÊU
            1. Về kiến thức
            2. Về năng lực (Năng lực chung + Năng lực đặc thù môn học)
            3. Về phẩm chất
            
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (Giáo viên & Học sinh)
            
            III. TIẾN TRÌNH DẠY HỌC
            Chia chi tiết cho cả {duration} tiết học với các Hoạt động (Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng).
            Mỗi hoạt động BẮT BUỘC có 4 mục:
            a) Mục tiêu
            b) Nội dung (Câu hỏi/Bài tập cụ thể)
            c) Sản phẩm (Đáp án/Kết quả chi tiết)
            d) Tổ chức thực hiện (Bước 1: Chuyển giao NV; Bước 2: Thực hiện NV; Bước 3: Báo cáo thảo luận; Bước 4: Kết luận, nhận định)
            """

            contents = [prompt_5512]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                contents.append({"mime_type": uploaded_sgv_file.type, "data": bytes_data})

            with st.spinner("⚡ Đang soạn Kế hoạch bài dạy 5512 chi tiết..."):
                res = model.generate_content(contents)
                doc_file = generate_doc(res.text)
                
                st.success("🎉 Đã tạo xong Giáo án Word 5512!")
                st.download_button(
                    label="📥 TẢI FILE WORD (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown(res.text)

        except Exception as e:
            st.error(f"❌ Lỗi: {str(e)}")
