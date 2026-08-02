import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# CSS Tăng cỡ chữ và giao diện chuẩn
st.markdown("""
    <style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; max-width: 1300px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    
    .step-header {
        color: #dc2626 !important; font-size: 24px !important; font-weight: 700 !important;
        margin-top: 20px !important; margin-bottom: 12px !important;
        padding-left: 10px; border-left: 5px solid #dc2626;
    }

    div[data-testid="stWidgetLabel"] p, .custom-label {
        font-size: 21px !important; font-weight: 700 !important; color: #1e293b !important;
    }

    /* Ép font to cho tất cả Selectbox và Input */
    .stSelectbox div[data-baseweb="select"] *, div[data-baseweb="input"] input, div[data-baseweb="textarea"] textarea {
        font-size: 21px !important; font-weight: 700 !important; color: #1e3a8a !important;
    }
    
    .stButton button p { font-size: 22px !important; font-weight: 700 !important; }
    </style>
""", unsafe_allow_html=True)

# SIDEBAR
with st.sidebar:
    st.title("🔑 CẤU HÌNH")
    api_key = st.text_input("Google Gemini API Key:", type="password")
    model_name = st.selectbox("Mô hình AI:", ["gemini-3.6-flash", "gemini-3.6-flash"])
    st.markdown("---")
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# HEADER
st.markdown("""
    <div style="text-align: center; background: #eff6ff; padding: 20px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <h2 style="color: #1e3a8a; margin:0;">HỆ THỐNG SOẠN KHBD TỰ ĐỘNG (CHUẨN 5512)</h2>
        <p style="color: #2563eb; font-weight:600; margin-top:5px;">Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI</p>
    </div>
""", unsafe_allow_html=True)

# BƯỚC 1
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    subject = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học"])
with col2:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

# BƯỚC 2
st.markdown('<div class="step-header">📖 BƯỚC 2: NHẬP THÔNG TIN BÀI HỌC & TRÍCH XUẤT SGV</div>', unsafe_allow_html=True)
col_a, col_b = st.columns([1, 2])

with col_a:
    chapter_title = st.text_input("Chương / Chủ đề:", value="Chương I.Ứng dụng đạo hàm để khảo sát hàm số")
    lesson_title = st.text_input("Tên bài dạy:", value="Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số")
    duration = st.number_input("Số tiết thực hiện:", value=2, min_value=1)

with col_b:
    st.markdown('<span class="custom-label">📌 Dán Yêu cầu cần đạt / Mục tiêu trích từ SGV (taphuan.nxbgd.vn):</span>', unsafe_allow_html=True)
    requirements = st.text_area(
        "Mục tiêu SGV:", 
        value="""- Nhận biết và tính được giá trị lớn nhất, giá trị nhỏ nhất của hàm số trên một tập hợp cho trước.
- Giải quyết được một số vấn đề thực tiễn gắn với tìm giá trị lớn nhất, giá trị nhỏ nhất của hàm số.""", 
        height=200,
        help="Thầy copy trực tiếp đoạn Mục tiêu/YCĐ trong SGV trên taphuan.nxbgd.vn dán vào đây để đảm bảo chính xác 100%!"
    )

# BƯỚC 3
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Geogebra, Padlet...)", "Tích hợp AI trong dạy học", "Giáo dục STEM/STEAM"],
    default=["Năng lực Số / Ứng dụng CNTT (Geogebra, Padlet...)"]
)

# TẠO WORD FILE
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

    # Bảng thông tin trường / GV
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    p_left = table.cell(0, 0).paragraphs[0]
    p_left.add_run(f"TRƯỜNG: {school_name.upper()}\nTỔ: {dept_name.upper()}").bold = True

    p_right = table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run(f"Họ và tên giáo viên:\n{teacher_name}").bold = True

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"\nTÊN BÀI DẠY: {lesson_title.upper()}\n")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"Môn học: {subject}; Lớp: {grade}\nThời gian thực hiện: ({duration} tiết)\n")
    r_sub.italic = True

    for line in content_text.split('\n'):
        line_str = line.strip()
        if not line_str or line_str.startswith("---"): continue
        p = doc.add_paragraph()
        clean = line_str.replace("**", "").replace("*", "")
        if clean.startswith("I. ") or clean.startswith("II. ") or clean.startswith("III. "):
            run = p.add_run(clean)
            run.bold = True
            run.font.size = Pt(14)
        elif clean.startswith("1. ") or clean.startswith("2. ") or clean.startswith("3. ") or clean.startswith("a)") or clean.startswith("b)"):
            run = p.add_run(clean)
            run.bold = True
        else:
            p.add_run(clean)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# THỰC THI SOẠN BÀI
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập API Key!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            
            # PROMPT ÉP CHẶT NỘI DUNG SGV CỦA THẦY TIẾN
            prompt = f"""
            Bạn là trợ lý biên soạn giáo án. Hãy sử dụng NGUYÊN VĂN dữ liệu Sách Giáo Viên (SGV) do giáo viên cung cấp dưới đây để lập Kế hoạch bài dạy chuẩn Công văn 5512.

            [DỮ LIỆU SGV GỐC BẮT BUỘC DÙNG]:
            {requirements}

            [YÊU CẦU BẮT BUỘC]:
            1. Ở mục "I. MỤC TIÊU": 
               - Mục "1. Kiến thức" và "2. Năng lực đặc thù": KHÔNG ĐƯỢC TỰ Ý THÊM BỚT HOẶC THAY ĐỔI CÂU CHỮ. Bắt buộc phải giữ nguyên văn các ý từ [DỮ LIỆU SGV GỐC] ở trên.
               - Mục "2.2 Năng lực chung" và "3. Phẩm chất": Bổ sung ngắn gọn phù hợp với bài học.
            2. Trình bày tiến trình bài dạy (III. TIẾN TRÌNH DẠY HỌC) theo chuẩn 4 hoạt động của CV 5512 (Mỗi hoạt động có đủ: a. Mục tiêu, b. Nội dung, c. Sản phẩm, d. Tổ chức thực hiện với 4 bước).
            3. Không viết lời chào, xuất thẳng nội dung giáo án.
            """

            with st.spinner("✨ Đang chuyển đổi dữ liệu SGV sang KHBD 5512..."):
                res = model.generate_content(prompt)
                st.success("🎉 Đã tạo xong giáo án chuẩn 100% nội dung SGV!")
                doc = generate_doc(res.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD (.DOCX)",
                    data=doc,
                    file_name=f"KHBD_{lesson_title}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                st.markdown(res.text)
        except Exception as e:
            st.error(f"Lỗi: {e}")
