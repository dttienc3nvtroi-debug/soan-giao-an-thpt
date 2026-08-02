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

# Cấu hình trang Streamlit
st.set_page_config(page_title="Hệ thống Soạn Giáo án Tự Động 5512", layout="wide", page_icon="📝")

# ==========================================
# CẤU HÌNH GIAO DIỆN CHUẨN DASHBOARD MODERN (THEO ẢNH MẪU)
# ==========================================
st.markdown("""
    <style>
    /* Font chữ mặc định */
    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', Roboto, sans-serif;
    }
    
    /* Nền tổng thể xám xịn hiện đại */
    .stApp {
        background-color: #F4F6F9;
    }
    
    /* Định dạng Khung / Card trắng bo tròn bóng đổ */
    div[data-testid="stVerticalBlock"] > div.element-container {
        font-family: 'Segoe UI', Roboto, sans-serif;
    }

    /* Tùy chỉnh các khối container chính */
    .stCard {
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.05);
        border: 1px solid #E2E8F0;
        margin-bottom: 20px;
    }
    
    /* Căn giữa tiêu đề chính */
    .main-header {
        text-align: center;
        margin-top: -20px;
        margin-bottom: 30px;
    }
    .main-title {
        font-size: 28px;
        font-weight: 800;
        color: #1E293B;
        letter-spacing: -0.5px;
    }
    .author-title {
        font-size: 18px;
        font-weight: 700;
        color: #007BFF;
        margin-top: 5px;
    }

    /* Thiết kế nút bấm chuyển màu Gradient sinh động */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #FF4B4B 0%, #007BFF 100%) !important;
        color: white !important;
        font-weight: 800 !important;
        font-size: 18px !important;
        border-radius: 30px !important;
        padding: 12px 35px !important;
        border: none !important;
        box-shadow: 0 6px 20px rgba(0, 123, 255, 0.3) !important;
        display: block !important;
        margin: 0 auto !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(0, 123, 255, 0.5) !important;
    }

    /* Tùy chỉnh Nút bấm thường */
    div.stButton > button:not([kind="primary"]) {
        background-color: #3B82F6 !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        width: 100% !important;
        border: none !important;
    }

    /* Ẩn viền cạnh mặc định */
    header[data-testid="stHeader"] {
        background-color: transparent;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP & CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown("### 🔑 ĐĂNG NHẬP & CẤU HÌNH")
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã AI Key vào đây...",
        help="Nhập API Key để kích hoạt trợ lý AI"
    )
    
    model_name = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"],
        index=0,
        help="Khuyên dùng 2.5-flash cho tốc độ soạn thảo nhanh nhất"
    )
    
    st.markdown("---")
    st.markdown("### 👤 THÔNG TIN GIÁO VIÊN")
    
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")
    
    st.markdown("---")
    st.caption("🟢 **Trạng thái:** Hệ thống sẵn sàng")

# ==========================================
# TIÊU ĐỀ CHÍNH GIỮA (GIỐNG HÌNH MẪU)
# ==========================================
st.markdown("""
    <div class="main-header">
        <div class="main-title">HỆ THỐNG SOẠN KHBD (có tích hợp NLS, AI, STEM,...)</div>
        <div class="author-title">Tác giả: DƯƠNG TẤN TIẾN - GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# KHỐI BƯỚC 1 & BƯỚC 2 (CHIA 2 CỘT NHƯ HÌNH)
# ==========================================
col_left, col_right = st.columns(2, gap="large")

with col_left:
    with st.container():
        st.markdown("<h4 style='color: #1E293B; margin-bottom: 15px;'>📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</h4>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            subject = st.selectbox("Môn học/Hoạt động GD", [
                "Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", 
                "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"
            ])
        with c2:
            grade = st.selectbox("Khối lớp", ["Lớp 10", "Lớp 11", "Lớp 12"])

with col_right:
    with st.container():
        st.markdown("<h4 style='color: #1E293B; margin-bottom: 15px;'>📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN</h4>", unsafe_allow_html=True)
        
        btn_fetch = st.button("Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn")
        
        if btn_fetch:
            if not api_key:
                st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
            else:
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(model_name)
                    
                    prompt_fetch = f"""
                    Hãy đóng vai Cơ sở dữ liệu chính thức của NXB Giáo dục Việt Nam (taphuan.nxbgd.vn).
                    Liệt kê ĐẦY ĐỦ, ĐÚNG THỨ TỰ tất cả các Bài học thuộc môn {subject} - {grade} (Bộ sách Kết nối tri thức/GDPT 2018).
                    
                    Trả về duy nhất dạng JSON theo cấu trúc:
                    [
                      {{
                        "chapter": "Tên Chương 1",
                        "lesson": "Tên Bài 1",
                        "duration": 2,
                        "req": "Yêu cầu cần đạt chuẩn của bài 1"
                      }}
                    ]
                    Chỉ trả về mã JSON nguyên bản, không thêm văn bản giải thích hay định dạng markdown khác.
                    """
                    
                    with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade}..."):
                        res = model.generate_content(prompt_fetch)
                        clean_json = res.text.replace("```json", "").replace("```", "").strip()
                        st.session_state['fetched_lessons'] = json.loads(clean_json)
                        st.success("🎉 Đã tải xong danh mục bài học chuẩn!")
                except Exception as e:
                    st.error(f"Lỗi khi tải danh mục bài học: {e}")

        if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
            lessons_data = st.session_state['fetched_lessons']
            lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
            selected_idx = st.selectbox("👉 Chọn Bài học chuẩn:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x])
            
            current_item = lessons_data[selected_idx]
            chapter_title = st.text_input("Tên Chương/Chủ đề", value=current_item['chapter'])
            lesson_title = st.text_input("Tên Bài dạy", value=current_item['lesson'])
            duration = st.number_input("Số tiết", value=current_item['duration'])
            requirements = st.text_area("Yêu cầu cần đạt", value=current_item['req'], height=100)
        else:
            st.info("💡 <b>Gợi ý:</b> Hãy bấm nút cập nhật ở trên để tự động tải danh mục bài học.", icon="💡")
            chapter_title = st.text_input("Tên Chương/Chủ đề", placeholder="Nhập tên chương...")
            lesson_title = st.text_input("Tên Bài dạy", placeholder="Nhập tên bài dạy...")
            duration = st.number_input("Số tiết", value=2)
            requirements = st.text_area("Yêu cầu cần đạt", placeholder="Nhập yêu cầu cần đạt...", height=80)

# ==========================================
# KHỐI BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
with st.container():
    st.markdown("<h4 style='color: #1E293B; margin-bottom: 15px;'>🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</h4>", unsafe_allow_html=True)
    integrations = st.multiselect(
        "Lựa chọn các yếu tố tích hợp:",
        [
            "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
            "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
            "Giáo dục STEM / STEAM", 
            "Phát triển Tư duy phản biện", 
            "Học tập qua Dự án (PBL)"
        ],
        default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
    )

# ==========================================
# NÚT BẮT ĐẦU TẠO KHBD NẰM CHÍNH GIỮA (GRADIENT)
# ==========================================
st.markdown("<br><br>", unsafe_allow_html=True)
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])

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
            run.font.size = Pt(13)
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

with col_btn2:
    submit = st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary")

if submit:
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
              2. Năng lực: (2.1. Năng lực toán học, 2.2. Năng lực chung, 2.3. Năng lực Số / Ứng dụng CNTT và AI)
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
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
