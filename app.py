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
st.set_page_config(page_title="Hệ Thống Soạn KHBD", layout="wide", page_icon="📝")

# ==========================================
# NHÚNG TAILWIND CSS & FONTAWESOME TRỰC TIẾP
# ==========================================
st.markdown("""
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        /* Ghi đè phông chữ và nền Streamlit theo mẫu Tailwind */
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif !important;
        }
        .stApp {
            background-color: #F8FAFC !important;
        }
        /* Tùy chỉnh Sidebar sang trọng */
        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E5E7EB !important;
        }
        /* Tùy chỉnh Nút bấm Streamlit theo chuẩn Tailwind */
        div.stButton > button[kind="primary"] {
            background: linear-gradient(to right, #2563EB, #4F46E5) !important;
            color: white !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            border-radius: 16px !important;
            padding: 14px 32px !important;
            border: none !important;
            box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3) !important;
            transition: all 0.2s ease-in-out !important;
            width: 100% !important;
        }
        div.stButton > button[kind="primary"]:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 20px 25px -5px rgba(37, 99, 235, 0.4) !important;
        }
        /* CSS cho nút cập nhật danh sách */
        div.stButton > button:not([kind="primary"]) {
            background-color: #2563EB !important;
            color: white !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            border-radius: 12px !important;
            padding: 10px 16px !important;
            border: none !important;
            width: 100% !important;
            transition: background-color 0.2s !important;
        }
        div.stButton > button:not([kind="primary"]):hover {
            background-color: #1D4ED8 !important;
        }
        /* Bỏ khoảng cách thừa của Streamlit */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 64rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP & CẤU HÌNH
# ==========================================
with st.sidebar:
    st.markdown("""
        <div class="space-y-4">
            <div class="flex items-center gap-2 text-slate-800 font-bold text-base border-b pb-3">
                <i class="fa-solid fa-key text-blue-600"></i>
                <span>ĐĂNG NHẬP & CẤU HÌNH</span>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
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
    
    st.markdown("""
        <div class="pt-4 space-y-4">
            <div class="flex items-center gap-2 text-slate-800 font-bold text-base border-b pb-3">
                <i class="fa-solid fa-user-gear text-blue-600"></i>
                <span>THÔNG TIN GIÁO VIÊN</span>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")
    
    st.markdown("---")
    st.markdown("""
        <div class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-green-50 text-green-700 border border-green-200">
            <span class="w-2 h-2 rounded-full bg-green-500"></span> Hệ thống sẵn sàng
        </div>
    """, unsafe_allow_html=True)

# ==========================================
# HEADER TIÊU ĐỀ
# ==========================================
st.markdown("""
    <header class="text-center space-y-2 mb-8">
        <h1 class="text-2xl md:text-3xl font-bold tracking-tight text-[#1F2937]">
            HỆ THỐNG SOẠN KHBD <span class="text-lg md:text-2xl font-semibold text-gray-600">(có tích hợp NLS, AI, STEM,...)</span>
        </h1>
        <p class="text-sm md:text-base font-semibold text-[#2563EB]">
            Tác giả: DƯƠNG TẤN TIẾN - GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
        </p>
    </header>
""", unsafe_allow_html=True)

# ==========================================
# KHỐI BƯỚC 1 & BƯỚC 2 (ĐỊNH DẠNG TAILWIND CARD)
# ==========================================
col_left, col_right = st.columns(2, gap="medium")

with col_left:
    st.markdown("""
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow mb-6">
            <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2 mb-4">
                <span>📚</span> BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
            </h2>
        </div>
    """, unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        subject = st.selectbox("Môn học/Hoạt động GD", [
            "Vật lí", "Toán học", "Ngữ văn", "Tiếng Anh", "Hóa học", 
            "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"
        ])
    with c2:
        grade = st.selectbox("Khối lớp", ["Khối 10", "Khối 11", "Khối 12"])

    # TRẠNG THÁI HỆ THỐNG CARD
    st.markdown("""
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm mt-6 space-y-3">
            <div class="flex items-center justify-between border-b pb-3">
                <span class="text-xs font-semibold uppercase text-gray-400">Trạng thái hệ thống</span>
                <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-green-50 text-green-700 border border-green-200">
                    <span class="w-2 h-2 rounded-full bg-green-500"></span> Sẵn sàng
                </span>
            </div>
            <p class="text-sm text-gray-500">
                Vui lòng hoàn tất các bước tra cứu và tích hợp năng lực trước khi tiến hành xuất file Word chuẩn 5512.
            </p>
        </div>
    """, unsafe_allow_html=True)

with col_right:
    st.markdown("""
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow space-y-4">
            <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
                <span>📖</span> BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN
            </h2>
        </div>
    """, unsafe_allow_html=True)
    
    btn_fetch = st.button("Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn")
    
    st.markdown("""
        <div class="p-3.5 bg-blue-50 border border-blue-100 rounded-xl flex items-start gap-2.5 text-xs text-blue-900 my-3">
            <i class="fa-regular fa-lightbulb text-blue-600 text-base shrink-0 mt-0.5"></i>
            <span><strong>Gợi ý:</strong> Hãy chọn Môn học và Khối lớp trước để xem danh sách.</span>
        </div>
    """, unsafe_allow_html=True)

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
        selected_idx = st.selectbox("👉 Chọn Bài học chuẩn từ danh sách:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x])
        
        current_item = lessons_data[selected_idx]
        chapter_title = st.text_input("Tên Chương/Chủ đề", value=current_item['chapter'])
        lesson_title = st.text_input("Tên Bài dạy", value=current_item['lesson'])
        
        c_dur, c_req = st.columns([1, 2])
        with c_dur:
            duration = st.number_input("Số tiết", value=current_item['duration'])
        with c_req:
            requirements = st.text_input("Yêu cầu cần đạt", value=current_item['req'])
    else:
        chapter_title = st.text_input("Tên Chương/Chủ đề", placeholder="Nhập tên chương...")
        lesson_title = st.text_input("Tên Bài dạy", placeholder="Nhập tên bài dạy...")
        
        c_dur, c_req = st.columns([1, 2])
        with c_dur:
            duration = st.number_input("Số tiết", value=2)
        with c_req:
            requirements = st.text_input("Yêu cầu cần đạt", placeholder="Yêu cầu...")

# ==========================================
# KHỐI BƯỚC 3: TÍCH HỢP NĂNG LỰC
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
    <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow space-y-4">
        <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
            <span>🚀</span> BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
        </h2>
    </div>
""", unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn các yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Năng lực Hợp tác & Tự học",
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện"
    ],
    default=[
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"
    ]
)

# ==========================================
# NÚT BẮT ĐẦU TẠO KHBD WORD (GRADIENT CÓ TÊN CỦA THẦY)
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)

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
