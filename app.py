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
st.set_page_config(page_title="Hệ Thống Soạn KHBD", layout="wide", page_icon="📝")

# ==========================================
# CẤU HÌNH GIAO DIỆN CHUẨN TAILWIND CSS GỐC
# ==========================================
st.markdown("""
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif !important;
        }
        .stApp {
            background-color: #F8FAFC !important;
        }
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 64rem !important;
        }
        div[data-baseweb="input"] > div, div[data-baseweb="select"] > div {
            background-color: #FFFFFF !important;
            border-color: #E5E7EB !important;
            border-radius: 0.75rem !important;
        }
        div.stButton > button[kind="primary"] {
            background: linear-gradient(to right, #2563EB, #4F46E5) !important;
            color: white !important;
            font-weight: 700 !important;
            font-size: 1rem !important;
            border-radius: 1rem !important;
            padding: 0.875rem 2rem !important;
            border: none !important;
            box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3) !important;
            width: 100% !important;
        }
        div.stButton > button[kind="primary"]:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 20px 25px -5px rgba(37, 99, 235, 0.4) !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP API KEY
# ==========================================
with st.sidebar:
    st.subheader("🔑 Đăng Nhập System")
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Nhập API Key vào đây...")
    model_name = st.selectbox("Mô hình AI:", ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"])
    st.divider()
    st.subheader("👤 Thông tin Giáo viên")
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN - VẬT LÍ")
    teacher_name = st.text_input("Họ tên GV:", "Dương Tấn Tiến")

# ==========================================
# HEADER
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
# GRID BƯỚC 1 & BƯỚC 2
# ==========================================
col_left, col_right = st.columns(2, gap="medium")

with col_left:
    st.markdown("""
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm mb-4">
            <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2 mb-2">
                <span>📚</span> BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
            </h2>
        </div>
    """, unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        subject = st.selectbox("Môn học/Hoạt động GD", ["Vật lí", "Toán học", "Hóa học", "Ngữ văn", "Tiếng Anh", "Sinh học", "Tin học"])
    with c2:
        grade = st.selectbox("Khối lớp", ["Khối 10", "Khối 11", "Khối 12"])

    st.markdown("""
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm space-y-3 mt-4">
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
        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm mb-2">
            <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
                <span>📖</span> BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN
            </h2>
        </div>
    """, unsafe_allow_html=True)
    
    btn_fetch = st.button("Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True)
    
    st.markdown("""
        <div class="p-3.5 bg-blue-50 border border-blue-100 rounded-xl flex items-start gap-2.5 text-xs text-blue-900 my-2">
            <i class="fa-regular fa-lightbulb text-blue-600 text-base shrink-0 mt-0.5"></i>
            <span><strong>Gợi ý:</strong> Hãy chọn Môn học và Khối lớp trước để xem danh sách.</span>
        </div>
    """, unsafe_allow_html=True)

    if btn_fetch:
        if not api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh bên trái!")
        else:
            try:
                genai.configure(api_key=api_key)
                # Sử dụng response_mime_type để ép Gemini trả về đúng định dạng JSON
                model = genai.GenerativeModel(
                    model_name=model_name,
                    generation_config={"response_mime_type": "application/json"}
                )
                
                prompt_fetch = f"""
                Liệt kê danh sách bài học môn {subject} - {grade} (Chương trình GDPT 2018).
                Trả về một mảng JSON chính xác bao gồm các đối tượng có cấu trúc sau:
                [
                  {{
                    "chapter": "Tên chương",
                    "lesson": "Tên bài học",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt"
                  }}
                ]
                """
                with st.spinner("Đang tải danh mục bài học..."):
                    res = model.generate_content(prompt_fetch)
                    raw_text = res.text.strip()
                    
                    # Bóc tách chuỗi JSON bằng Regex để tránh lỗi ký tự thừa
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    if json_match:
                        clean_json = json_match.group(0)
                    else:
                        clean_json = raw_text

                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("Tải dữ liệu thành công!")
            except Exception as e:
                st.error(f"Lỗi truy xuất danh mục: {e}")

    if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
        lessons = st.session_state['fetched_lessons']
        titles = [f"{i.get('chapter', '')} - {i.get('lesson', '')}" for i in lessons]
        sel_idx = st.selectbox("👉 Chọn bài học chuẩn:", range(len(titles)), format_func=lambda x: titles[x])
        curr = lessons[sel_idx]
        chapter_title = st.text_input("Tên Chương/Chủ đề", value=curr.get('chapter', ''))
        lesson_title = st.text_input("Tên Bài dạy", value=curr.get('lesson', ''))
        c_dur, c_req = st.columns([1, 2])
        with c_dur: duration = st.number_input("Số tiết", value=int(curr.get('duration', 2)))
        with c_req: requirements = st.text_input("Yêu cầu cần đạt", value=curr.get('req', ''))
    else:
        chapter_title = st.text_input("Tên Chương/Chủ đề", placeholder="Nhập tên chương...")
        lesson_title = st.text_input("Tên Bài dạy", placeholder="Nhập tên bài dạy...")
        c_dur, c_req = st.columns([1, 2])
        with c_dur: duration = st.number_input("Số tiết", value=2)
        with c_req: requirements = st.text_input("Yêu cầu cần đạt", placeholder="Yêu cầu...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
    <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm mb-2">
        <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
            <span>🚀</span> BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
        </h2>
    </div>
""", unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn các yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)", 
        "Năng lực Hợp tác & Tự học", 
        "Giáo dục STEM / STEAM"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy học (Gemini, ChatGPT, Canva...)"]
)

# ==========================================
# XỬ LÝ XUẤT FILE WORD 5512
# ==========================================
def create_word_document(content_text):
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
        clean_line = line.strip().replace("**", "").replace("*", "")
        if not clean_line or clean_line.startswith("---"): continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        
        if any(clean_line.startswith(x) for x in ["I. ", "II. ", "III. ", "IV. ", "HOẠT ĐỘNG ", "TIẾT "]):
            run = p.add_run(clean_line)
            run.bold = True
            run.font.size = Pt(13)
        elif any(clean_line.startswith(x) for x in ["1. ", "2. ", "3. ", "a)", "b)", "c)", "d)", "Bước 1:", "Bước 2:", "Bước 3:", "Bước 4:"]):
            run = p.add_run(clean_line)
            run.bold = True
        else:
            p.add_run(clean_line)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

st.markdown("<br>", unsafe_allow_html=True)
submit = st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary")

if submit:
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập hoặc chọn Tên bài dạy!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            
            prompt = f"""
            Soạn Kế hoạch bài dạy chuẩn Công văn 5512/BGDĐT:
            Môn: {subject} ({grade}) | Bài: {lesson_title} | Số tiết: {duration}
            Yêu cầu cần đạt: {requirements}
            Yếu tố tích hợp: {', '.join(integrations)}
            Trình bày rõ các mục I. MỤC TIÊU, II. THIẾT BỊ, III. TIẾN TRÌNH DẠY HỌC (đủ 4 bước cho mỗi hoạt động).
            """
            
            with st.spinner("✨ AI đang tạo Kế hoạch bài dạy..."):
                res = model.generate_content(prompt)
                st.success("🎉 Tạo giáo án thành công!")
                doc_file = create_word_document(res.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.divider()
                st.markdown(res.text)
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
