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
        font-size: 22px !important;
        font-weight: 700;
        margin-bottom: 14px;
        border-bottom: 3px solid #2563eb;
        padding-bottom: 6px;
    }
    .step-header {
        color: #dc2626 !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        margin-top: 24px !important;
        margin-bottom: 14px !important;
        padding-left: 10px;
        border-left: 5px solid #dc2626;
    }
    div[data-testid="stWidgetLabel"] p, .custom-label {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 6px !important;
    }
    .stSelectbox div[data-baseweb="select"] *,
    .stSelectbox [data-testid="stMarkdownContainer"] p,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="input"] input, .stTextInput input, .stNumberInput input {
        font-size: 21px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    div[data-baseweb="textarea"] textarea, .stTextArea textarea {
        font-size: 21px !important;
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
    .stButton button p {
        font-size: 21px !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CƠ SỞ DỮ LIỆU BÀI HỌC CHUẨN 100% NGUYÊN VĂN (MẪU ĐIỂN HÌNH)
# ==========================================
DEFAULT_LESSONS_DATABASE = {
    "Toán học - Lớp 12": [
        {
            "chapter": "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị của hàm số",
            "lesson": "Bài 1. Tính đơn điệu và cực trị của hàm số",
            "duration": 3,
            "req": "- Nhận biết được tính đồng biến, nghịch biến của một hàm số trên một khoảng thông qua dấu của đạo hàm cấp một.\n- Thể hiện được tính đồng biến, nghịch biến của hàm số trên bảng biến thiên.\n- Nhận biết được điểm cực trị, giá trị cực trị của hàm số."
        },
        {
            "chapter": "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị của hàm số",
            "lesson": "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số",
            "duration": 2,
            "req": "- Nhận biết được giá trị lớn nhất, giá trị nhỏ nhất của hàm số trên một tập xác định cho trước.\n- Tìm được giá trị lớn nhất, giá trị nhỏ nhất của hàm số đơn giản bằng đạo hàm."
        },
        {
            "chapter": "Chương II. Vectơ và hệ trục tọa độ trong không gian",
            "lesson": "Bài 7. Hệ trục tọa độ trong không gian",
            "duration": 3,
            "req": "- Nhận biết được hệ trục tọa độ Oxyz trong không gian.\n- Nhận biết được tọa độ của một điểm, tọa độ của một vectơ đối với hệ trục tọa độ Oxyz.\n- Tính được tọa độ của tổng, hiệu hai vectơ, tích của một số với một vectơ.\n- Vận dụng được biểu thức tọa độ của các phép toán vectơ để giải quyết một số bài toán thực tế."
        },
        {
            "chapter": "Chương II. Vectơ và hệ trục tọa độ trong không gian",
            "lesson": "Bài 8. Biểu thức tọa độ của các phép toán vectơ",
            "duration": 2,
            "req": "- Tự định nghĩa và tính toán biểu thức tọa độ phép cộng, trừ, nhân vô hướng của hai vectơ trong không gian."
        }
    ],
    "Toán học - Lớp 11": [
        {
            "chapter": "Chương I. Hàm số lượng giác và phương trình lượng giác",
            "lesson": "Bài 1. Giá trị lượng giác của góc lượng giác",
            "duration": 3,
            "req": "- Nhận biết được các khái niệm về góc lượng giác, đơn vị đo góc rad.\n- Tính được các giá trị lượng giác của một góc lượng giác."
        }
    ]
}

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
        
    return {
        "mime_type": mime,
        "data": bytes_data
    }

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
    
    api_key = st.text_input(
        "Google Gemini API Key:", 
        type="password", 
        placeholder="Dán mã API Key vào đây...",
        help="Nhập API Key từ Google AI Studio"
    )
    
    model_name = st.selectbox(
        "Mô hình AI ưu tiên:",
        ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
        index=0,
        help="Mô hình Flash xử lý đọc file và tạo giáo án nhanh nhất"
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
    <div style="text-align: center; margin-bottom: 25px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 22px; border-radius: 12px; border: 1px solid #bfdbfe;">
        <div style="font-size: 33px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 100% SGV (5512)
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
# BƯỚC 2: NẠP DỮ LIỆU & TỆP SGV / SGK
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU HOẶC TẢI LÊN SGV (CHÍNH XÁC 100%)</div>', unsafe_allow_html=True)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Tải danh mục bài học đầy đủ (Bộ sách KNTT/GDPT 2018):</span>', unsafe_allow_html=True)
    if st.button("🔍 Tra cứu toàn bộ Danh mục bài học từ AI", use_container_width=True):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                prompt_fetch = f"""
                Hãy tra cứu và trả về danh sách TẤT CẢ các bài học thuộc môn {subject} - {grade} (Bộ sách Kết nối tri thức với cuộc sống).
                YÊU CẦU ĐẶC BIỆT: Tên Chương và Tên Bài học PHẢI GHI ĐẦY ĐỦ 100% NGUYÊN VĂN THEO SGK, KHÔNG DÙNG TỪ TÓM TẮT.
                Ví dụ đúng:
                - chapter: "Chương II. Vectơ và hệ trục tọa độ trong không gian"
                - lesson: "Bài 7. Hệ trục tọa độ trong không gian"

                Trả về duy nhất mã JSON mảng dạng:
                [
                  {{
                    "chapter": "Tên chương nguyên văn đầy đủ",
                    "lesson": "Tên bài học nguyên văn đầy đủ",
                    "duration": 3,
                    "req": "Yêu cầu cần đạt nguyên văn SGV"
                  }}
                ]
                Không viết thêm lời chào hay giải thích nào khác ngoài mã JSON.
                """
                
                with st.spinner(f"⚡ Đang truy xuất danh mục bài học nguyên văn SGK..."):
                    res = call_gemini(clean_api_key, model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã tải xong danh mục đầy đủ bài học!")
            except Exception as e:
                db_key = f"{subject} - {grade}"
                st.session_state['fetched_lessons'] = DEFAULT_LESSONS_DATABASE.get(db_key, DEFAULT_LESSONS_DATABASE["Toán học - Lớp 12"])
                st.info("📌 Đã tải dữ liệu danh mục bài học chuẩn SGK từ bộ lưu trữ!")

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tải lên File/Ảnh trang SGV/SGK (PDF, JPG, PNG):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải lên File SGV:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

if uploaded_sgv_file is not None:
    st.info(f"✅ Đã nhận tệp: **{uploaded_sgv_file.name}**. AI sẽ trích xuất 100% Tiêu đề bài học và Mục tiêu từ tệp này.")

# Tải dữ liệu mặc định nếu chưa bấm tra cứu
key_pair = f"{subject} - {grade}"
if 'fetched_lessons' not in st.session_state:
    st.session_state['fetched_lessons'] = DEFAULT_LESSONS_DATABASE.get(key_pair, DEFAULT_LESSONS_DATABASE["Toán học - Lớp 12"])

lessons_data = st.session_state['fetched_lessons']
lesson_titles = [f"{item['chapter']} ➔ {item['lesson']}" for item in lessons_data]

st.markdown('<span class="custom-label">👉 Chọn Bài học từ danh mục SGK:</span>', unsafe_allow_html=True)
selected_idx = st.selectbox("Chọn bài:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")

current_item = lessons_data[selected_idx]

col_i1, col_i2 = st.columns([1, 2], gap="large")
with col_i1:
    chapter_title = st.text_input("Tên Chương / Chủ đề (Đầy đủ):", value=current_item['chapter'])
    lesson_title = st.text_input("Tên Bài dạy (Nguên văn SGK - Thầy có thể sửa trực tiếp):", value=current_item['lesson'])
    duration = st.number_input("Số tiết thực hiện:", value=int(current_item.get('duration', 3)))

with col_i2:
    requirements = st.text_area("📌 Yêu cầu cần đạt / Mục tiêu SGV (Nguyên văn):", value=current_item.get('req', ''), height=230)

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ (BỔ SUNG VÀO CẤU TRÚC 5512)</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    label_visibility="collapsed"
)

# ==========================================
# XỬ LÝ XUẤT FILE WORD 5512
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

    # In Tên Chương nguyên văn
    if full_chapter:
        p_chap = doc.add_paragraph()
        p_chap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_chap.paragraph_format.space_before = Pt(16)
        p_chap.paragraph_format.space_after = Pt(2)
        r_chap = p_chap.add_run(f"{full_chapter.upper()}")
        r_chap.bold = True
        r_chap.font.size = Pt(13)

    # In Tên Bài học nguyên văn
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

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary", use_container_width=True):
    clean_api_key = api_key.strip() if api_key else ""
    if not clean_api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh menu bên trái!")
    elif not lesson_title:
        st.error("⚠️ Vui lòng nhập hoặc chọn Tên bài dạy!")
    else:
        try:
            integration_str = ", ".join(integrations) if integrations else "Không"
            
            contents = []
            
            file_part = process_uploaded_file(uploaded_sgv_file)
            if file_part is not None:
                contents.append(file_part)
                file_note = f"⚠️ ĐÃ ĐÍNH KÈM TỆP SGV/SGK GỐC. HÃY ĐỌC VÀ TRÍCH XUẤT ĐÚNG NGUYÊN VĂN TÊN CHƯƠNG VÀ TÊN BÀI. NẾU TRÊN FILE CÓ TÊN BÀI HỌC CỤ THỂ THÌ PHẢI DÙNG CHÍNH XÁC TÊN ĐÓ."
            else:
                file_note = f"MỤC TIÊU/YÊU CẦU CẦN ĐẠT SGV: {requirements}"

            prompt = f"""
            Bạn là Chuyên gia Giáo dục. Hãy biên soạn Kế hoạch bài dạy (Giáo án) theo chuẩn Công văn 5512/BGDĐT.

            {file_note}

            MÔN HỌC: {subject} - {grade}
            CHƯƠNG / CHỦ ĐỀ NGUYÊN VĂN: {chapter_title}
            TÊN BÀI DẠY NGUYÊN VĂN 100% KHÔNG ĐƯỢC TÓM TẮT: {lesson_title}
            SỐ TIẾT: {duration}
            YẾU TỐ TÍCH HỢP: {integration_str}

            YÊU CẦU BẮT BUỘC VỀ CẤU TRÚC GIÁO ÁN:
            I. MỤC TIÊU
            1. Về kiến thức: (Trích xuất NGUYÊN VĂN từng gạch đầu dòng từ tệp đính kèm hoặc mục tiêu trên).
            2. Về năng lực: 
               - Năng lực chung & Năng lực đặc thù môn học.
               - Yếu tố tích hợp: {integration_str}
            3. Về phẩm chất: (Trích xuất NGUYÊN VĂN từ SGV).

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            - GV: ...
            - HS: ...

            III. TIẾN TRÌNH DẠY HỌC
            Trình bày chi tiết các Hoạt động (Khởi động, Khám phá/Hình thành kiến thức, Luyện tập, Vận dụng).
            Mỗi Hoạt động PHẢI có đủ 4 phần:
            a) Mục tiêu
            b) Nội dung
            c) Sản phẩm
            d) Tổ chức thực hiện:
               - Bước 1: Chuyển giao nhiệm vụ
               - Bước 2: Thực hiện nhiệm vụ
               - Bước 3: Báo cáo, thảo luận
               - Bước 4: Kết luận, nhận định
            """

            contents.append(prompt)

            with st.spinner("⚡ AI đang phân tích dữ liệu SGV và khởi tạo giáo án 5512..."):
                response = call_gemini(clean_api_key, model_name, contents)
                st.success("🎉 Đã tạo xong giáo án chuẩn 5512!")
                
                # Ép buộc truyền chính xác tên chương và tên bài do giáo viên nhập/chọn vào file Word
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
            err_str = str(e)
            st.error(f"❌ Lỗi xử lý: `{err_str}`. Thầy vui lòng kiểm tra lại API Key!")
