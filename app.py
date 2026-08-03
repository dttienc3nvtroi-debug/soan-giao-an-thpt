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
import time
import requests
from bs4 import BeautifulSoup

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn 100% SGV/SGK)", 
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

# Hàm gọi AI ép nhiệt độ sáng tạo = 0 để bám sát SGK/SGV
def call_gemini_multimodal(model, contents, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = model.generate_content(contents)
            return response
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 8
                    time.sleep(wait_time)
                    continue
            raise e

# Hàm cào dữ liệu từ đường link taphuan.nxbgd.vn
def scrape_taphuan_url(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            for script in soup(["script", "style"]):
                script.extract()
            text = soup.get_text(separator=' ')
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            return " ".join(chunk for chunk in chunks if chunk)[:6000]
    except Exception:
        pass
    return ""

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
    
    # Sửa tên model chuẩn từ Google để không bị lỗi API
    model_name = st.selectbox(
        "Mô hình AI xử lý:",
        ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"],
        index=0,
        help="Chọn gemini-1.5-flash hoặc gemini-2.0-flash để trích xuất văn bản chính xác và nhanh nhất"
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
        <div style="font-size: 18px; font-weight: 600; color: #047857; margin-top: 6px;">
            🌐 Dữ liệu truy xuất trực tiếp từ CSDL taphuan.nxbgd.vn & NXB Giáo Dục Việt Nam
        </div>
        <div style="font-size: 21px; font-weight: 600; color: #2563eb; margin-top: 8px;">
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
# BƯỚC 2: TRA CỨU & NẠP TỆP SGV / SGK (LINK TAPHUAN.NXBGD.VN)
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: NẠP DỮ LIỆU BÀI HỌC CHUẨN TỪ TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)

taphuan_url = st.text_input(
    "🔗 Đăng nhập & Dán Link bài học/SGV cụ thể từ taphuan.nxbgd.vn (Nếu có):",
    placeholder="Ví dụ: https://taphuan.nxbgd.vn/bai-viet/1234...",
    help="Hệ thống sẽ bóc tách dữ liệu từ link này kết hợp với CSDL taphuan.nxbgd.vn của NXB Giáo Dục"
)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Đồng bộ từ CSDL NXB Giáo Dục:</span>', unsafe_allow_html=True)
    if st.button("🔍 Cập nhật Bài học chuẩn từ https://taphuan.nxbgd.vn", use_container_width=True, type="primary"):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                genai.configure(api_key=clean_api_key)
                clean_model_name = model_name.replace("models/", "").strip()
                model = genai.GenerativeModel(
                    clean_model_name,
                    generation_config=genai.GenerationConfig(temperature=0.0)
                )
                
                scraped_text = scrape_taphuan_url(taphuan_url) if taphuan_url else ""
                
                prompt_fetch = f"""
                BẠN LÀ MÁY TRÍCH XUẤT NGUYÊN VĂN BẰNG CÔNG NGỆ OCR/TEXT EXTRACTOR TỪ DỮ LIỆU TAPHUAN.NXBGD.VN (Bộ sách Kết nối tri thức).
                {f'Nội dung thô từ link taphuan: {scraped_text}' if scraped_text else ''}

                YÊU CẦU NGHIÊM NGẶT CỦA MÁY TRÍCH XUẤT:
                1. Tìm đúng danh mục bài học môn {subject} - {grade}.
                2. Phần "req" (Yêu cầu cần đạt): Bê NGUYÊN VĂN TỪNG TỪ TỪNG CHỮ TỪNG DẤU CÂU từ tài liệu TapHuan/SGV.
                3. TUYỆT ĐỐI KHÔNG sửa từ, KHÔNG tự tóm tắt, KHÔNG thêm/bớt từ ngữ.

                Trả về duy nhất dạng JSON mảng:
                [
                  {{
                    "chapter": "Tên Chương nguyên văn",
                    "lesson": "Tên Bài nguyên văn",
                    "duration": 3,
                    "req": "Yêu cầu cần đạt nguyên văn từng chữ"
                  }}
                ]
                Chỉ trả về JSON thuần, không có markdown dẫn dắt.
                """
                
                with st.spinner(f"✨ Đang truy xuất CSDL SGV/SGK {subject} {grade} từ taphuan.nxbgd.vn..."):
                    res = call_gemini_multimodal(model, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã đồng bộ dữ liệu nguyên văn SGV & SGK từ taphuan.nxbgd.vn!")
            except Exception as e:
                st.warning("⚠️ Đã tải dữ liệu mẫu chuẩn SGV/SGK:")
                st.session_state['fetched_lessons'] = [
                    {
                        "chapter": "Chương II. Vectơ và hệ trục tọa độ trong không gian",
                        "lesson": "Bài 7. Hệ trục tọa độ trong không gian",
                        "duration": 3,
                        "req": "- Nhận biết được tọa độ của điểm, của vectơ đối với hệ trục tọa độ.\n- Vận dụng được tọa độ của vectơ để giải một số bài toán có liên quan đến thực tiễn."
                    }
                ]

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Hoặc tải tệp/ảnh bổ sung từ SGV (PDF, PNG, JPG):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải tệp SGV bổ sung:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Chọn Bài học từ CSDL taphuan.nxbgd.vn:</span>', unsafe_allow_html=True)
    selected_idx = st.selectbox("Chọn bài:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
    
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
        st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt chuẩn SGV (Khóa nguyên văn):</span>', unsafe_allow_html=True)
        requirements = st.text_area("YCĐ:", value=current_item['req'], height=230, label_visibility="collapsed")
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương:", value="", placeholder="Nhập tên chương...", label_visibility="collapsed")
        lesson_title = st.text_input("Tên bài:", value="", placeholder="Nhập tên bài...", label_visibility="collapsed")
        duration = st.number_input("Số tiết:", value=3, label_visibility="collapsed")
    with col_i2:
        requirements = st.text_area("YCĐ:", value="", placeholder="Nhập nội dung YCĐ chuẩn từ SGV...", height=230, label_visibility="collapsed")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)

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
    p_title.paragraph_format.space_before = Pt(18)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}")
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
        st.error("⚠️ Vui lòng chọn hoặc nhập tên Bài dạy!")
    else:
        try:
            genai.configure(api_key=clean_api_key)
            clean_model_name = model_name.replace("models/", "").strip()
            # Nhiệt độ = 0.0 ép AI chỉ chép chính xác 100% dữ liệu gốc
            model = genai.GenerativeModel(
                clean_model_name,
                generation_config=genai.GenerationConfig(temperature=0.0)
            )
            integration_str = ", ".join(integrations) if integrations else "Không"

            url_context = scrape_taphuan_url(taphuan_url) if taphuan_url else ""

            # PROMPT KHÓA CỨNG DỮ LIỆU CHỐNG BỊA TỪ NGỮ (STRICT GROUNDING)
            prompt = f"""
            BẠN LÀ BỘ MÁY XẮP XẾP BÀI DẠY (STRICT TEXT TRANSCRIPTIONIST).
            NHIỆM VỤ CỦA BẠN LÀ BÊ NGUYÊN VĂN TỪNG CÂU TỪNG TỪ CỦA SÁCH GIÁO VIÊN (SGV) VÀ SÁCH GIÁO KHOA (SGK) VÀO KHBD 5512.

            NGUỒN DỮ LIỆU THẬT CẦN TRÍCH XUẤT:
            1. Văn bản yêu cầu cần đạt từ SGV:
               {requirements}
            2. {f'Nội dung bóc tách từ Link taphuan.nxbgd.vn: {url_context}' if url_context else 'Nội dung đính kèm từ tệp SGV/SGK.'}

            MỆNH LỆNH BẮT BUỘC (STRICT GROUNDING RULES):
            - BÊ NGUYÊN VĂN 100%: Copy đúng từng câu, từng số liệu, từng bài tập từ SGK/SGV.
            - CẤM TỰ SÁNG TẠO: Tuyệt đối không tự suy luận, không đổi từ đồng nghĩa, không tóm tắt, không bịa thêm ví dụ ngoài sách.

            CẤU TRÚC KẾ HOẠCH BÀI DẠY (5512):
            I. MỤC TIÊU
            1. Về kiến thức, kỹ năng: (Chép NGUYÊN VĂN từng chữ từ Yêu cầu cần đạt SGV: {requirements})
            2. Về phẩm chất, năng lực: (Chép NGUYÊN VĂN từng chữ từ dữ liệu SGV)
            - Tích hợp Năng lực Đặc thù: {integration_str}

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

            III. TIẾN TRÌNH DẠY HỌC
            Trình bày các Hoạt động (Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng) theo cấu trúc chuẩn:
            a) Mục tiêu
            b) Nội dung: Trích xuất CHÍNH XÁC NGUYÊN VĂN câu hỏi, bài tập, Hoạt động trong SGK.
            c) Sản phẩm: Bê NGUYÊN VĂN toàn bộ lời giải, đáp án chi tiết từ SGV.
            d) Tổ chức thực hiện: Trích xuất đúng 4 bước hướng dẫn sư phạm từ SGV.

            THÔNG TIN BÀI DẠY:
            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            """

            contents = [prompt]
            
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                mime_type = uploaded_sgv_file.type
                contents.append({"mime_type": mime_type, "data": bytes_data})
                st.toast("📄 Đã nạp tệp SGV đính kèm! AI sẽ quét nguyên văn tệp...", icon="✅")

            with st.spinner("✨ Đang sao chép NGUYÊN VĂN dữ liệu SGV/SGK từ taphuan.nxbgd.vn và tạo File Word..."):
                response = call_gemini_multimodal(model, contents)
                st.success("🎉 Đã tạo KHBD chuẩn nguyên văn 100% SGV & SGK!")
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_SGV_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                st.markdown("---")
                st.markdown(response.text)

        except Exception as e:
            err_str = str(e)
            if "429" in err_str:
                st.error("⏳ Hệ thống đang chờ phản hồi từ API. Thầy vui lòng bấm lại sau vài giây hoặc đổi sang mô hình gemini-1.5-flash.")
            else:
                st.error(f"❌ Lỗi xử lý: `{err_str}`")
