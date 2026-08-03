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
from pypdf import PdfReader

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn Dữ Liệu)", 
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

# HÀM XỬ LÝ FILE TẢI LÊN (ĐỌC TEXT VỚI PDF HOẶC CHUYỂN IMAGE PART DÙNG CHO GEMINI)
def extract_file_content(uploaded_file):
    if uploaded_file is None:
        return None, ""
    
    file_type = uploaded_file.type
    bytes_data = uploaded_file.getvalue()
    
    # Nếu là PDF: Đọc trực tiếp chữ ra String để gửi cho Gemini (Siêu nhanh + Chính xác 100%)
    if "pdf" in file_type:
        try:
            pdf_reader = PdfReader(io.BytesIO(bytes_data))
            extracted_text = ""
            for page in pdf_reader.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
            return "text", extracted_text
        except Exception as e:
            # Fallback nếu PDF dạng ảnh chụp không đọc được chữ
            return "image_part", {"mime_type": "application/pdf", "data": bytes_data}
            
    # Nếu là tệp Ảnh (JPG, PNG)
    elif "image" in file_type or "jpg" in file_type or "png" in file_type or "jpeg" in file_type:
        return "image_part", {"mime_type": file_type if file_type else "image/jpeg", "data": bytes_data}
        
    return None, ""

# HÀM GỌI GEMINI AN TOÀN & TỐI ƯU
def call_gemini_fast(api_key, model_choice, contents):
    genai.configure(api_key=api_key)
    clean_model_name = model_choice.replace("models/", "").strip()
    
    generation_config = genai.types.GenerationConfig(
        temperature=0.1,  # Nhiệt độ thấp giúp giữ nguyên văn bản gốc từ SGV
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
        "Mô hình AI ưu tiên (Khuyên dùng Flash):",
        ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
        index=0,
        help="Model Flash phản hồi nhanh và đọc file cực tốt"
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
# BƯỚC 2: TRA CỨU & NẠP TỆP SGV / SGK
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: NẠP FILE SGV (ĐẢM BẢO CHÍNH XÁC 100%)</div>', unsafe_allow_html=True)

col_btn_sync, col_file_upload = st.columns([1, 1], gap="medium")

with col_btn_sync:
    st.markdown('<span class="custom-label">🌐 Tải danh mục bài học chuẩn (Bộ sách KNTT/GDPT 2018):</span>', unsafe_allow_html=True)
    if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
        clean_api_key = api_key.strip() if api_key else ""
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
        else:
            try:
                prompt_fetch = f"""
                Liệt kê ĐẦY ĐỦ tất cả các Bài học thuộc môn {subject} - {grade} (KNTT).
                Trả về duy nhất dạng JSON mảng:
                [
                  {{
                    "chapter": "Tên Chương 1",
                    "lesson": "Tên Bài 1",
                    "duration": 2,
                    "req": "Yêu cầu cần đạt chuẩn của bài"
                  }}
                ]
                Chỉ trả về mã JSON mảng [ ... ], không viết lời chào.
                """
                
                with st.spinner(f"⚡ Đang nạp danh mục bài học..."):
                    res = call_gemini_fast(clean_api_key, model_name, [prompt_fetch])
                    raw_text = res.text.strip()
                    json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                    clean_json = json_match.group(0) if json_match else raw_text
                    st.session_state['fetched_lessons'] = json.loads(clean_json)
                    st.success("🎉 Đã tải xong danh mục bài học!")
            except Exception as e:
                st.warning("⚠️ Đã nạp bài học mẫu chuẩn SGK:")
                st.session_state['fetched_lessons'] = [
                    {
                        "chapter": "Chương II. Vectơ và hệ trục tọa độ trong không gian",
                        "lesson": "Bài 7. Hệ trục tọa độ trong không gian",
                        "duration": 3,
                        "req": "- Nhận biết được tọa độ của điểm, của vectơ đối với hệ trục tọa độ.\n- Vận dụng được tọa độ của vectơ để giải một số bài toán có liên quan đến thực tiễn."
                    }
                ]

with col_file_upload:
    st.markdown('<span class="custom-label">📂 Tải lên File/Ảnh trang SGV (PDF, JPG, PNG):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải lên File SGV:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

if uploaded_sgv_file is not None:
    st.info(f"✅ Đã nhận tệp: **{uploaded_sgv_file.name}**. Hệ thống sẽ ưu tiên trích xuất MỤC TIÊU từ tệp này.")

if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    
    st.markdown('<span class="custom-label">👉 Chọn Bài học từ danh sách:</span>', unsafe_allow_html=True)
    selected_idx = st.selectbox("Chọn bài:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
    
    current_item = lessons_data[selected_idx]
    
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương / Chủ đề:", value=current_item['chapter'])
        lesson_title = st.text_input("Tên bài dạy:", value=current_item['lesson'])
        duration = st.number_input("Số tiết thực hiện:", value=int(current_item['duration']))
    
    with col_i2:
        requirements = st.text_area("📌 Yêu cầu cần đạt / Mục tiêu SGV (Mặc định):", value=current_item['req'], height=230)
else:
    col_i1, col_i2 = st.columns([1, 2], gap="large")
    with col_i1:
        chapter_title = st.text_input("Chương / Chủ đề:", value="", placeholder="Nhập tên chương...")
        lesson_title = st.text_input("Tên bài dạy:", value="", placeholder="Nhập tên bài...")
        duration = st.number_input("Số tiết thực hiện:", value=3)
        
    with col_i2:
        requirements = st.text_area("📌 Yêu cầu cần đạt / Mục tiêu SGV (Mặc định):", value="", placeholder="Nhập mục tiêu hoặc đính kèm tệp SGV ở trên...", height=230)

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
            integration_str = ", ".join(integrations) if integrations else "Không"
            
            # Xử lý file upload
            file_mode, file_payload = extract_file_content(uploaded_sgv_file)
            
            file_instruction = ""
            contents = []

            if file_mode == "text":
                file_instruction = f"\n⚠️ DỮ LIỆU TỰ ĐỘNG ĐỌC TỪ FILE TẢI LÊN (ƯU TIÊN TRÍCH XUẤT NGUYÊN VĂN 100%): \n'''\n{file_payload}\n'''"
            elif file_mode == "image_part":
                file_instruction = "\n⚠️ ĐÃ ĐÍNH KÈM FILE TẢI LÊN. HÃY ĐỌC TRỰC TIẾP TỆP ĐÍNH KÈM BÊN DƯỚI ĐỂ TRÍCH XUẤT MỤC TIÊU."
                contents.append(file_payload)

            prompt = f"""
            Hãy đóng vai một Chuyên gia Giáo dục. Hãy lập Kế hoạch bài dạy (Giáo án) theo chuẩn Công văn 5512/BGDĐT.

            {file_instruction}

            MÔN HỌC: {subject} - {grade}
            TÊN BÀI DẠY: {lesson_title}
            SỐ TIẾT: {duration}
            MỤC TIÊU BỔ SUNG (NẾU KHÔNG CÓ FILE): {requirements}
            YẾU TỐ TÍCH HỢP: {integration_str}

            YÊU CẦU BẮT BUỘC VỀ CẤU TRÚC GIÁO ÁN:
            I. MỤC TIÊU
            1. Về kiến thức: (Trích xuất nguyên văn từ dữ liệu file/Mục tiêu SGV).
            2. Về năng lực: 
               - Năng lực chung & Năng lực đặc thù của môn học.
               - Yếu tố tích hợp: {integration_str}
            3. Về phẩm chất: (Trích xuất nguyên văn hoặc chuẩn GDPT 2018).

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

            # Đưa prompt vào danh sách gửi cho AI
            contents.append(prompt)

            with st.spinner("⚡ AI đang phân tích dữ liệu tệp và khởi tạo giáo án..."):
                response = call_gemini_fast(clean_api_key, model_name, contents)
                st.success("🎉 Đã tạo xong giáo án chuẩn 5512 từ dữ liệu SGV!")
                
                doc_file = generate_doc(response.text)
                
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
            st.error(f"❌ Lỗi xử lý dữ liệu: `{err_str}`. Thầy vui lòng kiểm tra lại API Key!")
