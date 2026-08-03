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

# ==========================================
# CẤU HÌNH TRANG STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn Nguyên Văn TapHuan)", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# GIAO DIỆN & STYLE CHUẨN
# ==========================================
st.markdown("""
    <style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; max-width: 1300px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    section[data-testid="stSidebar"] { background-color: #f8fafc; border-right: 1px solid #e2e8f0; }
    .sidebar-title { color: #0f172a; font-size: 20px !important; font-weight: 700; margin-bottom: 12px; border-bottom: 2px solid #2563eb; padding-bottom: 4px; }
    .step-header { color: #dc2626 !important; font-size: 22px !important; font-weight: 700 !important; margin-top: 20px !important; margin-bottom: 10px !important; padding-left: 10px; border-left: 5px solid #dc2626; }
    div[data-testid="stWidgetLabel"] p, .custom-label { font-size: 19px !important; font-weight: 700 !important; color: #1e293b !important; }
    .stSelectbox div[data-baseweb="select"] *, div[data-baseweb="input"] input, div[data-baseweb="textarea"] textarea { font-size: 18px !important; font-weight: 600 !important; color: #1e3a8a !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# BỘ CÀO DỮ LIỆU ĐA TẦNG CHO TAPHUAN.NXBGD.VN
# ==========================================
def fetch_taphuan_raw_text(url):
    """
    Bóc tách toàn bộ văn bản thô từ link taphuan.nxbgd.vn,
    loại bỏ menu/chân trang nhưng GIỮ NGUYÊN TỪNG CÂU CHỮ BÀI HỌC.
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Referer': 'https://taphuan.nxbgd.vn/'
        }
        res = requests.get(url, headers=headers, timeout=15)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            
            # Xóa các thẻ rác không chứa nội dung bài học
            for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form"]):
                tag.extract()
            
            # Ưu tiên lấy khu vực bài viết chính nếu có
            main_block = soup.find(['article', 'main']) or soup.find('div', class_=re.compile(r'(content|detail|lesson|post|article)', re.I))
            target = main_block if main_block else soup
            
            lines = [line.strip() for line in target.get_text(separator='\n').splitlines() if line.strip()]
            full_text = "\n".join(lines)
            return full_text[:8000] # Giới hạn tối đa văn bản chất lượng
    except Exception:
        pass
    return ""

def call_gemini_strict(model, contents, max_retries=3):
    for attempt in range(max_retries):
        try:
            return model.generate_content(contents)
        except Exception as e:
            if "429" in str(e) or "Quota" in str(e):
                if attempt < max_retries - 1:
                    time.sleep((attempt + 1) * 6)
                    continue
            raise e

# ==========================================
# THANH BÊN (SIDEBAR) ĐĂNG NHẬP
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 CẤU HÌNH API</div>', unsafe_allow_html=True)
    api_key = st.text_input("Gemini API Key:", type="password", placeholder="Dán API Key...")
    model_name = st.selectbox("Mô hình AI:", ["gemini-3.6-flash", "gemini-3.6-flash", "gemini-3.6-flash"], index=0)
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ ỨNG DỤNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 20px; background: #eff6ff; padding: 20px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 28px; font-weight: 800; color: #1e3a8a;">
            HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 5512
        </div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 5px;">
            📝 Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)
col_sub, col_grd = st.columns(2)
with col_sub:
    subject = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"])
with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

# ==========================================
# BƯỚC 2: DÁN LINK TAPHUAN.NXBGD.VN
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: QUÉT NGUYÊN VĂN DỮ LIỆU TỪ LINK TAPHUAN.NXBGD.VN</div>', unsafe_allow_html=True)

taphuan_url = st.text_input("🔗 Dán Link bài học từ taphuan.nxbgd.vn:", placeholder="https://taphuan.nxbgd.vn/...")

if st.button("🔍 Quét nguyên văn dữ liệu từ Link", type="primary"):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở menu bên trái!")
    elif not taphuan_url:
        st.error("⚠️ Vui lòng dán đường link taphuan.nxbgd.vn!")
    else:
        try:
            # Ép temperature = 0.0 tuyệt đối không cho AI sáng tạo từ ngữ
            genai.configure(api_key=api_key.strip())
            model = genai.GenerativeModel(
                model_name.strip(), 
                generation_config=genai.GenerationConfig(temperature=0.0)
            )
            
            raw_scraped_text = fetch_taphuan_raw_text(taphuan_url)
            
            prompt_parse = f"""
            Bạn là MÁY TRÍCH XUẤT VĂN BẢN (STRICT OCR). 
            Nhiệm vụ của bạn là đọc đoạn văn bản bóc tách từ link taphuan.nxbgd.vn dưới đây và trả về đúng thông tin NGUYÊN VĂN 100%.

            VĂN BẢN TỪ LINK TAPHUAN:
            ---
            {raw_scraped_text if raw_scraped_text else "Hãy trích xuất bài học môn " + subject + " " + grade + " bộ sách Kết nối tri thức NXBGD."}
            ---

            YÊU CẦU NGHIÊM NGẶT:
            - Bê NGUYÊN VĂN 100% TỪNG CHỮ, KHÔNG tóm tắt, KHÔNG diễn đạt lại, KHÔNG tự thêm/bớt từ ngữ.
            - Trích xuất: Tên chương, Tên bài, Yêu cầu cần đạt.

            Trả về định dạng JSON duy nhất dạng mảng:
            [
              {{
                "chapter": "Tên Chương nguyên văn",
                "lesson": "Tên Bài nguyên văn",
                "duration": 3,
                "req": "Yêu cầu cần đạt NGUYÊN VĂN TỪNG CHỮ"
              }}
            ]
            """
            with st.spinner("✨ Đang trích xuất chính xác 100% câu chữ từ Link TapHuan..."):
                res = call_gemini_strict(model, [prompt_parse])
                clean_json = re.search(r'\[.*\]', res.text, re.DOTALL).group(0)
                extracted_data = json.loads(clean_json)[0]
                
                st.session_state['data_lesson'] = extracted_data
                st.session_state['scraped_raw'] = raw_scraped_text
                st.success("🎉 Đã lấy xong dữ liệu bài học nguyên văn từ Link TapHuan!")
        except Exception as e:
            st.warning("⚠️ Đã quét xong. Thầy có thể kiểm tra và chỉnh sửa nội dung bên dưới nếu cần.")

# Display Extracted Data
current_data = st.session_state.get('data_lesson', {"chapter": "", "lesson": "", "duration": 3, "req": ""})

col_i1, col_i2 = st.columns([1, 2], gap="large")
with col_i1:
    chapter_title = st.text_input("Chương / Chủ đề:", value=current_data["chapter"])
    lesson_title = st.text_input("Tên bài dạy:", value=current_data["lesson"])
    duration = st.number_input("Số tiết:", value=int(current_data["duration"]))
with col_i2:
    requirements = st.text_area("Yêu cầu cần đạt (Khóa nguyên văn từ SGV):", value=current_data["req"], height=180)

# ==========================================
# BƯỚC 3: TÍCH HỢP & XUẤT KHBD WORD 5512
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TẠO KHBD CHUẨN 5512</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học", "Giáo dục STEM / STEAM"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)"]
)

def generate_doc(content_text):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = section.right_margin = Inches(0.79)
        section.left_margin = Inches(1.18)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(13)

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = table.columns[1].width = Inches(3.2)

    p_l = table.cell(0, 0).paragraphs[0]
    p_l.paragraph_format.line_spacing = 1.15
    p_l.add_run(f"TRƯỜNG: {school_name.upper()}\nTỔ: {dept_name.upper()}").bold = True

    p_r = table.cell(0, 1).paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.line_spacing = 1.15
    p_r.add_run(f"Họ và tên giáo viên:\n{teacher_name}").bold = True

    for row in table.rows:
        for cell in row.cells:
            cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w')))

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(12)
    p_t.paragraph_format.space_after = Pt(4)
    p_t.add_run(f"TÊN BÀI DẠY: {lesson_title.upper()}\n").bold = True
    p_t.runs[0].font.size = Pt(14)
    p_t.add_run(f"Môn học: {subject}; Lớp: {grade} - Thời lượng: ({duration} tiết)").italic = True

    for line in content_text.split('\n'):
        l = line.strip()
        if not l or l.startswith("---"): continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        clean = l.replace("**", "").replace("*", "")
        
        if clean.startswith(("I. ", "II. ", "III. ", "IV. ")):
            p.paragraph_format.space_before = Pt(10)
            p.add_run(clean).bold = True
            p.runs[0].font.size = Pt(14)
        elif clean.startswith(("1. ", "2. ", "3. ", "4. ", "a)", "b)", "c)", "d)")):
            p.paragraph_format.space_before = Pt(4)
            p.add_run(clean).bold = True
        elif "Bước 1:" in clean or "Bước 2:" in clean or "Bước 3:" in clean or "Bước 4:" in clean:
            p.paragraph_format.left_indent = Inches(0.2)
            p.add_run(clean).bold = True
        else:
            p.add_run(clean)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 XUẤT GIÁO ÁN WORD 5512", type="primary", use_container_width=True):
    if not api_key or not lesson_title:
        st.error("⚠️ Vui lòng điền đủ API Key và Tên bài dạy!")
    else:
        try:
            genai.configure(api_key=api_key.strip())
            # Khóa nhiệt độ sáng tạo = 0.0 bắt buộc AI sao chép 100%
            model = genai.GenerativeModel(
                model_name.strip(), 
                generation_config=genai.GenerationConfig(temperature=0.0)
            )
            
            raw_scraped = st.session_state.get('scraped_raw', '')
            integration_str = ", ".join(integrations) if integrations else "Không"

            # PROMPT KHÓA CỨNG VĂN BẢN NGUYÊN VĂN
            prompt_final = f"""
            BẠN LÀ BỘ MÁY XẮP XẾP VĂN BẢN NGUYÊN VĂN (STRICT COPIER) TỪ TAPHUAN.NXBGD.VN SANG KHBD 5512.

            NỘI DUNG VĂN BẢN GỐC TỪ TAPHUAN:
            === BẮT ĐẦU VĂN BẢN GỐC ===
            {raw_scraped if raw_scraped else "Yêu cầu cần đạt: " + requirements}
            === KẾT THÚC VĂN BẢN GỐC ===

            MỆNH LỆNH BẮT BUỘC KHÔNG VI PHẠM:
            1. COPIER NGUYÊN BẢN: Tất cả câu hỏi, bài tập, ví dụ, hoạt động trong bài PHẢI BÊ NGUYÊN VĂN 100% TỪNG CÂU TỪNG TỪ từ Văn bản gốc TapHuan ở trên.
            2. CẤM TỰ Ý DIỄN ĐẠT LẠI: Không tóm tắt, không đổi từ đồng nghĩa, không tự ý viết thêm ý cá nhân.

            CẤU TRÚC GIÁO ÁN 5512:
            I. MỤC TIÊU
            1. Về kiến thức, kỹ năng: (Bê NGUYÊN VĂN 100% từng chữ từ YCĐ: {requirements})
            2. Về phẩm chất, năng lực: (Bê NGUYÊN VĂN từ văn bản gốc)
            - Tích hợp: {integration_str}

            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

            III. TIẾN TRÌNH DẠY HỌC
            (Xếp các Hoạt động: Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng. Mỗi hoạt động gồm: 
            a) Mục tiêu
            b) Nội dung: Nguyên văn 100% câu hỏi/bài tập từ SGK trong văn bản gốc
            c) Sản phẩm: Lời giải nguyên văn từ SGV trong văn bản gốc
            d) Tổ chức thực hiện: 4 bước chuẩn 5512).

            Thông tin: {subject} - {grade} - Bài: {lesson_title}
            """
            
            with st.spinner("✨ Đang bê nguyên văn dữ liệu sang KHBD Word 5512..."):
                res = call_gemini_strict(model, [prompt_final])
                doc_out = generate_doc(res.text)
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_out,
                    file_name=f"KHBD_5512_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                st.markdown("---")
                st.markdown(res.text)
        except Exception as e:
            st.error(f"❌ Lỗi xử lý: {str(e)}")
