import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io
import json
import re
from PIL import Image

# ==========================================
# 1. CẤU HÌNH TRANG & CSS TỐI ƯU GIAO DIỆN
# ==========================================
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# CSS Điều chỉnh giao diện chính xác theo hình ảnh mẫu của thầy
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@700&display=swap');

    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1250px;
    }

    /* Tiêu đề các Bước (BƯỚC 1, BƯỚC 2, BƯỚC 3) */
    .step-header {
        color: #dc2626 !important;
        font-size: 21px !important;
        font-weight: 800 !important;
        font-family: 'Roboto Condensed', sans-serif, Arial !important;
        margin-top: 15px !important;
        margin-bottom: 12px !important;
        padding-left: 8px;
        border-left: 5px solid #dc2626;
        display: flex;
        align-items: center;
    }

    /* Nhãn chữ màu ĐEN, ĐẬM, NÉT HẸP chuẩn theo ảnh mẫu */
    .field-label {
        font-family: 'Roboto Condensed', sans-serif, Arial !important;
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #000000 !important;
        margin-bottom: 4px !important;
        display: block;
        white-space: nowrap;
    }

    /* Chữ bên trong ô nhập liệu/dropdown có màu XANH DƯƠNG ĐẬM (Bold Navy) */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p,
    div[data-baseweb="input"] input,
    .stTextInput input,
    .stNumberInput input,
    .stTextArea textarea {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }

    /* Thu gọn File Uploader để nằm gọn gàng bên trái */
    div[data-testid="stFileUploader"] section {
        padding: 6px 12px !important;
    }
    div[data-testid="stFileUploader"] section > div {
        padding: 0px !important;
    }
    
    .element-container {
        margin-bottom: 6px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown("### 🔑 ĐĂNG NHẬP & CẤU HÌNH")
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán mã API Key vào đây...")
    model_name = st.selectbox("Mô hình AI xử lý:", ["gemini-3.6-flash"], index=0)
    
    st.markdown("---")
    st.markdown("### 👤 THÔNG TIN GIÁO VIÊN")
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# 3. BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)
col_sub, col_grd = st.columns(2)
with col_sub:
    st.markdown('<span class="field-label">Môn học/Hoạt động GD:</span>', unsafe_allow_html=True)
    subject = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học"], label_visibility="collapsed")
with col_grd:
    st.markdown('<span class="field-label">Khối lớp:</span>', unsafe_allow_html=True)
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"], index=2, label_visibility="collapsed")

# ==========================================
# 4. BƯỚC 2: ĐIỀU CHỈNH GIAO DIỆN CHUẨN XÁC
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

# HÀNG 1: Nút bấm Cập nhật full chiều rộng trên cùng
if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            # Sửa cặp {{ }} để tránh lỗi Invalid format specifier
            prompt_fetch = f"""Liệt kê ĐẦY ĐỦ các Bài học thuộc môn {subject} - {grade} (GDPT 2018) dạng JSON array: [{{"chapter": "...", "lesson": "...", "duration": 3, "req": "..."}}]"""
            
            with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade}..."):
                res = model.generate_content(prompt_fetch)
                raw_text = res.text.strip()
                json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                clean_json = json_match.group(0) if json_match else raw_text
                st.session_state['fetched_lessons'] = json.loads(clean_json)
                st.success("🎉 Đã tải xong danh mục bài học chuẩn đầy đủ!")
        except Exception as e:
            st.error(f"Lỗi khi tải danh mục bài học: {e}")

st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# HÀNG 2: Chia 2 Cột nằm ngang (Trái: Tải File SGV | Phải: Chọn Bài học)
col_b2_left, col_b2_right = st.columns(2)

with col_b2_left:
    st.markdown('<span class="field-label">📁 Tải lên File SGV (Sách Giáo Viên - PDF hoặc Ảnh):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Upload SGV", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

with col_b2_right:
    st.markdown('<span class="field-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
    
    val_chap = "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số"
    val_less = "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số"
    val_dur = 3
    val_req = ""

    if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
        lessons_data = st.session_state['fetched_lessons']
        lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
        selected_idx = st.selectbox("Select lesson", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
        current_item = lessons_data[selected_idx]
        val_chap = current_item['chapter']
        val_less = current_item['lesson']
        val_dur = int(current_item['duration'])
        val_req = current_item['req']
    else:
        st.selectbox("Select lesson default", ["Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số - Bài 1. Tính đơn điệu và cực trị của hàm số"], label_visibility="collapsed")

st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# HÀNG 3: Chia 3 Cột (Chương | Tên bài | Số tiết) tỉ lệ chuẩn
c_chap, c_less, c_dur = st.columns([4.2, 4.2, 1.6])

with c_chap:
    st.markdown('<span class="field-label">Chương:</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương", value=val_chap, label_visibility="collapsed")

with c_less:
    st.markdown('<span class="field-label">Tên bài:</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài", value=val_less, label_visibility="collapsed")

with c_dur:
    st.markdown('<span class="field-label">Số tiết:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết", value=val_dur, min_value=1, label_visibility="collapsed")

st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# HÀNG 4: Yêu cầu cần đạt
st.markdown('<span class="field-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
requirements = st.text_area("YCĐ", value=val_req, height=100, label_visibility="collapsed")

# ==========================================
# 5. BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)

# ==========================================
# 6. HÀM TẠO FILE DOCX CÔNG VĂN 5512 (LOGIC GỐC)
# ==========================================
def set_cell_background(cell, fill_color):
    """Đặt màu nền cho cell trong bảng docx"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def create_5512_docx(data):
    """Tạo file Docx chuẩn Công văn 5512"""
    doc = docx.Document()

    # Cấu hình Lề trang chuẩn (Top/Bottom 2cm, Left 3cm, Right 1.5cm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)

    # Thẻ Style mặc định Times New Roman 13pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    # 1. BẢNG THÔNG TIN ĐẦU TRANG (Trường - GV - Môn - Bài)
    tbl_header = doc.add_table(rows=2, cols=2)
    tbl_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_header.autofit = False

    cell_00 = tbl_header.cell(0, 0)
    p00 = cell_00.paragraphs[0]
    p00.add_run(f"TRƯỜNG: {school_name.upper()}\nTỔ: {dept_name.upper()}").bold = True
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER

    cell_01 = tbl_header.cell(0, 1)
    p01 = cell_01.paragraphs[0]
    p01.add_run(f"Họ và tên giáo viên:\n{teacher_name}").bold = True
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tiêu đề bài học
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"\nTÊN BÀI DẠY: {data.get('lesson_title', '').upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Khối lớp: {grade}\nThời lượng thực hiện: ({data.get('duration', 1)} tiết)\n")

    # 2. NỘI DUNG GIÁO ÁN
    sections_5512 = [
        ("I. MỤC TIÊU", data.get("muc_tieu", "")),
        ("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU", data.get("thiet_bi", "")),
        ("III. TIẾN TRÌNH DẠY HỌC", "")
    ]

    for title, content in sections_5512:
        p_sec = doc.add_paragraph()
        r_sec = p_sec.add_run(title)
        r_sec.bold = True
        r_sec.font.size = Pt(13)
        if content:
            doc.add_paragraph(content)

    # 3. BẢNG TIẾN TRÌNH HOẠT ĐỘNG DẠY HỌC
    activities = data.get("hoat_dong", [])
    for act in activities:
        p_act = doc.add_paragraph()
        r_act = p_act.add_run(f"\n{act.get('ten_hoat_dong', 'Hoạt động')}")
        r_act.bold = True

        # Bảng 2 cột (Hoạt động của GV và HS | Nội dung cần đạt)
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header bảng
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "HOẠT ĐỘNG CỦA GV VÀ HS"
        hdr_cells[1].text = "SẢN PHẨM DỰ KIẾN / NỘI DUNG"
        
        for cell in hdr_cells:
            set_cell_background(cell, "E0E0E0")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.bold = True

        row_cells = table.add_row().cells
        row_cells[0].text = act.get("gv_hs", "")
        row_cells[1].text = act.get("san_pham", "")

    # Xuất ra bộ nhớ đệm (BytesIO)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 7. NÚT KÍCH HOẠT SOẠN GIÁO ÁN TỰ ĐỘNG (AI)
# ==========================================
st.markdown("---")
if st.button("🚀 BẮT ĐẦU SOẠN GIÁO ÁN CHUẨN 5512 (GEMINI AI)", use_container_width=True, type="primary"):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            
            prompt_lesson = f"""
            Bạn là một chuyên gia giáo dục. Hãy soạn kế hoạch bài dạy (Giáo án) theo chuẩn Công văn 5512 Bộ GD&ĐT Việt Nam cho bài học sau:
            - Môn học: {subject} - {grade}
            - Chương: {chapter_title}
            - Tên bài: {lesson_title} ({duration} tiết)
            - Yêu cầu cần đạt: {requirements}
            - Tích hợp: {', '.join(integrations)}

            Trả về nội dung dạng JSON cấu trúc sau (không chứa markdown thừa):
            {{
                "muc_tieu": "...",
                "thiet_bi": "...",
                "hoat_dong": [
                    {{
                        "ten_hoat_dong": "Hoạt động 1: Mở đầu / Khởi động",
                        "gv_hs": "...",
                        "san_pham": "..."
                    }},
                    {{
                        "ten_hoat_dong": "Hoạt động 2: Hình thành kiến thức mới",
                        "gv_hs": "...",
                        "san_pham": "..."
                    }},
                    {{
                        "ten_hoat_dong": "Hoạt động 3: Luyện tập",
                        "gv_hs": "...",
                        "san_pham": "..."
                    }},
                    {{
                        "ten_hoat_dong": "Hoạt động 4: Vận dụng",
                        "gv_hs": "...",
                        "san_pham": "..."
                    }}
                ]
            }}
            """

            with st.spinner("🤖 Gemini AI đang phân tích tài liệu và soạn giáo án 5512..."):
                response = model.generate_content(prompt_lesson)
                raw_text = response.text.strip()
                json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                clean_json = json_match.group(0) if json_match else raw_text
                plan_data = json.loads(clean_json)
                plan_data['lesson_title'] = lesson_title
                plan_data['duration'] = duration

                # Tạo file Word Docx
                docx_file = create_5512_docx(plan_data)
                
                st.success("🎉 Đã hoàn thành soạn Giáo án chuẩn 5512!")
                
                # Nút tải file Word về máy
                st.download_button(
                    label="📥 TẢI VỀ FILE GIÁO ÁN WORD (.DOCX)",
                    data=docx_file,
                    file_name=f"Giao_An_5512_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra trong quá trình soạn giáo án: {e}")
