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
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Sách KNTT)", 
    layout="wide", 
    page_icon="📝"
)

# Giao diện & Font chữ
st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 1250px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    .sidebar-title { color: #0f172a; font-size: 19px !important; font-weight: 700; margin-bottom: 10px; border-bottom: 3px solid #2563eb; padding-bottom: 4px; }
    .step-header { color: #dc2626 !important; font-size: 21px !important; font-weight: 700 !important; margin-top: 15px !important; margin-bottom: 8px !important; padding-left: 8px; border-left: 5px solid #dc2626; }
    div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #1e293b !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CƠ SỞ DỮ LIỆU CÁC MÔN - SÁCH KẾT NỐI TRI THỨC
# ==========================================
DATABASE_KNTT = {
    "Toán học": {
        "Lớp 10": {
            "Chương I. Mệnh đề và Tập hợp": [
                "Bài 1. Mệnh đề",
                "Bài 2. Tập hợp và các phép toán trên tập hợp"
            ],
            "Chương II. Bất phương trình và Hệ bất phương trình bậc nhất hai ẩn": [
                "Bài 3. Bất phương trình bậc nhất hai ẩn",
                "Bài 4. Hệ bất phương trình bậc nhất hai ẩn"
            ],
            "Chương III. Hàm số bậc hai và Đồ thị": [
                "Bài 5. Giá trị lượng giác của một góc từ 0 đến 180 độ",
                "Bài 6. Hệ thức lượng trong tam giác",
                "Bài 7. Các khái niệm mở đầu về Vectơ"
            ]
        },
        "Lớp 11": {
            "Chương I. Hàm số lượng giác và Phương trình lượng giác": [
                "Bài 1. Giá trị lượng giác của góc lượng giác",
                "Bài 2. Công thức lượng giác",
                "Bài 3. Hàm số lượng giác",
                "Bài 4. Phương trình lượng giác cơ bản"
            ],
            "Chương II. Dãy số. Cấp số cộng và Cấp số nhân": [
                "Bài 5. Dãy số",
                "Bài 6. Cấp số cộng",
                "Bài 7. Cấp số nhân"
            ]
        },
        "Lớp 12": {
            "Chương I. Ứng dụng Đạo hàm để Khảo sát và Vẽ đồ thị Hàm số": [
                "Bài 1. Tính đơn điệu và cực trị của hàm số",
                "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số",
                "Bài 3. Đường tiệm cận của đồ thị hàm số",
                "Bài 4. Khảo sát sự biến thiên và vẽ đồ thị của hàm số"
            ],
            "Chương II. Tọa độ của Vectơ trong Không gian": [
                "Bài 6. Vectơ trong không gian",
                "Bài 7. Tọa độ của vectơ trong không gian"
            ]
        }
    },
    "Vật lý": {
        "Lớp 10": {
            "Chương I. Mở đầu": [
                "Bài 1. Tắc hại của việc không tuân thủ tắc an toàn trong phòng thực hành",
                "Bài 2. Vấn đề an toàn trong Vật lý"
            ],
            "Chương II. Động học": [
                "Bài 4. Độ dịch chuyển và quãng đường đi được",
                "Bài 5. Tốc độ và vận tốc",
                "Bài 7. Gia tốc"
            ]
        },
        "Lớp 11": {
            "Chương I. Dao động": [
                "Bài 1. Dao động điều hòa",
                "Bài 2. Mô tả dao động điều hòa",
                "Bài 3. Vận tốc, gia tốc trong dao động điều hòa"
            ]
        },
        "Lớp 12": {
            "Chương I. Vật lý Nhiệt": [
                "Bài 1. Cấu trúc của chất. Sự chuyển thể",
                "Bài 2. Thang nhiệt độ",
                "Bài 3. Nhiệt lượng và nhiệt dung riêng"
            ]
        }
    }
}

# ==========================================
# HÀM KẾT NỐI GEMINI SPEED
# ==========================================
def get_working_model(api_key):
    genai.configure(api_key=api_key)
    try:
        models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
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
    st.markdown('<div class="sidebar-title">🔑 CẤU HÌNH HỆ THỐNG</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key...")
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN TRƯỜNG & GV</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ UNG DUNG
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 15px; background: #eff6ff; padding: 15px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 26px; font-weight: 800; color: #1e3a8a;">HỆ THỐNG SOẠN GIÁO ÁN 5512 TỰ ĐỘNG</div>
        <div style="font-size: 16px; font-weight: 700; color: #047857; margin-top: 4px;">📘 BỘ SÁCH: KẾT NỐI TRI THỨC VỚI CUỘC SỐNG</div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 4px;">Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: BẤM CHỌN CHƯƠNG & BÀI HỌC
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN, LỚP, CHƯƠNG & BÀI HỌC</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    subject_list = list(DATABASE_KNTT.keys()) + ["Ngữ văn", "Tiếng Anh", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"]
    subject = st.selectbox("Môn học:", subject_list)

with c2:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

with c3:
    duration = st.number_input("Số tiết dạy:", value=2, min_value=1, max_value=10)

# Lấy danh sách Chương & Bài theo Môn + Lớp
available_chapters = DATABASE_KNTT.get(subject, {}).get(grade, {})

col_ch, col_les = st.columns([1, 1])

if available_chapters:
    with col_ch:
        selected_chapter = st.selectbox("Chọn Chương / Chủ đề:", list(available_chapters.keys()))
    with col_les:
        lessons_in_chap = available_chapters[selected_chapter]
        selected_lesson = st.selectbox("Chọn Tên Bài dạy cụ thể:", lessons_in_chap)
else:
    with col_ch:
        selected_chapter = st.text_input("Tên Chương / Chủ đề:", placeholder="Nhập tên chương...")
    with col_les:
        selected_lesson = st.text_input("Tên Bài dạy cụ thể:", placeholder="Nhập tên bài...")

# ==========================================
# BƯỚC 2: YÊU CẦU CẦN ĐẠT (AI TỰ TẠO RA KHI BẤM CHỌN BÀI)
# ==========================================
st.markdown('<div class="step-header">🎯 BƯỚC 2: YÊU CẦU CẦN ĐẠT (SGV KẾT NỐI TRI THỨC)</div>', unsafe_allow_html=True)

# Tự động load YCĐ khi bấm đổi bài
if 'current_lesson_key' not in st.session_state or st.session_state['current_lesson_key'] != selected_lesson:
    st.session_state['current_lesson_key'] = selected_lesson
    st.session_state['auto_ycd'] = ""

col_load_ycd, col_upload = st.columns([1, 1])
with col_load_ycd:
    if st.button("⚡ Tải nhanh Yêu cầu cần đạt chuẩn SGV bài này", type="secondary", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng dán Gemini API Key ở menu bên trái!")
        else:
            try:
                model_name = get_working_model(api_key.strip())
                model = genai.GenerativeModel(model_name)
                
                prompt_ycd = f"""Bạn là Chuyên gia Giáo dục. Hãy trích xuất NGUYÊN VĂN các Yêu cầu cần đạt (về kiến thức, năng lực) chuẩn SGV bộ sách Kết nối tri thức với cuộc sống - môn {subject} ({grade}) cho bài: "{selected_lesson}" (Thuộc {selected_chapter}). Viết ngắn gọn dạng gạch đầu dòng."""
                
                with st.spinner("⚡ AI đang tải YCĐ từ SGV Kết nối tri thức..."):
                    res = model.generate_content(prompt_ycd)
                    st.session_state['auto_ycd'] = res.text
                    st.success("✅ Đã tải xong YCĐ!")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

with col_upload:
    uploaded_sgv_file = st.file_uploader("Hoặc Tải file ảnh/PDF trang SGV (Nếu có):", type=["pdf", "png", "jpg", "jpeg"])

ycd_content = st.text_area("Mô tả Yêu cầu cần đạt (Có thể sửa thêm):", value=st.session_state.get('auto_ycd', ''), height=140, placeholder="Nhấn nút 'Tải nhanh Yêu cầu cần đạt' ở trên để AI tự điền...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: YẾU TỐ TÍCH HỢP</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn tích hợp:",
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
    r_title = p_title.add_run(f"KẾ HOẠCH BÀI DẠY: {selected_lesson.upper()}")
    r_title.bold = True
    r_title.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.add_run(f"Môn: {subject} ({grade}) - Sách: Kết nối tri thức với cuộc sống\nThời lượng: {duration} tiết").italic = True

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
if st.button("🚀 BẤM TẠO GIÁO ÁN WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở bên trái!")
    elif not selected_lesson:
        st.error("⚠️ Vui lòng chọn Tên Bài dạy!")
    else:
        try:
            model_name = get_working_model(api_key.strip())
            model = genai.GenerativeModel(model_name)
            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt_5512 = f"""
            Đóng vai Giáo viên Giỏi môn {subject}. Soạn Kế hoạch bài dạy CHUẨN CÔNG VĂN 5512 đầy đủ và chi tiết.
            
            THÔNG TIN BÀI DẠY:
            - Môn: {subject} ({grade}) - Sách: Kết nối tri thức với cuộc sống
            - Bài dạy: {selected_lesson} (Thuộc {selected_chapter})
            - Thời lượng: {duration} tiết
            - Yếu tố tích hợp: {integration_str}
            - Yêu cầu cần đạt SGV: {ycd_content}

            CẤU TRÚC BẮT BUỘC 5512:
            I. MỤC TIÊU (1. Kiến thức; 2. Năng lực; 3. Phẩm chất)
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Chi tiết cả {duration} tiết)
            Các Hoạt động (Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng).
            Mỗi hoạt động phải có 4 mục:
            a) Mục tiêu
            b) Nội dung (Nêu chi tiết bài tập/câu hỏi trong SGK Kết nối tri thức)
            c) Sản phẩm (Đáp án, lời giải chi tiết)
            d) Tổ chức thực hiện (Đủ 4 bước: Bước 1 Chuyển giao; Bước 2 Thực hiện; Bước 3 Báo cáo; Bước 4 Kết luận)
            """

            contents = [prompt_5512]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                contents.append({"mime_type": uploaded_sgv_file.type, "data": bytes_data})

            with st.spinner("⚡ Đang tạo Giáo án Word 5512 chi tiết..."):
                res = model.generate_content(contents)
                doc_file = generate_doc(res.text)
                
                st.success("🎉 Đã tạo xong Giáo án!")
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{selected_lesson.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown(res.text)

        except Exception as e:
            st.error(f"❌ Lỗi: {str(e)}")
