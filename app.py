import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import time

st.set_page_config(
    page_title="Hệ thống Soạn Giáo án 5512 (KNTT Chuẩn)", 
    layout="wide", 
    page_icon="📝"
)

st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 1250px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, sans-serif; }
    .sidebar-title { color: #0f172a; font-size: 19px !important; font-weight: 700; margin-bottom: 10px; border-bottom: 3px solid #2563eb; padding-bottom: 4px; }
    .step-header { color: #dc2626 !important; font-size: 21px !important; font-weight: 700 !important; margin-top: 15px !important; margin-bottom: 8px !important; padding-left: 8px; border-left: 5px solid #dc2626; }
    div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #1e293b !important; }
    </style>
""", unsafe_allow_html=True)

# DATABASE KNTT
DATABASE_KNTT = {
    "Toán học": {
        "Lớp 10": {
            "Chương I. Mệnh đề và Tập hợp": ["Bài 1. Mệnh đề", "Bài 2. Tập hợp và các phép toán trên tập hợp"],
            "Chương II. Bất phương trình và Hệ bất phương trình bậc nhất hai ẩn": ["Bài 3. Bất phương trình bậc nhất hai ẩn", "Bài 4. Hệ bất phương trình bậc nhất hai ẩn"],
            "Chương III. Hệ thức lượng trong tam giác": ["Bài 5. Giá trị lượng giác của một góc từ 0° đến 180°", "Bài 6. Hệ thức lượng trong tam giác"],
            "Chương IV. Vectơ": ["Bài 7. Các khái niệm mở đầu", "Bài 8. Tổng và hiệu của hai vectơ", "Bài 9. Tích của một số với một vectơ", "Bài 10. Tích vô hướng của hai vectơ"],
            "Chương V. Các số đặc trưng của mẫu số liệu không ghép nhóm": ["Bài 11. Số gần đúng và sai số", "Bài 12. Số trung bình và trung vị của mẫu số liệu không ghép nhóm", "Bài 13. Tứ phân vị và mốt của mẫu số liệu không ghép nhóm", "Bài 14. Các số đặc trưng đo độ phân tán"],
            "Chương VI. Hàm số, đồ thị và ứng dụng": ["Bài 15. Hàm số và đồ thị", "Bài 16. Hàm số bậc hai", "Bài 17. Dấu của tam thức bậc hai", "Bài 18. Phương trình quy về phương trình bậc hai"],
            "Chương VII. Phương pháp tọa độ trong mặt phẳng": ["Bài 19. Phương trình đường thẳng", "Bài 20. Vị trí tương đối giữa hai đường thẳng. Góc và khoảng cách", "Bài 21. Đường tròn trong mặt phẳng tọa độ", "Bài 22. Ba đường conic trong mặt phẳng tọa độ"],
            "Chương VIII. Đại số tổ hợp": ["Bài 23. Quy tắc đếm", "Bài 24. Hoán vị, chỉnh hợp và tổ hợp", "Bài 25. Nhị thức Newton"],
            "Chương IX. Tính xác suất theo định nghĩa cổ điển": ["Bài 26. Biến cố và định nghĩa cổ điển của xác suất", "Bài 27. Thực hành tính xác suất theo định nghĩa cổ điển"]
        },
        "Lớp 11": {
            "Chương I. Hàm số lượng giác và Phương trình lượng giác": ["Bài 1. Giá trị lượng giác của góc lượng giác", "Bài 2. Công thức lượng giác", "Bài 3. Hàm số lượng giác", "Bài 4. Phương trình lượng giác cơ bản"],
            "Chương II. Dãy số. Cấp số cộng và Cấp số nhân": ["Bài 5. Dãy số", "Bài 6. Cấp số cộng", "Bài 7. Cấp số nhân"],
            "Chương III. Các số đặc trưng đo xu thế trung tâm của mẫu số liệu ghép nhóm": ["Bài 8. Mẫu số liệu ghép nhóm", "Bài 9. Các số đặc trưng đo xu thế trung tâm"],
            "Chương IV. Quan hệ song song trong không gian": ["Bài 10. Đường thẳng và mặt phẳng trong không gian", "Bài 11. Hai đường thẳng song song", "Bài 12. Đường thẳng song song với mặt phẳng", "Bài 13. Hai mặt phẳng song song", "Bài 14. Phép chiếu song song"],
            "Chương V. Giới hạn. Hàm số liên tục": ["Bài 15. Giới hạn của dãy số", "Bài 16. Giới hạn của hàm số", "Bài 17. Hàm số liên tục"],
            "Chương VI. Hàm số mũ và Hàm số lôgarit": ["Bài 18. Lũy thừa với số mũ thực", "Bài 19. Lôgarit", "Bài 20. Hàm số mũ và hàm số lôgarit", "Bài 21. Phương trình, bất phương trình mũ và lôgarit"],
            "Chương VII. Quan hệ vuông góc trong không gian": ["Bài 22. Hai đường thẳng vuông góc", "Bài 23. Đường thẳng vuông góc với mặt phẳng", "Bài 24. Hai mặt phẳng vuông góc", "Bài 25. Khoảng cách trong không gian", "Bài 26. Góc giữa đường thẳng và mặt phẳng. Góc giữa hai mặt phẳng"],
            "Chương VIII. Các quy tắc tính xác suất": ["Bài 27. Biến cố xung khắc và quy tắc cộng xác suất", "Bài 28. Biến cố độc lập và quy tắc nhân xác suất"],
            "Chương IX. Đạo hàm": ["Bài 29. Định nghĩa và ý nghĩa của đạo hàm", "Bài 30. Các quy tắc tính đạo hàm"]
        },
        "Lớp 12": {
            "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số": ["Bài 1. Tính đơn điệu và cực trị của hàm số", "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số", "Bài 3. Đường tiệm cận của đồ thị hàm số", "Bài 4. Khảo sát sự biến thiên và vẽ đồ thị của hàm số", "Bài 5. Ứng dụng đạo hàm để giải quyết một số vấn đề thực tiễn"],
            "Chương II. Tọa độ của vectơ trong không gian": ["Bài 6. Vectơ trong không gian", "Bài 7. Hệ trục tọa độ trong không gian", "Bài 8. Biểu thức tọa độ của các phép toán vectơ"],
            "Chương III. Các số đặc trưng đo độ phân tán của mẫu số liệu ghép nhóm": ["Bài 9. Khoảng biến thiên và khoảng tứ phân vị của mẫu số liệu ghép nhóm", "Bài 10. Phương sai và độ lệch chuẩn của mẫu số liệu ghép nhóm"],
            "Chương IV. Nguyên hàm và Tích phân": ["Bài 11. Nguyên hàm", "Bài 12. Tích phân", "Bài 13. Ứng dụng hình học của tích phân"],
            "Chương V. Phương pháp tọa độ trong không gian": ["Bài 14. Phương trình mặt phẳng", "Bài 15. Phương trình đường thẳng trong không gian", "Bài 16. Công thức tính góc và khoảng cách trong không gian", "Bài 17. Phương trình mặt cầu"],
            "Chương VI. Xác suất có điều kiện": ["Bài 18. Xác suất có điều kiện", "Bài 19. Công thức xác suất toàn phần và công thức Bayes"]
        }
    }
}

# HÀM GỌI GEMINI VỚI GEMINI-1.5-FLASH
def generate_content_safe(api_key, contents):
    genai.configure(api_key=api_key)
    
    # Sử dụng gemini-1.5-flash để tránh nghẽn Quota của dòng 2.0
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        generation_config=genai.GenerationConfig(
            temperature=0.3,
            max_output_tokens=2048
        )
    )
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            return model.generate_content(contents)
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "Quota" in err_msg:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 15
                    st.warning(f"⏳ Tạm thời đạt trần Token. Tự động nghỉ {wait_time}s... (Lần {attempt + 1}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    raise Exception("❌ Đã hết Quota miễn phí trong phút này! Vui lòng tạo API Key ở Gmail KHÁC hoặc chờ 1-2 phút rồi bấm lại.")
            else:
                raise e

# SIDEBAR
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 CẤU HÌNH HỆ THỐNG</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key...")
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN TRƯỜNG & GV</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

st.markdown("""
    <div style="text-align: center; margin-bottom: 15px; background: #eff6ff; padding: 15px; border-radius: 10px; border: 1px solid #bfdbfe;">
        <div style="font-size: 26px; font-weight: 800; color: #1e3a8a;">HỆ THỐNG SOẠN GIÁO ÁN 5512 TỰ ĐỘNG CHUẨN KNTT</div>
        <div style="font-size: 16px; font-weight: 700; color: #047857; margin-top: 4px;">📘 BỘ SÁCH: KẾT NỐI TRI THỨC VỚI CUỘC SỐNG</div>
        <div style="font-size: 18px; font-weight: 600; color: #2563eb; margin-top: 4px;">Tác giả: DƯƠNG TẤN TIẾN — THPT NGUYỄN VĂN TRỖI</div>
    </div>
""", unsafe_allow_html=True)

# BƯỚC 1: CHỌN MÔN, LỚP
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN, LỚP, CHƯƠNG & BÀI HỌC</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    subject_list = list(DATABASE_KNTT.keys()) + ["Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"]
    subject = st.selectbox("Môn học:", subject_list)

with c2:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

with c3:
    duration = st.number_input("Số tiết dạy:", value=2, min_value=1, max_value=10)

available_chapters = DATABASE_KNTT.get(subject, {}).get(grade, {})
col_ch, col_les = st.columns([1, 1])

if available_chapters:
    with col_ch:
        selected_chapter = st.selectbox("Chọn Chương / Chủ đề:", list(available_chapters.keys()))
    with col_les:
        lessons_in_chap = available_chapters[selected_chapter]
        selected_lesson = st.selectbox("Chọn Bài dạy cụ thể:", lessons_in_chap)
else:
    with col_ch:
        selected_chapter = st.text_input("Tên Chương / Chủ đề:", placeholder="Nhập tên chương...")
    with col_les:
        selected_lesson = st.text_input("Tên Bài dạy cụ thể:", placeholder="Nhập tên bài...")

# BƯỚC 2: YCĐ
st.markdown('<div class="step-header">🎯 BƯỚC 2: YÊU CẦU CẦN ĐẠT (SGVKẾT NỐI TRI THỨC)</div>', unsafe_allow_html=True)

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
                prompt_ycd = f"Liệt kê ngắn gọn Yêu cầu cần đạt SGV Kết nối tri thức - {grade}, môn {subject}, bài '{selected_lesson}'."
                with st.spinner("⚡ AI đang tải YCĐ..."):
                    res = generate_content_safe(api_key.strip(), [prompt_ycd])
                    st.session_state['auto_ycd'] = res.text
                    st.success("✅ Đã tải xong YCĐ!")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

with col_upload:
    uploaded_sgv_file = st.file_uploader("Hoặc Tải file ảnh/PDF trang SGV (Nếu có):", type=["pdf", "png", "jpg", "jpeg"])

ycd_content = st.text_area("Mô tả Yêu cầu cần đạt:", value=st.session_state.get('auto_ycd', ''), height=100, placeholder="Bấm 'Tải nhanh Yêu cầu cần đạt' để tự điền...")

# BƯỚC 3: TÍCH HỢP
st.markdown('<div class="step-header">🚀 BƯỚC 3: YẾU TỐ TÍCH HỢP</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn tích hợp:",
    ["Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI (Gemini, ChatGPT...)", "Giáo dục STEM/STEAM", "Phát triển Tư duy phản biện"],
    default=["Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI (Gemini, ChatGPT...)"]
)

# XUẤT DOCX
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

# TẠO GIÁO ÁN
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 BẤM TẠO GIÁO ÁN WORD CHUẨN 5512", type="primary", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở bên trái!")
    elif not selected_lesson:
        st.error("⚠️ Vui lòng chọn Tên Bài dạy!")
    else:
        try:
            integration_str = ", ".join(integrations) if integrations else "Không"
            
            # XỬ LÝ PHẦN 1
            st.info("🔄 Đang xử lý Phần 1: Mục tiêu, Thiết bị & Khởi động...")
            prompt_part1 = f"""
            Soạn Phần 1 KHBD 5512 bài: {selected_lesson} ({grade}, {subject}, Kết nối tri thức).
            YCĐ: {ycd_content}. Tích hợp: {integration_str}.
            Gồm:
            I. MỤC TIÊU (Kiến thức, Năng lực, Phẩm chất)
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC:
            Hoạt động 1: Mở đầu (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện 4 bước).
            """
            
            contents1 = [prompt_part1]
            if uploaded_sgv_file is not None:
                bytes_data = uploaded_sgv_file.getvalue()
                contents1.append({"mime_type": uploaded_sgv_file.type, "data": bytes_data})
                
            res1 = generate_content_safe(api_key.strip(), contents1)
            
            # Tạm nghỉ 5 giây để Google reset Quota
            time.sleep(5)
            
            # XỬ LÝ PHẦN 2
            st.info("🔄 Đang xử lý Phần 2: Hình thành kiến thức, Luyện tập & Vận dụng...")
            prompt_part2 = f"""
            Soạn tiếp Phần 2 KHBD 5512 bài: {selected_lesson} ({grade}, {subject}, Kết nối tri thức).
            Gồm các Hoạt động còn lại (mỗi hoạt động đủ 4 mục a, b, c, d):
            Hoạt động 2: Hình thành kiến thức mới
            Hoạt động 3: Luyện tập (có bài tập + lời giải)
            Hoạt động 4: Vận dụng
            """
            res2 = generate_content_safe(api_key.strip(), [prompt_part2])
            
            # TỔNG HỢP NỘI DUNG
            full_text = res1.text + "\n\n" + res2.text
            doc_file = generate_doc(full_text)
            
            st.success("🎉 Đã tạo xong toàn bộ Giáo án 5512!")
            st.download_button(
                label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                data=doc_file,
                file_name=f"KHBD_5512_{selected_lesson.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown(full_text)

        except Exception as e:
            st.error(f"{str(e)}")
