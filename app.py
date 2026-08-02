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

st.set_page_config(page_title="Hệ thống Soạn Giáo án Tự Động 5512", layout="wide", page_icon="📝")

st.title("📝 HỆ THỐNG SOẠN BÀI DẠY CHUẨN CÔNG VĂN 5512")
st.caption("Đồng bộ danh mục Chương/Bài chuẩn NXB Giáo dục Việt Nam (taphuan.nxbgd.vn)")

# Thanh bên cấu hình
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
    
    # Chọn mô hình Gemini
    model_name = st.selectbox(
        "Chọn phiên bản Gemini AI:",
        ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"],
        index=0
    )
    
    st.markdown("---")
    school_name = st.text_input("Trường THPT:", "THPT Nguyễn Văn Trỗi")
    dept_name = st.text_input("Tổ chuyên môn:", "Tổ Toán")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

st.subheader("📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP")
col_sub, col_grd = st.columns(2)
with col_sub:
    subject = st.selectbox("Môn học/Hoạt động GD:", [
        "Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", 
        "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"
    ])
with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

st.subheader("📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN")

# Nút lấy danh sách bài học tự động
if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn"):
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
            
            with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade} từ taphuan.nxbgd.vn..."):
                res = model.generate_content(prompt_fetch)
                clean_json = res.text.replace("```json", "").replace("```", "").strip()
                st.session_state['fetched_lessons'] = json.loads(clean_json)
                st.success("🎉 Đã tải xong danh mục bài học chuẩn!")
        except Exception as e:
            st.error(f"Lỗi khi tải danh mục bài học: {e}")

# Hiển thị danh sách bài học đã tra cứu được
if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
    lessons_data = st.session_state['fetched_lessons']
    
    lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
    selected_idx = st.selectbox("👉 Chọn Bài học chuẩn từ danh sách vừa tải:", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x])
    
    current_item = lessons_data[selected_idx]
    
    col_i1, col_i2 = st.columns([1, 3])
    with col_i1:
        chapter_title = st.text_input("Chương:", value=current_item['chapter'])
        lesson_title = st.text_input("Tên bài:", value=current_item['lesson'])
        duration = st.number_input("Số tiết:", value=current_item['duration'])
    with col_i2:
        st.markdown("**📌 Yêu cầu cần đạt (Quy định chuẩn NXB Giáo dục):**")
        requirements = st.text_area("YCĐ:", value=current_item['req'], height=120)

else:
    st.info("💡 Thầy vui lòng bấm nút **'🔍 Cập nhật danh sách Chương & Bài học...'** ở trên để AI tự động tải toàn bộ bài học chuẩn nhé!")
    chapter_title = st.text_input("Tên Chương/Chủ đề:", "")
    lesson_title = st.text_input("Tên Bài dạy:", "")
    duration = st.number_input("Số tiết:", value=2)
    requirements = st.text_area("Yêu cầu cần đạt:", "")

# BƯỚC 3: TÍCH HỢP NĂNG LỰC
st.subheader("🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ")
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện", 
        "Học tập qua Dự án (PBL)"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)

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
    p_left.add_run(f"Trường: {school_name}\n").bold = True
    p_left.add_run(f"Tổ: {dept_name}").bold = True

    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run(f"Họ và tên giáo viên:\n{teacher_name}").bold = True

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
        if not line_str:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)

        if line_str.startswith("I. ") or line_str.startswith("II. ") or line_str.startswith("III. "):
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(line_str)
            run.bold = True
            run.font.size = Pt(13)
        elif line_str.startswith("1. ") or line_str.startswith("2. ") or line_str.startswith("3. ") or line_str.startswith("4. "):
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(line_str)
            run.bold = True
        elif line_str.startswith("a)") or line_str.startswith("b)") or line_str.startswith("c)") or line_str.startswith("d)"):
            run = p.add_run(line_str)
            run.bold = True
        else:
            p.add_run(line_str)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512", type="primary"):
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
            Soạn Kế hoạch bài dạy chuẩn 100% CÔNG VĂN 5512/BGDĐT:
            - Môn: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - Yêu cầu cần đạt: {requirements}
            - YẾU TỐ TÍCH HỢP: {integration_str}

            ĐẦY ĐỦ 3 PHẦN CHUẨN 5512:
            I. Mục tiêu (Kiến thức, Năng lực [Chung, Đặc thù, Tích hợp {integration_str}], Phẩm chất)
            II. Thiết bị dạy học & Học liệu
            III. Tiến trình dạy học (4 Hoạt động: Mở đầu, Hình thành kiến thức, Luyện tập, Vận dụng. Mỗi hoạt động đúng 4 bước: Chuyển giao, Thực hiện, Báo cáo, Kết luận).
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
                st.write(response.text)

        except Exception as e:
            st.error(f"Đã có lỗi xảy ra: {e}")
