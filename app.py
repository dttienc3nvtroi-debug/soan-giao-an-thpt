import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# Cấu hình trang Streamlit
st.set_page_config(page_title="Hệ thống Soạn Giáo án Tự Động 5512", layout="wide", page_icon="📝")

st.title("📝 HỆ THỐNG SOẠN BÀI DẠY CHUẨN CÔNG VĂN 5512")
st.caption("Cập nhật theo Danh mục NXB Giáo dục Việt Nam (taphuan.nxbgd.vn) | Tích hợp AI, STEM & Năng lực số")

# Thanh bên cấu hình
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
    
    st.markdown("---")
    school_name = st.text_input("Trường THPT:", "THPT Nguyễn Văn Trỗi")
    dept_name = st.text_input("Tổ chuyên môn:", "Tổ Toán")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")
    
    st.markdown("---")
    st.info("💡 **Mẹo:** Dữ liệu được đồng bộ theo cấu trúc Chương trình GDPT mới nhất từ Cổng thông tin NXB Giáo dục Việt Nam.")

# BƯỚC 1: CHỌN MÔN VÀ LỚP
st.subheader("📚 Bước 1: Chọn Môn học & Khối lớp")
col_sub, col_grd, col_book = st.columns(3)

with col_sub:
    subject = st.selectbox("Môn học/Hoạt động GD:", [
        "Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", 
        "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"
    ])

with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])

with col_book:
    book_series = st.selectbox("Bộ sách giáo khoa:", [
        "Kết nối tri thức với cuộc sống", 
        "Cánh diều", 
        "Chân trời sáng tạo"
    ])

# BƯỚC 2: NHẬP/CHỌN BÀI HỌC
st.subheader("📖 Bước 2: Thông tin bài dạy")

col_input1, col_input2 = st.columns(2)

with col_input1:
    chapter_title = st.text_input("Tên Chương / Chủ đề:", placeholder="Ví dụ: CHƯƠNG I: MỆNH ĐỀ VÀ TẬP HỢP")
    lesson_title = st.text_input("TÊN BÀI DẠY:", placeholder="Ví dụ: BÀI 1: MỆNH ĐỀ")
    duration = st.number_input("Thời gian thực hiện (Số tiết):", min_value=1, max_value=20, value=2)

with col_input2:
    requirements = st.text_area(
        "Yêu cầu cần đạt / Nội dung chính (Gõ vắn tắt, AI sẽ tự phát triển đầy đủ):", 
        placeholder="Ví dụ: Phát biểu mệnh đề, xác định tính đúng sai, ký hiệu ∀, ∃...", 
        height=110
    )

# BƯỚC 3: TÍCH HỢP PHƯƠNG PHÁP MỚI
st.subheader("🚀 Bước 3: Tích hợp Năng lực & Phương pháp hiện đại")
integrations = st.multiselect(
    "Lựa chọn các yếu tố tích hợp vào Kế hoạch bài dạy:",
    [
        "Năng lực Số / Ứng dụng CNTT (Phần mềm, Padlet, Kahoot...)", 
        "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)", 
        "Giáo dục STEM / STEAM", 
        "Phát triển Tư duy phản biện & Giải quyết vấn đề", 
        "Học tập qua Dự án (PBL) / Thảo luận nhóm"
    ],
    default=["Năng lực Số / Ứng dụng CNTT (Phần mềm, Padlet, Kahoot...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)

def generate_doc(content_text):
    doc = docx.Document()

    # Lề trang A4 chuẩn 5512
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    # Style Times New Roman 13pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)

    # Header Trường / Tổ / GV
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

    # Xóa viền bảng Header
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # Tiêu đề Bài dạy
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
    r_sub = p_sub.add_run(f"Môn học/Hoạt động giáo dục: {subject}; Lớp: {grade} ({book_series})\nThời gian thực hiện: ({duration} tiết)")
    r_sub.italic = True

    # Xuất các mục nội dung
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
        st.error("⚠️ Vui lòng nhập Tên bài dạy!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash-latest')

            integration_str = ", ".join(integrations) if integrations else "Không"

            prompt = f"""
            Bạn là một Chuyên gia Giáo dục THPT hàng đầu tại Việt Nam, am hiểu sâu sắc Chương trình GDPT 2018 và tài liệu từ Cổng thông tin Tập huấn Nhà xuất bản Giáo dục Việt Nam (taphuan.nxbgd.vn).

            Hãy soạn một Kế hoạch bài dạy (KHBD) hoàn chỉnh, chi tiết và chuẩn 100% theo CÔNG VĂN 5512/BGDĐT cho bài học sau:
            - Môn học: {subject} ({grade})
            - Bộ sách giáo khoa: {book_series}
            - Chương/Chủ đề: {chapter_title}
            - Tên bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - Yêu cầu cần đạt cơ bản: {requirements}
            - CÁC YẾU TỐ TÍCH HỢP BẮT BUỘC: {integration_str}

            YÊU CẦU TRÌNH BÀY CHUẨN KHUNG CÔNG VĂN 5512:

            I. Mục tiêu
            1. Về kiến thức: (Bám sát chuẩn YCĐ của Bộ GD&ĐT cho bài học này trong bộ sách {book_series}).
            2. Về năng lực:
               - Năng lực chung: Tự chủ và tự học, giao tiếp và hợp tác, giải quyết vấn đề và sáng tạo.
               - Năng lực đặc thù môn học: (Nêu rõ các biểu hiện năng lực môn {subject}).
               - Năng lực tích hợp ({integration_str}): Mô tả cụ thể hoạt động học sinh thực hiện (sử dụng công cụ số, AI, STEM...).
            3. Về phẩm chất: (Yêu nước, nhân ái, chăm chỉ, trung thực, trách nhiệm).

            II. Thiết bị dạy học và học liệu
            - Giáo viên: Kế hoạch bài dạy, bài giảng điện tử, thiết bị công nghệ, phần mềm hỗ trợ, phiếu học tập...
            - Học sinh: SGK {subject} {grade} ({book_series}), vở ghi, dụng cụ học tập...

            III. Tiến trình dạy học
            1. Hoạt động 1: Xác định vấn đề/nhiệm vụ học tập/Mở đầu
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện:
            - Bước 1: Chuyển giao nhiệm vụ: ...
            - Bước 2: Thực hiện nhiệm vụ: ...
            - Bước 3: Báo cáo, thảo luận: ...
            - Bước 4: Kết luận, nhận định: ...

            2. Hoạt động 2: Hình thành kiến thức mới/giải quyết vấn đề/thực thi nhiệm vụ đặt ra từ Hoạt động 1
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện: (Trình bày đầy đủ 4 bước: Chuyển giao, Thực hiện, Báo cáo, Kết luận. Lồng ghép các hoạt động tích hợp đã chọn).

            3. Hoạt động 3: Luyện tập
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện: (Trình bày đầy đủ 4 bước).

            4. Hoạt động 4: Vận dụng và tìm tòi mở rộng
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện: (Trình bày đầy đủ 4 bước, giao nhiệm vụ ứng dụng thực tế hoặc tra cứu công nghệ).
            """

            with st.spinner("✨ AI đang tra cứu chuẩn NXB Giáo dục và soạn thảo bài dạy 5512, thầy vui lòng đợi giây lát..."):
                response = model.generate_content(prompt)
                
                st.success("🎉 Đã tạo thành công Kế hoạch bài dạy!")
                
                doc_file = generate_doc(response.text)
                
                st.download_button(
                    label="📥 TẢI FILE WORD GIÁO ÁN (.DOCX)",
                    data=doc_file,
                    file_name=f"KHBD_5512_{lesson_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                st.markdown("---")
                st.subheader("👁️ Xem trước nội dung:")
                st.write(response.text)

        except Exception as e:
            st.error(f"Đã có lỗi xảy ra: {e}")
