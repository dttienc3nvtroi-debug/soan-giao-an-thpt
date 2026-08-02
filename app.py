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
st.set_page_config(page_title="Hệ thống Soạn Giáo án THPT 5512", layout="wide", page_icon="📝")

st.title("📝 HỆ THỐNG TẠO KHBD CHUẨN CÔNG VĂN 5512")
st.caption("Ứng dụng hỗ trợ Giáo viên THPT tạo Kế hoạch bài dạy chuẩn định dạng Bộ GD&ĐT xuất file Word")

# Thanh bên tả cấu hình
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
    
    st.markdown("---")
    school_name = st.text_input("Trường THPT:", "THPT Nguyễn Văn Trỗi")
    dept_name = st.text_input("Tổ chuyên môn:", "Tổ Toán")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# Form nhập thông tin bài học
col1, col2 = st.columns(2)
with col1:
    subject = st.selectbox("Môn học/Hoạt động giáo dục:", ["Toán học", "Vật lý", "Hóa học", "Ngữ văn", "Tiếng Anh", "Lịch sử", "Địa lý", "Tin học", "Sinh học", "GDKT&PL"])
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
    lesson_title = st.text_input("TÊN BÀI DẠY:", "BÀI 1: MỆNH ĐỀ")

with col2:
    duration = st.number_input("Thời gian thực hiện (Số tiết):", min_value=1, max_value=20, value=4)
    chapter_title = st.text_input("Tên Chương/Chủ đề:", "CHƯƠNG I: MỆNH ĐỀ VÀ TẬP HỢP")
    requirements = st.text_area("Yêu cầu cần đạt / Nội dung chính:", "Thiết lập, phát biểu mệnh đề phủ định, kéo theo, tương đương, kí hiệu ∀,∃. Xác định tính đúng sai...", height=100)

def generate_doc(content_text):
    doc = docx.Document()

    # Cấu hình lề trang A4 chuẩn (Top 2cm, Bottom 2cm, Left 3cm, Right 2cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    # Style mặc định Times New Roman 13pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)

    # Bảng Header thông tin Trường/Tổ - Họ tên GV
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)

    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    r_l1 = p_left.add_run(f"Trường: {school_name}\n")
    r_l1.bold = True
    r_l2 = p_left.add_run(f"Tổ: {dept_name}")
    r_l2.bold = True

    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_r1 = p_right.add_run(f"Họ và tên giáo viên:\n{teacher_name}")
    r_r1.bold = True

    # Xóa viền bảng Header
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
            tcPr.append(tcBorders)

    # Tiêu đề bài dạy
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

    # Nội dung chính từ AI
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
        else:
            p.add_run(line_str)

    # Xuất thành Stream Bytes
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD chuẩn 5512", type="primary"):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google Gemini API Key ở thanh menu bên trái!")
    else:
        try:
            genai.configure(api_key=api_key)
            # Dùng model chuẩn nhất hiện nay
            model = genai.GenerativeModel('gemini-3.6-flash')

            prompt = f"""
            Hãy đóng vai là một Chuyên gia Giáo dục THPT tại Việt Nam. Hãy soạn Kế hoạch bài dạy (KHBD) chuẩn 100% theo định dạng CÔNG VĂN 5512/BGDĐT cho bài học sau:
            - Môn học: {subject} ({grade})
            - Chương/Chủ đề: {chapter_title}
            - Tên bài dạy: {lesson_title}
            - Thời lượng: {duration} tiết
            - Yêu cầu cần đạt: {requirements}

            TRÌNH BÀY CHÍNH XÁC THEO KHUNG CẤU TRÚC SAU (Không thêm bớt tiêu đề mục lớn):

            I. Mục tiêu
            1. Về kiến thức: (Nêu cụ thể nội dung kiến thức học sinh cần học theo YCĐ).
            2. Về năng lực: (Nêu cụ thể năng lực chung và năng lực đặc thù môn học).
            3. Về phẩm chất: (Nêu cụ thể các phẩm chất gắn với bài dạy).

            II. Thiết bị dạy học và học liệu
            - Giáo viên: (Màn hình chiếu, phiếu học tập, SGK,...)
            - Học sinh: (Vở ghi, SGK,...)

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
            d) Tổ chức thực hiện: (Ghi rõ 4 bước: Chuyển giao, Thực hiện, Báo cáo, Kết luận)

            3. Hoạt động 3: Luyện tập
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện: (Ghi rõ 4 bước)

            4. Hoạt động 4: Vận dụng và tìm tòi mở rộng
            a) Mục tiêu: ...
            b) Nội dung: ...
            c) Sản phẩm: ...
            d) Tổ chức thực hiện: (Ghi rõ 4 bước)
            """

            with st.spinner("✨ Hệ thống đang soạn thảo KHBD chuẩn 5512, thầy vui lòng đợi giây lát..."):
                response = model.generate_content(prompt)
                
                st.success("🎉 Đã tạo thành công Kế hoạch bài dạy!")
                
                # Tạo file docx
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
