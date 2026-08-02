import streamlit as st
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import google.generativeai as genai
import io

# 1. CẤU HÌNH GIAO DIỆN
st.set_page_config(page_title="Hệ thống Soạn Giáo án THPT - Kết nối tri thức", layout="wide")

st.title("📚 HỆ THỐNG TẠO KHBD NỘI BỘ - THPT")
st.caption("Dành riêng cho Giáo viên THPT - Áp dụng Bộ sách Kết nối tri thức (2026)")

# Sidebar nhập API Key và cấu hình chung
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
    
    st.divider()
    school_name = st.text_input("Trường THPT:", "THPT ...")
    group_name = st.text_input("Tổ chuyên môn:", "Tổ Toán - Tin")
    teacher_name = st.text_input("Họ và tên GV:", "Nguyễn Văn A")

# 2. KHU VỰC NHẬP THÔNG TIN BÀI HỌC
col1, col2 = st.columns(2)

with col1:
    subject = st.selectbox("Môn học (THPT):", [
        "Toán học", "Ngữ văn", "Tiếng Anh", "Vật lí", "Hóa học", 
        "Sinh học", "Lịch sử", "Địa lí", "GDKT & PL", "Tin học", 
        "Công nghệ", "GDQP & AN", "Hoạt động trải nghiệm, HN"
    ])
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
    chapter = st.text_input("Tên Chương/Chủ đề:", "CHƯƠNG I: MỆNH ĐỀ VÀ TẬP HỢP")
    lesson_name = st.text_input("Tên Bài dạy:", "BÀI 1: MỆNH ĐỀ")

with col2:
    total_periods = st.number_input("Số tiết dạy:", min_value=1, max_value=10, value=4)
    integrated_skills = st.multiselect(
        "Tích hợp các năng lực đặc thù:",
        ["Năng lực Số", "Tích hợp STEM", "Tích hợp AI/Công nghệ", "Giáo dục Địa phương", "Học sinh hòa nhập/khuyết tật"],
        default=["Năng lực Số"]
    )
    lesson_content_summary = st.text_area("Yêu cầu cần đạt / Nội dung chính của bài:", 
                                         "Thiết lập, phát biểu mệnh đề phủ định, kéo theo, tương đương, kí hiệu ∀,∃. Xác định tính đúng sai...")

# 3. HÀM TẠO PROMPT GỬI CHO AI
def build_prompt(subject, grade, chapter, lesson_name, total_periods, integrated_skills, lesson_content_summary):
    skills_str = ", ".join(integrated_skills)
    prompt = f"""
Bạn là một Chuyên gia Giáo dục THPT tại Việt Nam. Hãy soạn Kế hoạch bài dạy (KHBD) cho môn {subject} - {grade} theo chuẩn Công văn BGDĐT.
Thông tin bài học:
- Bộ sách: Kết nối tri thức với cuộc sống (Chương trình thống nhất 2026).
- Chương: {chapter}
- Tên bài: {lesson_name} ({total_periods} tiết)
- Yêu cầu tích hợp đặc thù: {skills_str}
- Tóm tắt yêu cầu cần đạt: {lesson_content_summary}

MẪU CẤU TRÚC BẮT BUỘC SỬ DỤNG (Tuân thủ chính xác từng tiêu mục):

I. MỤC TIÊU
1. Về kiến thức: (Nêu cụ thể yêu cầu cần đạt)
2. Về năng lực:
- Năng lực chung: Tự chủ và tự học, Giao tiếp và hợp tác, Giải quyết vấn đề và sáng tạo.
- Năng lực riêng (Môn học): (Ghi chi tiết)
- Năng lực số: (Nêu chi tiết khai thác dữ liệu, giao tiếp môi trường số, sáng tạo nội dung số nếu có chọn)
3. Về phẩm chất: Chăm chỉ, trung thực, trách nhiệm...

II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
1. Đối với GV: SGK, máy tính, bảng phụ...
2. Đối với HS: SGK, vở ghi, thiết bị công nghệ (nếu có)...

III. TIẾN TRÌNH DẠY HỌC
Hãy chia tiến trình dạy học chi tiết cho từng TIẾT (Ví dụ Tiết 1, Tiết 2...).
Mỗi tiết gồm các Hoạt động (Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng).
Mỗi Hoạt động ĐẦY ĐỦ 4 phần:
a) Mục tiêu
b) Nội dung
c) Sản phẩm
d) Tổ chức thực hiện (Bước 1: Chuyển giao, Bước 2: Thực hiện, Bước 3: Báo cáo thảo luận, Bước 4: Kết luận nhận định).

LƯU Ý RẤT QUAN TRỌNG: Không dùng định dạng Markdown hoa mỹ, viết văn phong sư phạm chuẩn mực, mạch lạc, chính xác kiến thức THPT.
"""
    return prompt

# 4. XỬ LÝ KHI BẤM NÚT TẠO KHBD
if st.button("🚀 BẮT ĐẦU TẠO KHBD WORD", type="primary"):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Google API Key ở thanh bên trái!")
    else:
        try:
            with st.spinner("AI đang soạn thảo Kế hoạch bài dạy đúng chuẩn... Vui lòng đợi trong giây lát!"):
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # Gọi AI sinh nội dung
                prompt = build_prompt(subject, grade, chapter, lesson_name, total_periods, integrated_skills, lesson_content_summary)
                response = model.generate_content(prompt)
                ai_text = response.text

                # Khởi tạo File Word (.docx)
                doc = docx.Document()
                
                # Cấu hình Margins (Lề chuẩn: Trên/Dưới/Phải 2cm, Trái 3cm)
                for section in doc.sections:
                    section.top_margin = Inches(0.79)
                    section.bottom_margin = Inches(0.79)
                    section.left_margin = Inches(1.18)
                    section.right_margin = Inches(0.79)

                # Cấu hình Header thông tin Trường/Tổ/GV (Bảng 2 cột ẩn viền)
                table_header = doc.add_table(rows=2, cols=2)
                table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                cell_top_left = table_header.cell(0, 0)
                cell_top_right = table_header.cell(0, 1)
                cell_bot_left = table_header.cell(1, 0)
                cell_bot_right = table_header.cell(1, 1)

                cell_top_left.paragraphs[0].add_run(f"Trường: {school_name}\nTổ: {group_name}")
                cell_top_right.paragraphs[0].add_run(f"Họ và tên giáo viên:\n{teacher_name}")
                cell_bot_left.paragraphs[0].add_run("Ngày soạn: ……………")
                cell_bot_right.paragraphs[0].add_run("Ngày dạy: ……………")

                doc.add_paragraph().paragraph_format.space_after = Pt(12)

                # Tiêu đề Chương và Bài
                p_chap = doc.add_paragraph()
                p_chap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_chap = p_chap.add_run(f"{chapter.upper()}")
                run_chap.bold = True
                run_chap.font.size = Pt(14)

                p_less = doc.add_paragraph()
                p_less.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_less = p_less.add_run(f"{lesson_name.upper()} ({total_periods} tiết)")
                run_less.bold = True
                run_less.font.size = Pt(14)

                doc.add_paragraph()

                # Đưa toàn bộ nội dung AI vào File Word
                for line in ai_text.split('\n'):
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(3)

                # Lưu vào bộ nhớ tạm để cho Tải về
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.success("🎉 Đã khởi tạo KHBD thành công!")
                st.download_button(
                    label="📥 TẢI FILE WORD (.DOCX) VỀ MÁY",
                    data=buffer,
                    file_name=f"KHBD_{subject}_{lesson_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Đã có lỗi xảy ra: {str(e)}")
