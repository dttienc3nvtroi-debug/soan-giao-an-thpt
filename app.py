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
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# CSS ĐIỀU CHỈNH CHÍNH XÁC THEO HÌNH ẢNH
# ==========================================
st.markdown("""
    <style>
    /* Khoảng cách lề trang */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
        max-width: 1250px;
    }

    /* Font chữ chung */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* Tiêu đề Bước 2 */
    .step-header {
        color: #dc2626 !important;
        font-size: 20px !important;
        font-weight: 700 !important;
        margin-top: 15px !important;
        margin-bottom: 15px !important;
        padding-left: 8px;
        border-left: 4px solid #dc2626;
    }

    /* Nhãn tiêu đề các ô (Chương:, Tên bài:, Số tiết:...) - Chữ đen, đậm chuẩn ảnh */
    .field-label {
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #111827 !important;
        margin-bottom: 6px !important;
        display: block;
    }

    /* Màu chữ XANH DƯƠNG ĐẬM bên trong các ô input/selectbox chuẩn ảnh */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p,
    div[data-baseweb="input"] input,
    .stTextInput input,
    .stNumberInput input {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important; /* Màu xanh đậm đúng hình */
    }

    /* Tối ưu khoảng cách giữa các hàng */
    .row-spacer {
        margin-top: 12px;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# THANH BÊN (SIDEBAR)
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
# BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
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
# BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV
# (TẠO LẠI CHÍNH XÁC 100% THEO ẢNH CHỤP)
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

# HÀNG 1: Nút bấm Cập nhật full viền
if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            prompt_fetch = f"""Liệt kê danh mục bài học môn {subject} {grade} GDPT 2018 dạng JSON array: [{"chapter": "...", "lesson": "...", "duration": 3, "req": "..."}]"""
            res = model.generate_content(prompt_fetch)
            json_match = re.search(r'\[.*\]', res.text.strip(), re.DOTALL)
            if json_match:
                st.session_state['fetched_lessons'] = json.loads(json_match.group(0))
                st.success("🎉 Đã tải xong danh mục!")
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('<div class="row-spacer"></div>', unsafe_allow_html=True)

# HÀNG 2: Chia 2 Cột (Trái: Tải File SGV | Phải: Chọn Bài học chuẩn)
col_h2_left, col_h2_right = st.columns(2)

with col_h2_left:
    st.markdown('<span class="field-label">📁 Tải lên File SGV (Sách Giáo Viên - PDF hoặc Ảnh):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Upload SGV", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

with col_h2_right:
    st.markdown('<span class="field-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
    
    val_chap = "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số"
    val_less = "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số"
    val_dur = 3
    val_req = ""

    if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
        lessons_data = st.session_state['fetched_lessons']
        lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
        selected_idx = st.selectbox("Select lesson", range(len(lesson_titles)), format_func=lambda x: lesson_titles[x], label_visibility="collapsed")
        val_chap = lessons_data[selected_idx]['chapter']
        val_less = lessons_data[selected_idx]['lesson']
        val_dur = int(lessons_data[selected_idx]['duration'])
        val_req = lessons_data[selected_idx]['req']
    else:
        st.selectbox("Select lesson default", ["Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số - Bài 1. Tính đơn điệu và cực trị của hàm số"], label_visibility="collapsed")

st.markdown('<div class="row-spacer"></div>', unsafe_allow_html=True)

# HÀNG 3: Chia 3 Cột tỉ lệ chính xác [4.2, 4.2, 1.6] khớp hoàn toàn với ảnh
col_h3_1, col_h3_2, col_h3_3 = st.columns([4.2, 4.2, 1.6])

with col_h3_1:
    st.markdown('<span class="field-label">Chương:</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương", value=val_chap, label_visibility="collapsed")

with col_h3_2:
    st.markdown('<span class="field-label">Tên bài:</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài", value=val_less, label_visibility="collapsed")

with col_h3_3:
    st.markdown('<span class="field-label">Số tiết:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết", value=val_dur, min_value=1, label_visibility="collapsed")

st.markdown('<div class="row-spacer"></div>', unsafe_allow_html=True)

# HÀNG 4: Yêu cầu cần đạt
st.markdown('<span class="field-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
requirements = st.text_area("YCĐ", value=val_req, height=120, label_visibility="collapsed")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)
