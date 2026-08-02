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

# ==========================================
# 1. CẤU HÌNH TRANG STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# 2. CẤU HÌNH CSS TỐI ƯU GIAO DIỆN MẪU
# ==========================================
st.markdown("""
    <style>
    /* Import font chữ hẹp & đậm giống hình mẫu */
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@700&display=swap');

    /* Lề trang gọn gàng */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1250px;
    }

    /* Tiêu đề Bước */
    .step-header {
        color: #dc2626 !important;
        font-size: 21px !important;
        font-weight: 800 !important;
        font-family: 'Roboto Condensed', sans-serif, Arial !important;
        margin-top: 12px !important;
        margin-bottom: 12px !important;
        padding-left: 8px;
        border-left: 5px solid #dc2626;
        display: flex;
        align-items: center;
    }

    /* Nhãn tiêu đề các ô (Chương:, Tên bài:, Số tiết:...) - Chữ đen, đậm */
    .custom-label {
        font-family: 'Roboto Condensed', sans-serif, Arial !important;
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #000000 !important;
        margin-bottom: 4px !important;
        display: block;
        white-space: nowrap;
    }

    /* Màu chữ XANH DƯƠNG ĐẬM bên trong các ô Nhập liệu/Selectbox */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] p,
    div[data-baseweb="input"] input,
    .stTextInput input,
    .stNumberInput input,
    .stTextArea textarea {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important; /* Màu xanh đậm đúng hình */
    }

    /* Tối ưu khoảng cách giữa các phần tử */
    .element-container {
        margin-bottom: 6px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. THANH BÊN (SIDEBAR)
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
# 4. BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)
col_sub, col_grd = st.columns(2)
with col_sub:
    st.markdown('<span class="custom-label">Môn học/Hoạt động GD:</span>', unsafe_allow_html=True)
    subject = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học"], label_visibility="collapsed")
with col_grd:
    st.markdown('<span class="custom-label">Khối lớp:</span>', unsafe_allow_html=True)
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"], index=2, label_visibility="collapsed")

# ==========================================
# 5. BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

# --- HÀNG 1: Nút Cập nhật chạy rộng full dòng ---
if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            
            # ĐÃ SỬA LỖI: Dùng {{ và }} để không bị lỗi Format specifier
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

st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

# --- HÀNG 2: CHIA 2 CỘT SONG SONG (TẢI FILE SGV | CHỌN BÀI HỌC) ---
col_h2_left, col_h2_right = st.columns(2)

with col_h2_left:
    st.markdown('<span class="custom-label">📁 Tải lên File SGV (Sách Giáo Viên - PDF hoặc Ảnh):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader("Upload SGV", type=["pdf", "png", "jpg", "jpeg"], label_visibility="collapsed")

with col_h2_right:
    st.markdown('<span class="custom-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
    
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

st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

# --- HÀNG 3: CHIA 3 CỘT (CHƯƠNG | TÊN BÀI | SỐ TIẾT) ---
col_h3_1, col_h3_2, col_h3_3 = st.columns([4.2, 4.2, 1.6])

with col_h3_1:
    st.markdown('<span class="custom-label">Chương:</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương", value=val_chap, label_visibility="collapsed")

with col_h3_2:
    st.markdown('<span class="custom-label">Tên bài:</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài", value=val_less, label_visibility="collapsed")

with col_h3_3:
    st.markdown('<span class="custom-label">Số tiết:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết", value=val_dur, min_value=1, label_visibility="collapsed")

st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

# --- HÀNG 4: YÊU CẦU CẦN ĐẠT ---
st.markdown('<span class="custom-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
requirements = st.text_area("YCĐ", value=val_req, height=110, label_visibility="collapsed")

# ==========================================
# 6. BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ</div>', unsafe_allow_html=True)
integrations = st.multiselect(
    "Lựa chọn yếu tố tích hợp:",
    ["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"],
    default=["Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra...)", "Tích hợp AI trong dạy và học (Gemini, ChatGPT, Canva...)"]
)
