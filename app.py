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

# ==========================================
# CẤU HÌNH TRANG STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Hệ thống Soạn Giáo án Tự Động 5512 (Chuẩn SGV)", 
    layout="wide", 
    page_icon="📝"
)

# ==========================================
# GIAO DIỆN & STYLE
# ==========================================
st.markdown("""<style>
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 1300px; }
html, body, [class*="css"] { font-family: 'Segoe UI', Roboto, -apple-system, BlinkMacSystemFont, Arial, sans-serif; }
section[data-testid="stSidebar"] { background-color: #f8fafc; border-right: 1px solid #e2e8f0; }
.sidebar-title { color: #0f172a; font-size: 19px !important; font-weight: 700; margin-bottom: 10px; border-bottom: 3px solid #2563eb; padding-bottom: 4px; }
.step-header { color: #dc2626 !important; font-size: 21px !important; font-weight: 700 !important; margin-top: 18px !important; margin-bottom: 10px !important; padding-left: 10px; border-left: 5px solid #dc2626; }
div[data-testid="stWidgetLabel"] p { font-size: 18px !important; font-weight: 700 !important; color: #1e3a8a !important; }
</style>""", unsafe_allow_html=True)

# ==========================================
# HÀM BÓC TÁCH JSON AN TOÀN TUYỆT ĐỐI
# ==========================================
def clean_and_parse_json(raw_text):
    text = raw_text.strip()
    
    if "```" in text:
        text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s*```$', '', text)
    text = text.strip()

    try:
        data = json.loads(text)
        if isinstance(data, dict):
            for key, val in data.items():
                if isinstance(val, list):
                    return val
            return [data]
        return data
    except Exception:
        pass

    match_array = re.search(r'\[\s*\{.*?\}\s*\]', text, re.DOTALL)
    if match_array:
        try:
            return json.loads(match_array.group(0))
        except Exception:
            pass

    match_object = re.search(r'\{\s*".*?"\s*:.*?\}', text, re.DOTALL)
    if match_object:
        try:
            res = json.loads(match_object.group(0))
            return [res] if isinstance(res, dict) else res
        except Exception:
            pass

    raise ValueError("Dữ liệu trả về chưa khớp cấu trúc. Thầy bấm tra cứu lại lần nữa nhé!")

# ==========================================
# HÀM XỬ LÝ TỆP ĐÍNH KÈM
# ==========================================
def process_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    file_type = uploaded_file.type
    bytes_data = uploaded_file.getvalue()
    if "pdf" in file_type:
        mime = "application/pdf"
    elif "png" in file_type:
        mime = "image/png"
    elif "jpeg" in file_type or "jpg" in file_type:
        mime = "image/jpeg"
    else:
        mime = file_type if file_type else "application/octet-stream"
    return {"mime_type": mime, "data": bytes_data}

# ==========================================
# HÀM GỌI API GEMINI (TỐI ƯU TỐC ĐỘ, CHỐNG TREO)
# ==========================================
def call_gemini(api_key, preferred_model, contents, force_json=False):
    genai.configure(api_key=api_key)
    
    pref_clean = preferred_model.replace("models/", "").strip()
    models_to_try = [pref_clean, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    
    seen = set()
    models_to_try = [m for m in models_to_try if not (m in seen or seen.add(m))]

    config_dict = {"temperature": 0.2}
    if force_json:
        config_dict["response_mime_type"] = "application/json"
        
    generation_config = genai.types.GenerationConfig(**config_dict)

    last_error = None
    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name=m_name, generation_config=generation_config)
            response = model.generate_content(contents)
            if response and response.text:
                return response
        except Exception as e:
            last_error = e
            if force_json:
                try:
                    model_retry = genai.GenerativeModel(model_name=m_name)
                    response = model_retry.generate_content(contents)
                    if response and response.text:
                        return response
                except Exception as e2:
                    last_error = e2
                    continue
            continue
                
    raise Exception(f"Không thể kết nối với API Gemini. Lỗi chi tiết: {last_error}")

# ==========================================
# THANH BÊN (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown('<div class="sidebar-title">🔑 ĐĂNG NHẬP & CẤU HÌNH</div>', unsafe_allow_html=True)
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="Dán API Key vào đây...")
    model_name = st.selectbox("Mô hình AI ưu tiên:", ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"], index=0)
    
    st.markdown("---")
    st.markdown('<div class="sidebar-title">👤 THÔNG TIN GIÁO VIÊN</div>', unsafe_allow_html=True)
    school_name = st.text_input("Trường THPT:", "THPT NGUYỄN VĂN TRỖI")
    dept_name = st.text_input("Tổ chuyên môn:", "TỔ TOÁN")
    teacher_name = st.text_input("Họ và tên GV:", "Dương Tấn Tiến")

# ==========================================
# TIÊU ĐỀ TRANG
# ==========================================
header_html = """<div style="text-align: center; margin-bottom: 20px; background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 16px; border-radius: 12px; border: 1px solid #bfdbfe;"><div style="font-size: 26px; font-weight: 800; color: #1e3a8a;">HỆ THỐNG SOẠN KHBD TỰ ĐỘNG CHUẨN 100% SGV (5512)</div><div style="font-size: 17px; font-weight: 600; color: #2563eb; margin-top: 4px;">📝 Tác giả: DƯƠNG TẤN TIẾN — GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI</div></div>"""
st.markdown(header_html, unsafe_allow_html=True)

# ==========================================
# BƯỚC 1: CHỌN MÔN HỌC
# ==========================================
st.markdown('<div class="step-header">📚 BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP</div>', unsafe_allow_html=True)

col_sub, col_grd, col_book = st.columns(3)
with col_sub:
    subject = st.selectbox("Môn học/Hoạt động GD:", ["Toán học", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học", "Lịch sử", "Địa lý", "Tin học", "GDKT&PL", "Công nghệ"])
with col_grd:
    grade = st.selectbox("Khối lớp:", ["Lớp 10", "Lớp 11", "Lớp 12"])
with col_book:
    book_series = st.selectbox("Bộ sách (Áp dụng từ 2026-2027):", ["Kết nối tri thức với cuộc sống"])

# ==========================================
# BƯỚC 2: XÁC ĐỊNH BÀI HỌC & MỤC TIÊU
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU BÀI HỌC NXB GIÁO DỤC & YÊU CẦU CẦN ĐẠT</div>', unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1])

clean_api_key = api_
