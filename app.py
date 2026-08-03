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
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
        max-width: 1300px;
    }
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Roboto, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    .sidebar-title {
        color: #0f172a;
        font-size: 19px !important;
        font-weight: 700;
        margin-bottom: 10px;
        border-bottom: 3px solid #2563eb;
        padding-bottom: 4px;
    }
    .step-header {
        color: #dc2626 !important;
        font-size: 21px !important;
        font-weight: 700 !important;
        margin-top: 18px !important;
        margin-bottom: 10px !important;
        padding-left: 10px;
        border-left: 5px solid #dc2626;
    }
    div[data-testid="stWidgetLabel"] p {
        font-size: 18px !important;
        font-weight: 700 !important;
        color: #1e3a8a !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# HÀM BÓC TÁCH JSON AN TOÀN
# ==========================================
def clean_and_parse_json(raw_text):
    text = raw_text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[-1]
        if text.endswith("```"):
            text = text.rsplit("```", 1)[0]
    text = text.strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    array_match = re.search(r'\[\s*\{.*\}\s*\]', text, re.DOTALL)
    if array_match:
        try:
            return json.loads(array_match.group(0))
        except Exception:
            pass

    object_match = re.search(r'\{\s*".*"\s*:.*\}', text, re.DOTALL)
    if object_match:
        try:
            return json.loads(object_match.group(0))
        except Exception:
            pass

    raise ValueError("Lỗi định dạng dữ liệu từ AI. Thầy vui lòng bấm tra cứu lại!")

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
# HÀM GỌI API GEMINI
# ==========================================
def call_gemini(api_key, preferred_model, contents, force_json=False):
    genai.configure(api_key=api_key)
    
    available_models = []
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                clean_name = m.name.replace("models/", "")
                available_models.append(clean_name)
    except Exception:
        available_models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

    pref_clean = preferred_model.replace("models/", "").strip()
    models_to_try = [pref_clean] + [m for m in available_models if m != pref_clean]

    config_dict = {"temperature": 0.1, "top_p": 0.8, "top_k": 40}
    if force_json:
        config_dict["response_mime_type"] = "application/json"
        
    generation_config = genai.types.GenerationConfig(**config_dict)

    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name=m_name, generation_config=generation_config)
            response = model.generate_content(contents)
            return response
        except Exception:
            if force_json:
                try:
                    model_retry = genai.GenerativeModel(model_name=m_name)
                    return model_retry.generate_content(contents)
                except:
                    continue
            continue
                
    raise Exception("❌ Không kết nối được API. Thầy vui lòng kiểm tra lại API Key ở menu bên trái!")

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
st.markdown("""
    <div style="text-
