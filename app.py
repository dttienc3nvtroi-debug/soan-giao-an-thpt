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
# HÀM BÓC TÁCH JSON ĐƯỢC NÂNG CẤP MẠNH MẼ (CHỐNG LỖI AI TRẢ LỜI DÀI DÒNG)
# ==========================================
def clean_and_parse_json(raw_text):
    text = raw_text.strip()
    
    # 1. Loại bỏ các thẻ code block của Markdown
    if "```" in text:
        text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s*```$', '', text)
    text = text.strip()

    # 2. Thử parse trực tiếp xem AI có trả về JSON sạch không
    try:
        data = json.loads(text)
        # Nếu AI tự bọc mảng trong 1 Object (VD: {"lessons": [...]})
        if isinstance(data, dict):
            for key, val in data.items():
                if isinstance(val, list):
                    return val
            return [data] # Fallback
        return data
    except Exception:
        pass

    # 3. Quét sâu để tìm mảng JSON trong đống text bằng Regex an toàn
    match_array = re.search(r'\[\s*\{.*?\}\s*\]', text, re.DOTALL)
    if match_array:
        try:
            return json.loads(match_array.group(0))
        except Exception:
            pass

    # 4. Quét tìm Object JSON đơn lẻ nếu AI chỉ trả 1 bài học
    match_object = re.search(r'\{\s*".*?"\s*:.*?\}', text, re.DOTALL)
    if match_object:
        try:
            res = json.loads(match_object.group(0))
            return [res] if isinstance(res, dict) else res
        except Exception:
            pass

    raise ValueError("Không thể trích xuất dữ liệu bài học. Xin thầy nhấn tra cứu lại lần nữa để AI reset dữ liệu!")

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
                    # Chạy lại không ép kiểu json xem có pass được cơ chế bảo mật của model không
                    model_retry = genai.GenerativeModel(model_name=m_name)
                    return model_retry.generate_content(contents)
                except:
                    continue
            continue
                
    raise Exception("❌ Không kết nối được API. Thầy vui lòng kiểm tra lại API Key ở menu bên trái!")
