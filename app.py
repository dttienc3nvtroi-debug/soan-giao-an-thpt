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

# Kiểm tra và nạp pypdf an toàn (Tránh lỗi ModuleNotFoundError)
try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

# ... [Giữ nguyên phần Cấu hình giao diện] ...

# HÀM XỬ LÝ FILE TẢI LÊN (ĐÃ TỐI ƯU KHÔNG BỊ LỖI THƯ VIỆN)
def extract_file_content(uploaded_file):
    if uploaded_file is None:
        return None, ""
    
    file_type = uploaded_file.type
    bytes_data = uploaded_file.getvalue()
    
    # 1. Nếu là PDF và đã cài pypdf -> Ưu tiên đọc chữ trực tiếp
    if "pdf" in file_type:
        if PYPDF_AVAILABLE:
            try:
                pdf_reader = PdfReader(io.BytesIO(bytes_data))
                extracted_text = ""
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if text:
                        extracted_text += text + "\n"
                if extracted_text.strip():
                    return "text", extracted_text
            except Exception:
                pass
        
        # Nếu chưa cài pypdf hoặc file PDF dạng ảnh chụp -> Gửi trực tiếp dưới dạng Part cho Gemini
        return "image_part", {"mime_type": "application/pdf", "data": bytes_data}
            
    # 2. Nếu là tệp Ảnh (JPG, PNG)
    elif "image" in file_type or "jpg" in file_type or "png" in file_type or "jpeg" in file_type:
        return "image_part", {"mime_type": file_type if file_type else "image/jpeg", "data": bytes_data}
        
    return None, ""
