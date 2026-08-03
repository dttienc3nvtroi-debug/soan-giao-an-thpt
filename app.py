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
import time

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
# HÀM BÓC TÁCH JSON CHỐNG LỖI
# ==========================================
def clean_and_parse_json(raw_text):
    text = raw_text.strip()
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*
