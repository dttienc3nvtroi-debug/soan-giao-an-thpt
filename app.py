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
# HÀM BÓC TÁCH JSON NÂNG CẤP AN TOÀN
# ==========================================
def clean_and_parse_json(raw_text):
    text = raw_text.strip()
    
    # Loại bỏ codeblock markdown nếu có
    if "```" in text:
        text = re.sub(r'```json\s*', '', text)
        text = re.sub(r'```\s*', '', text)
    text = text.strip()

    # Thử parse trực tiếp
    try:
        return json.loads(text)
    except Exception:
        pass

    # Bóc tách lấy mảng JSON từ [...]
    match_array = re.search(r'\[\s*\{.*\}\s*\]', text, re.DOTALL)
    if match_array:
        try:
            return json.loads(match_array.group(0))
        except Exception:
            pass

    # Bóc tách lấy đối tượng JSON từ {...}
    match_object = re.search(r'\{\s*".*"\s*:.*\}', text, re.DOTALL)
    if match_object:
        try:
            res = json.loads(match_object.group(0))
            return [res] if isinstance(res, dict) else res
        except Exception:
            pass

    raise ValueError("Không thể bóc tách dữ liệu JSON từ AI. Thầy vui lòng nhấn tra cứu lại lần nữa!")

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
    book_series = st.selectbox("Bộ sách Giáo khoa:", ["Kết nối tri thức với cuộc sống", "Cánh diều", "Chân trời sáng tạo"])

# ==========================================
# BƯỚC 2: XÁC ĐỊNH BÀI HỌC & MỤC TIÊU
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU BÀI HỌC NXB GIÁO DỤC & YÊU CẦU CẦN ĐẠT</div>', unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1])

clean_api_key = api_key.strip() if api_key else ""

with col_left:
    st.subheader("1. Lấy thông tin Bài học (Chuẩn taphuan.nxbgd.vn)")
    if st.button("🔍 Tra cứu Danh sách Bài học chuẩn NXB Giáo Dục", use_container_width=True):
        if not clean_api_key:
            st.error("⚠️ Vui lòng nhập API Key ở menu bên trái!")
        else:
            with st.spinner("⚡ Đang kết nối dữ liệu NXB Giáo Dục..."):
                try:
                    prompt_lookup = f"""
                    Yêu cầu: Liệt kê danh sách các bài học môn {subject} {grade} - Bộ sách {book_series} theo chương trình GDPT 2018.
                    Định dạng trả về duy nhất là mảng JSON thuần túy (mỗi phần tử đại diện cho 1 bài học), không thêm lời dẫn hay giải thích:
                    [
                        {{
                            "chapter": "Tên Chương hoặc Tên Chủ đề",
                            "lesson": "Tên Bài học",
                            "duration": 3
                        }}
                    ]
                    """
                    res = call_gemini(clean_api_key, model_name, [prompt_lookup], force_json=True)
                    data = clean_and_parse_json(res.text)
                    if isinstance(data, list) and len(data) > 0:
                        st.session_state['ai_lessons'] = data
                        st.success(f"✅ Đã tìm thấy {len(data)} bài học chuẩn!")
                    else:
                        st.error("⚠️ Dữ liệu trả về không đúng danh sách. Thầy bấm tra cứu lại giúp em nhé!")
                except Exception as e:
                    st.error(f"❌ {e}")

    if 'ai_lessons' in st.session_state:
        lessons_list = st.session_state['ai_lessons']
        options = [f"{item.get('chapter', '')} ➔ {item.get('lesson', '')}" for item in lessons_list]
        selected_idx = st.selectbox("Chọn Bài học từ danh sách:", range(len(options)), format_func=lambda i: options[i])
        sel = lessons_list[selected_idx]
        chapter_title = st.text_input("Chương / Chủ đề:", value=sel.get('chapter', ''))
        lesson_title = st.text_input("Tên Bài dạy nguyên văn SGK:", value=sel.get('lesson', ''))
        try:
            duration = int(sel.get('duration', 3))
        except:
            duration = 3
        duration = st.number_input("Số tiết thực hiện:", value=duration, min_value=1, max_value=20)
    else:
        chapter_title = st.text_input("Chương / Chủ đề:", value="")
        lesson_title = st.text_input("Tên Bài dạy nguyên văn SGK (Hoặc nhập thủ công tại đây):", value="")
        duration = st.number_input("Số tiết thực hiện:", value=3, min_value=1, max_value=20)

with col_right:
    st.subheader("2. Xác định Yêu cầu cần đạt (Mục tiêu SGV)")
    uploaded_file = st.file_uploader("Tải lên PDF/Ảnh trang SGV (Nếu không tải, AI sẽ tự đề xuất mục tiêu):", type=["pdf", "png", "jpg", "jpeg"])
    
    requirements_text = ""
    
    if uploaded_file and clean_api_key:
        if st.button("⚡ Đọc trích xuất Mục tiêu từ File/Ảnh đính kèm", type="primary", use_container_width=True):
            with st.spinner("🔍 Đang đọc chính xác Mục tiêu từ File..."):
                try:
                    file_part = process_uploaded_file(uploaded_file)
                    prompt_extract = "Trích xuất đầy đủ, chính xác từng gạch đầu dòng các Yêu cầu cần đạt nguyên văn có trong trang SGV này."
                    res = call_gemini(clean_api_key, model_name, [file_part, prompt_extract])
                    st.session_state['extracted_reqs'] = res.text
                    st.success("✅ Trích xuất thành công!")
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}")

    if 'extracted_reqs' in st.session_state:
        requirements_text = st.text_area("Yêu cầu cần đạt (Được trích xuất từ File/Ảnh SGV):", value=st.session_state['extracted_reqs'], height=160)
    else:
        requirements_text = st.text_area("Yêu cầu cần đạt (Để trống nếu muốn AI tự đề xuất chuẩn SGV):", value="", height=160, placeholder="Thầy có thể để trống, AI sẽ tự tạo hệ thống mục tiêu chuẩn Công văn 5512 bám sát SGV...")

# ==========================================
# BƯỚC 3: TÍCH HỢP NĂNG LỰC 4.0
# ==========================================
st.markdown('<div class="step-header">🚀 BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ & YẾU TỐ CHUYỂN ĐỔI SỐ</div>', unsafe_allow_html=True)

integrations = st.multiselect(
    "Lựa chọn các yếu tố tích hợp hiện đại vào Kế hoạch bài dạy:",
    [
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)",
        "Tích hợp Công cụ AI trong dạy và học (Gemini, ChatGPT, Canva AI...)",
        "Giáo dục STEM / STEAM",
        "Phát triển Tư duy phản biện & Giải quyết vấn đề thực tậi",
        "Tích hợp Giáo dục Đạo đức, Căn cước công dân & Pháp luật"
    ],
    default=[
        "Năng lực Số / Ứng dụng CNTT (Padlet, Kahoot, Geogebra, Azota...)",
        "Tích hợp Công cụ AI trong dạy và học (Gemini, ChatGPT, Canva AI...)"
    ]
)

# ==========================================
# HÀM XUẤT FILE WORD
# ==========================================
def generate_doc(content_text, locked_chapter_title, locked_lesson_title):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.7
