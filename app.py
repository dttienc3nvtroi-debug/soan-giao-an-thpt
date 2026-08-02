# ==========================================
# BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV
# ==========================================
st.markdown('<div class="step-header">📖 BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN / CHỌN FILE SGV</div>', unsafe_allow_html=True)

# 1. HÀNG 1: NÚT BẤM CẬP NHẬT (FULL CHIỀU RỘNG TRÊN CÙNG)
if st.button("🔍 Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn", use_container_width=True):
    if not api_key:
        st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh menu bên trái trước!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            prompt_fetch = f"""Liệt kê ĐẦY ĐỦ các Bài học thuộc môn {subject} - {grade} (GDPT 2018) dạng JSON: [{"chapter": "...", "lesson": "...", "duration": 3, "req": "..."}]"""
            
            with st.spinner(f"✨ Đang đồng bộ danh mục bài học {subject} {grade}..."):
                res = model.generate_content(prompt_fetch)
                raw_text = res.text.strip()
                json_match = re.search(r'\[.*\]', raw_text, re.DOTALL)
                clean_json = json_match.group(0) if json_match else raw_text
                st.session_state['fetched_lessons'] = json.loads(clean_json)
                st.success("🎉 Đã tải xong danh mục bài học chuẩn đầy đủ!")
        except Exception as e:
            st.error(f"Lỗi khi tải danh mục bài học: {e}")

st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# 2. HÀNG 2: CHIA CỘT NGANG HÀNG (BÊN TRÁI: TẢI SGV | BÊN PHẢI: CHỌN BÀI HỌC)
col_left, col_right = st.columns(2)

# --- CỘT BÊN TRÁI ---
with col_left:
    st.markdown('<span class="field-label">📁 Tải lên File SGV (Sách Giáo Viên - PDF hoặc Ảnh):</span>', unsafe_allow_html=True)
    uploaded_sgv_file = st.file_uploader(
        "Tải lên File SGV:", 
        type=["pdf", "png", "jpg", "jpeg"], 
        label_visibility="collapsed"
    )

# --- CỘT BÊN PHẢI ---
with col_right:
    st.markdown('<span class="field-label">👉 Chọn Bài học chuẩn từ danh sách vừa tải:</span>', unsafe_allow_html=True)
    
    val_chap = "Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số"
    val_less = "Bài 2. Giá trị lớn nhất và giá trị nhỏ nhất của hàm số"
    val_dur = 3
    val_req = ""

    if 'fetched_lessons' in st.session_state and st.session_state['fetched_lessons']:
        lessons_data = st.session_state['fetched_lessons']
        lesson_titles = [f"{item['chapter']} - {item['lesson']}" for item in lessons_data]
        selected_idx = st.selectbox(
            "👉 Chọn Bài học chuẩn:", 
            range(len(lesson_titles)), 
            format_func=lambda x: lesson_titles[x],
            label_visibility="collapsed"
        )
        current_item = lessons_data[selected_idx]
        val_chap = current_item['chapter']
        val_less = current_item['lesson']
        val_dur = int(current_item['duration'])
        val_req = current_item['req']
    else:
        st.selectbox(
            "👉 Chọn Bài học chuẩn:",
            ["Chương I. Ứng dụng đạo hàm để khảo sát và vẽ đồ thị hàm số - Bài 1. Tính đơn điệu và cực trị của hàm số"],
            label_visibility="collapsed"
        )

st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# 3. HÀNG 3: CHIA 3 CỘT (CHƯƠNG | TÊN BÀI | SỐ TIẾT)
c1, c2, c3 = st.columns([4, 4, 1.5])

with c1:
    st.markdown('<span class="field-label">Chương:</span>', unsafe_allow_html=True)
    chapter_title = st.text_input("Chương", value=val_chap, label_visibility="collapsed")

with c2:
    st.markdown('<span class="field-label">Tên bài:</span>', unsafe_allow_html=True)
    lesson_title = st.text_input("Tên bài", value=val_less, label_visibility="collapsed")

with c3:
    st.markdown('<span class="field-label">Số tiết:</span>', unsafe_allow_html=True)
    duration = st.number_input("Số tiết", value=val_dur, min_value=1, label_visibility="collapsed")

st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# 4. HÀNG 4: YÊU CẦU CẦN ĐẠT
st.markdown('<span class="field-label">📌 Yêu cầu cần đạt:</span>', unsafe_allow_html=True)
requirements = st.text_area("YCĐ", value=val_req, height=120, label_visibility="collapsed")
