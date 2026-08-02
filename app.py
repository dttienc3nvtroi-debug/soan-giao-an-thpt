import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Hệ Thống Soạn KHBD", layout="wide", page_icon="📝")

# Đọc toàn bộ giao diện HTML nguyên bản
html_code = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #F8FAFC; }
    </style>
</head>
<body class="p-4 md:p-8 flex flex-col justify-center items-center">
    <div class="w-full max-w-5xl space-y-8">
        <header class="text-center space-y-2">
            <h1 class="text-2xl md:text-3xl font-bold tracking-tight text-[#1F2937]">
                HỆ THỐNG SOẠN KHBD <span class="text-lg md:text-2xl font-semibold text-gray-600">(có tích hợp NLS, AI, STEM,...)</span>
            </h1>
            <p class="text-sm md:text-base font-semibold text-[#2563EB]">
                Tác giả: DƯƠNG TẤN TIẾN - GIÁO VIÊN TRƯỜNG THPT NGUYỄN VĂN TRỖI
            </p>
        </header>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div class="space-y-6 flex flex-col">
                <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow">
                    <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2 mb-4">
                        <span>📚</span> BƯỚC 1: CHỌN MÔN HỌC & KHỐI LỚP
                    </h2>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-[#6B7280] mb-1">Môn học/Hoạt động GD</label>
                            <select class="w-full h-11 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none text-[#1F2937]">
                                <option>Vật lí</option>
                                <option>Toán học</option>
                                <option>Hóa học</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-[#6B7280] mb-1">Khối lớp</label>
                            <select class="w-full h-11 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none text-[#1F2937]">
                                <option>Khối 10</option>
                                <option>Khối 11</option>
                                <option>Khối 12</option>
                            </select>
                        </div>
                    </div>
                </div>

                <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm flex-1 space-y-3">
                    <div class="flex items-center justify-between border-b pb-3">
                        <span class="text-xs font-semibold uppercase text-gray-400">Trạng thái hệ thống</span>
                        <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-green-50 text-green-700 border border-green-200">
                            <span class="w-2 h-2 rounded-full bg-green-500"></span> Sẵn sàng
                        </span>
                    </div>
                    <p class="text-sm text-gray-500">
                        Vui lòng hoàn tất các bước tra cứu và tích hợp năng lực trước khi tiến hành xuất file Word chuẩn 5512.
                    </p>
                </div>
            </div>

            <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow space-y-4">
                <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
                    <span>📖</span> BƯỚC 2: TRA CỨU & CHỌN BÀI HỌC CHUẨN
                </h2>
                <button class="w-full py-2.5 px-4 bg-[#2563EB] hover:bg-[#1D4ED8] text-white font-semibold text-sm rounded-xl shadow-sm transition-colors text-center">
                    Cập nhật danh sách Chương & Bài học từ taphuan.nxbgd.vn
                </button>
                <div class="p-3.5 bg-blue-50 border border-blue-100 rounded-xl flex items-start gap-2.5 text-xs text-blue-900">
                    <i class="fa-regular fa-lightbulb text-blue-600 text-base shrink-0 mt-0.5"></i>
                    <span><strong>Gợi ý:</strong> Hãy chọn Môn học và Khối lớp trước để xem danh sách.</span>
                </div>
                <div class="space-y-3">
                    <div>
                        <label class="block text-xs font-medium text-[#6B7280] mb-1">Tên Chương/Chủ đề</label>
                        <input type="text" placeholder="Nhập tên chương..." class="w-full h-10 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-medium text-[#6B7280] mb-1">Tên Bài dạy</label>
                        <input type="text" placeholder="Nhập tên bài dạy..." class="w-full h-10 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none">
                    </div>
                    <div class="grid grid-cols-3 gap-3">
                        <div class="col-span-1">
                            <label class="block text-xs font-medium text-[#6B7280] mb-1">Số tiết</label>
                            <input type="number" placeholder="2" class="w-full h-10 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none">
                        </div>
                        <div class="col-span-2">
                            <label class="block text-xs font-medium text-[#6B7280] mb-1">Yêu cầu cần đạt</label>
                            <input type="text" placeholder="Yêu cầu..." class="w-full h-10 px-3 bg-white border border-[#E5E7EB] rounded-xl text-sm focus:ring-2 focus:ring-[#2563EB] focus:outline-none">
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm hover:shadow-md transition-shadow space-y-4">
            <h2 class="text-lg font-bold text-[#1F2937] flex items-center gap-2">
                <span>🚀</span> BƯỚC 3: TÍCH HỢP NĂNG LỰC ĐẶC THÙ
            </h2>
            <div class="flex flex-wrap gap-2.5">
                <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-red-50 text-red-600 border border-red-100">
                    <i class="fa-solid fa-laptop text-red-500"></i> Năng lực Số/Ứng dụng... <button class="hover:text-red-800 ml-1">✕</button>
                </span>
                <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-red-50 text-red-600 border border-red-100">
                    <i class="fa-solid fa-microchip text-red-500"></i> Tích hợp AI trong dạy học... <button class="hover:text-red-800 ml-1">✕</button>
                </span>
                <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-red-50 text-red-600 border border-red-100">
                    <i class="fa-solid fa-users text-red-500"></i> Năng lực Hợp tác & Tự học <button class="hover:text-red-800 ml-1">✕</button>
                </span>
                <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-red-50 text-red-600 border border-red-100">
                    <i class="fa-solid fa-flask text-red-500"></i> Giáo dục STEM/STEAM <button class="hover:text-red-800 ml-1">✕</button>
                </span>
            </div>
        </div>

        <div class="pt-2 text-center">
            <button class="w-full sm:w-auto px-8 py-4 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold text-base rounded-2xl shadow-lg hover:shadow-xl transition-all transform hover:-translate-y-0.5 inline-flex items-center justify-center gap-3">
                <i class="fa-solid fa-rocket text-xl text-yellow-300"></i>
                <span>BẮT ĐẦU TẠO KHBD WORD CHUẨN 5512</span>
            </button>
        </div>
    </div>
</body>
</html>
"""

# Render nguyên bản HTML
components.html(html_code, height=900, scrolling=True)Fgemin
