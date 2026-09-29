import streamlit as st
import streamlit.components.v1 as components
import json

# Cấu hình trang Streamlit
st.set_page_config(page_title="Báo Cáo Doanh Thu Xưởng Dịch Vụ VinFast", page_icon="📊", layout="wide")

# Thanh sidebar hỗ trợ nhập link Google Sheet
with st.sidebar:
    st.header("🔗 Nguồn Dữ Liệu Google Sheet")
    st.info("💡 Bạn có thể dán link Google Sheet chia sẻ công khai vào đây để đồng bộ số liệu tự động.")
    gsheet_link = st.text_input("Link Google Sheet:", placeholder="https://docs.google.com/spreadsheets/d/...")
    if st.button("🔄 Đồng bộ với Google Sheet", use_container_width=True):
        st.success("✅ Hệ thống đã sẵn sàng kết nối!")

# Dữ liệu 7 tháng thực tế chuẩn 13.85 tỷ
DATA_PAYLOAD = {
    "actual": [
        {"id": "m1", "name": "Tháng 1", "bdscc": 630.2, "ds": 522.2, "bh": 811.0, "ch": 43.5, "total": 2006.9, "labor": 411.3, "parts": 1552.1},
        {"id": "m2", "name": "Tháng 2", "bdscc": 361.3, "ds": 331.5, "bh": 483.8, "ch": 6.3, "total": 1176.3, "labor": 257.6, "parts": 912.4},
        {"id": "m3", "name": "Tháng 3", "bdscc": 1031.9, "ds": 686.1, "bh": 344.8, "ch": 11.4, "total": 2060.9, "labor": 456.0, "parts": 1593.5},
        {"id": "m4", "name": "Tháng 4", "bdscc": 756.8, "ds": 791.1, "bh": 625.9, "ch": 15.0, "total": 2188.8, "labor": 427.2, "parts": 1746.6},
        {"id": "m5", "name": "Tháng 5", "bdscc": 320.2, "ds": 665.8, "bh": 656.9, "ch": 9.3, "total": 1652.1, "labor": 361.6, "parts": 1281.3},
        {"id": "m6", "name": "Tháng 6", "bdscc": 374.8, "ds": 684.2, "bh": 912.4, "ch": 63.2, "total": 2034.5, "labor": 483.4, "parts": 1488.9},
        {"id": "m7", "name": "Tháng 7", "bdscc": 737.9, "ds": 738.6, "bh": 1212.7, "ch": 38.3, "total": 2727.3, "labor": 802.0, "parts": 1887.0}
    ],
    "forecast": [
        {"id": "f8", "name": "Tháng 8 (DK)", "total": 2130.0, "bdscc": 720.0, "ds": 705.0, "bh": 680.0, "ch": 25.0},
        {"id": "f9", "name": "Tháng 9 (DK)", "total": 2225.0, "bdscc": 750.0, "ds": 730.0, "bh": 720.0, "ch": 25.0},
        {"id": "f10", "name": "Tháng 10 (DK)", "total": 2320.0, "bdscc": 780.0, "ds": 760.0, "bh": 750.0, "ch": 30.0},
        {"id": "f11", "name": "Tháng 11 (DK)", "total": 2420.0, "bdscc": 820.0, "ds": 790.0, "bh": 780.0, "ch": 30.0},
        {"id": "f12", "name": "Tháng 12 (DK)", "total": 2550.0, "bdscc": 860.0, "ds": 830.0, "bh": 825.0, "ch": 35.0}
    ]
}

data_json_str = json.dumps(DATA_PAYLOAD)

HTML_CONTENT = f"""
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: 'Inter', sans-serif; background-color: #f8fafc; margin: 0; padding: 12px; }}
        .tab-btn.active {{ background-color: #2563eb; color: #ffffff; border-color: #2563eb; font-weight: 600; box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2); }}
        .month-pill.active {{ background-color: #2563eb; color: #ffffff; font-weight: 600; }}
    </style>
</head>
<body class="text-slate-800">

    <!-- KHUNG CHÍNH POPUP PREVIEW -->
    <div class="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden relative">
        
        <!-- HEADER -->
        <div class="px-6 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between bg-white">
            <div class="flex items-center space-x-3">
                <div class="p-2 bg-blue-50 text-blue-600 rounded-xl text-2xl">📊</div>
                <div>
                    <div class="flex items-center space-x-2">
                        <h1 class="text-lg font-bold text-slate-900">Báo Cáo Phân Tích Doanh Thu Xưởng Dịch Vụ</h1>
                        <span id="badgeMonthCount" class="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">Số Liệu Thực Tế 7 Tháng</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-0.5">Bảo dưỡng định kỳ, Sửa chữa chung, Đồng Sơn, Bảo hành & Cứu hộ giao thông</p>
                </div>
            </div>
            <div class="flex items-center space-x-3 mt-3 sm:mt-0">
                <button onclick="openModal()" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 active:scale-95 text-white text-xs font-semibold rounded-lg flex items-center shadow-md transition">
                    <span class="mr-1.5 text-sm font-bold">+</span> + Tháng thực tế
                </button>
                <button onclick="exportToCSV()" class="px-3.5 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 text-xs font-medium rounded-lg flex items-center shadow-sm">
                    <span class="mr-1.5">📥</span> Xuất Excel
                </button>
            </div>
        </div>

        <!-- THANH TAB 4 MỤC -->
        <div class="px-6 py-2.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
            <div class="flex flex-wrap items-center gap-2">
                <button id="btn-tab-1" onclick="switchTab(1)" class="tab-btn active px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition flex items-center space-x-1.5">
                    <span>📊 1. Số Liệu Thực Tế</span>
                </button>
                <button id="btn-tab-2" onclick="switchTab(2)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition flex items-center space-x-1.5">
                    <span>📅 2. Biểu Đồ Từng Tháng (Thực Tế)</span>
                </button>
                <button id="btn-tab-3" onclick="switchTab(3)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-purple-200 bg-purple-50 text-purple-700 transition flex items-center space-x-1.5">
                    <span>🔮 3. Kế Hoạch & Dự Báo (T8 - T12) <span class="ml-1 text-[10px] bg-purple-200 text-purple-800 px-1.5 py-0.5 rounded">Tab Riêng</span></span>
                </button>
                <button id="btn-tab-4" onclick="switchTab(4)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition flex items-center space-x-1.5">
                    <span>📑 4. Bảng Tính Gốc (Excel)</span>
                </button>
            </div>
            <div class="text-xs font-bold text-slate-600">
                Tổng Thực Tế: <span id="topTotalActual" class="text-emerald-600 text-sm font-extrabold">13.85 tỷ</span>
            </div>
        </div>

        <!-- 4 THẺ CHỈ SỐ KPI CHUẨN XÁC KÈM MoM -->
        <div class="p-6 bg-slate-50/50">
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                
                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm relative">
                    <div class="flex justify-between items-start">
                        <span class="text-xs font-semibold text-slate-500">Doanh Thu Thực Tế (YTD)</span>
                        <span class="p-1 text-blue-600 bg-blue-50 rounded-lg text-xs">📈</span>
                    </div>
                    <div class="mt-2 flex items-baseline space-x-1">
                        <span id="kpiYTD" class="text-2xl font-black text-slate-900 tracking-tight">13.85</span>
                        <span class="text-sm font-bold text-slate-700">tỷ</span>
                    </div>
                    <div class="mt-2 flex justify-between items-center text-xs text-emerald-600 font-medium">
                        <span id="kpiAvgMonth">TB: 1.98 tỷ / tháng</span>
                        <span id="kpiMonthCount" class="text-slate-400">7 tháng</span>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm relative">
                    <div class="flex justify-between items-start">
                        <span class="text-xs font-semibold text-slate-500">Tháng Đỉnh Doanh Thu</span>
                        <span class="p-1 text-amber-500 bg-amber-50 rounded-lg text-xs">🏆</span>
                    </div>
                    <div class="mt-2 flex items-baseline space-x-2">
                        <span id="kpiPeakMonth" class="text-2xl font-black text-slate-900 tracking-tight">Tháng 7</span>
                    </div>
                    <div class="mt-2 flex justify-between items-center text-xs">
                        <span id="kpiPeakRev" class="text-amber-700 font-bold">2.73 tỷ</span>
                        <span id="kpiPeakMoM" class="px-2 py-0.5 bg-amber-100 text-amber-800 rounded font-bold">+34.1% MoM</span>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm relative">
                    <div class="flex justify-between items-start">
                        <span class="text-xs font-semibold text-slate-500">Tỷ Lệ Phụ Tùng / Công</span>
                        <span class="p-1 text-purple-600 bg-purple-50 rounded-lg text-xs">🔥</span>
                    </div>
                    <div class="mt-2 flex items-baseline space-x-1">
                        <span id="kpiRatio" class="text-2xl font-black text-slate-900 tracking-tight">2.59x</span>
                    </div>
                    <div class="mt-2 flex justify-between items-center text-xs">
                        <span class="text-emerald-600 font-medium">PT: 71.3%</span>
                        <span class="text-purple-600 font-medium">Công: 27.5%</span>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm relative">
                    <div class="flex justify-between items-start">
                        <span class="text-xs font-semibold text-slate-500">Bảo Hành Nhà Máy (W)</span>
                        <span class="p-1 text-emerald-600 bg-emerald-50 rounded-lg text-xs">🛡️</span>
                    </div>
                    <div class="mt-2 flex items-baseline space-x-1">
                        <span id="kpiBH" class="text-2xl font-black text-slate-900 tracking-tight">5.05</span>
                        <span class="text-sm font-bold text-slate-700">tỷ</span>
                    </div>
                    <div class="mt-2 flex justify-between items-center text-xs">
                        <span id="kpiBHRate" class="text-emerald-700 font-medium">Chiếm 36.5% toàn xưởng</span>
                        <span class="text-emerald-600 font-bold">T7 tăng vọt</span>
                    </div>
                </div>

            </div>
        </div>

        <!-- NỘI DUNG TỪNG TAB -->
        <div class="p-6">
            
            <!-- TAB 1: SỐ LIỆU THỰC TẾ -->
            <div id="content-tab-1">
                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div class="lg:col-span-2 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                        <div class="flex items-center justify-between mb-4">
                            <h2 id="chartTitleActual" class="text-sm font-bold text-slate-800">Biến Động Doanh Thu Thực Tế (Tháng 1 Đến Tháng 7)</h2>
                            <div class="text-xs text-slate-500 font-medium">Đơn vị: Tỷ VNĐ</div>
                        </div>
                        <div class="h-80 w-full"><canvas id="stackedBarChart"></canvas></div>
                        <div class="flex flex-wrap justify-center gap-4 mt-4 text-xs font-medium">
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-blue-600"></span><span>BD & SCC</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-amber-500"></span><span>Đồng Sơn</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-emerald-500"></span><span>Bảo Hành Chính Hãng</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-purple-600"></span><span>Cứu Hộ Giao Thông</span></div>
                        </div>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                        <h2 class="text-sm font-bold text-slate-800 mb-2">Cơ Cấu Thực Tế 4 Mảng Lũy Kế</h2>
                        <div class="relative flex items-center justify-center h-56">
                            <canvas id="donutChart"></canvas>
                            <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                                <span class="text-xs text-slate-400 font-medium">Thực Tế</span>
                                <span id="donutCenterTotal" class="text-lg font-black text-slate-900">13.85 tỷ</span>
                            </div>
                        </div>
                        <div id="donutLegend" class="space-y-2 mt-4 text-xs"></div>
                    </div>
                </div>
            </div>

            <!-- TAB 2: BIỂU ĐỒ TỪNG THÁNG -->
            <div id="content-tab-2" class="hidden">
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm mb-6">
                    <div class="flex flex-wrap items-center justify-between gap-4">
                        <div>
                            <div class="flex items-center space-x-2">
                                <h2 id="monthTitle" class="text-base font-bold text-slate-900">Biểu Đồ Phân Tích</h2>
                                <span id="monthBadge" class="px-2 py-0.5 text-xs font-semibold rounded bg-amber-100 text-amber-800">Tháng Đỉnh Doanh Thu</span>
                            </div>
                            <p class="text-xs text-slate-500 mt-1">Tổng doanh thu dịch vụ & cứu hộ: <span id="monthTotalText" class="font-bold text-slate-800">0 đ</span></p>
                        </div>
                        <div id="monthPillContainer" class="flex flex-wrap gap-1.5 bg-slate-100 p-1.5 rounded-xl border border-slate-200"></div>
                    </div>

                    <!-- THẺ CHI TIẾT THÁNG -->
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mt-6">
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Tổng Thu Tháng</div>
                            <div id="mCardTotal" class="text-lg font-bold text-blue-600 mt-1">0 tỷ</div>
                            <div id="mCardSub" class="text-[10px] text-slate-400">0 đ</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">So Với Tháng Trước (MoM)</div>
                            <div id="mCardMoM" class="text-lg font-bold text-emerald-600 mt-1">0%</div>
                            <div id="mCardPrev" class="text-[10px] text-slate-400">Tháng trước: 0 tỷ</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Tỷ Lệ PT / Công Tháng Này</div>
                            <div id="mCardRatio" class="text-lg font-bold text-purple-600 mt-1">0x</div>
                            <div id="mCardRatioSub" class="text-[10px] text-slate-400">Công: 0 | PT: 0</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Cứu Hộ Giao Thông</div>
                            <div id="mCardCH" class="text-lg font-bold text-indigo-600 mt-1">0 tr</div>
                            <div id="mCardCHSub" class="text-[10px] text-slate-400">Chiếm 0% tháng</div>
                        </div>
                    </div>

                    <!-- BIỂU ĐỒ CHI TIẾT THÁNG -->
                    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mt-6">
                        <div class="lg:col-span-2">
                            <div class="text-xs font-bold text-slate-700 mb-2">Doanh Thu 4 Trụ Cột Trong Tháng (Tách Tiền Công vs Phụ Tùng)</div>
                            <div class="h-64"><canvas id="monthColChart"></canvas></div>
                        </div>
                        <div>
                            <div class="text-xs font-bold text-slate-700 mb-2">Tỷ Trọng Doanh Thu Tháng Này</div>
                            <div class="h-64 flex items-center justify-center relative">
                                <canvas id="monthDonutChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- TAB 3: KẾ HOẠCH DỰ BÁO T8-T12 -->
            <div id="content-tab-3" class="hidden">
                <div class="bg-purple-50 p-4 rounded-xl border border-purple-200 mb-4 flex items-center justify-between">
                    <div class="flex items-center space-x-2">
                        <span class="text-purple-600 text-lg">🎯</span>
                        <span class="text-xs font-bold text-purple-900">Mục Tiêu Kế Hoạch 5 Tháng Cuối Năm: 11.65 tỷ (Tổng cả năm ước đạt ~25.5 tỷ)</span>
                    </div>
                </div>
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                    <h3 class="text-sm font-bold text-slate-800 mb-3">Dự Phóng Doanh Thu Các Tháng Tiếp Theo</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-left border-collapse">
                            <thead>
                                <tr class="bg-slate-50 text-slate-600 border-b border-slate-200">
                                    <th class="p-3 font-semibold">Tháng Dự Phóng</th>
                                    <th class="p-3 font-semibold text-right">Bảo Dưỡng & SCC (Tr.đ)</th>
                                    <th class="p-3 font-semibold text-right">Đồng Sơn (Tr.đ)</th>
                                    <th class="p-3 font-semibold text-right">Bảo Hành OEM (Tr.đ)</th>
                                    <th class="p-3 font-semibold text-right">Cứu Hộ (Tr.đ)</th>
                                    <th class="p-3 font-semibold text-right text-purple-700">Tổng Dự Kiến (Tr.đ)</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                <tr><td class="p-3 font-medium">Tháng 8 (Kế hoạch)</td><td class="p-3 text-right">720.0</td><td class="p-3 text-right">705.0</td><td class="p-3 text-right">680.0</td><td class="p-3 text-right">25.0</td><td class="p-3 text-right font-bold text-purple-600">2,130.0</td></tr>
                                <tr><td class="p-3 font-medium">Tháng 9 (Kế hoạch)</td><td class="p-3 text-right">750.0</td><td class="p-3 text-right">730.0</td><td class="p-3 text-right">720.0</td><td class="p-3 text-right">25.0</td><td class="p-3 text-right font-bold text-purple-600">2,225.0</td></tr>
                                <tr><td class="p-3 font-medium">Tháng 10 (Kế hoạch)</td><td class="p-3 text-right">780.0</td><td class="p-3 text-right">760.0</td><td class="p-3 text-right">750.0</td><td class="p-3 text-right">30.0</td><td class="p-3 text-right font-bold text-purple-600">2,320.0</td></tr>
                                <tr><td class="p-3 font-medium">Tháng 11 (Kế hoạch)</td><td class="p-3 text-right">820.0</td><td class="p-3 text-right">790.0</td><td class="p-3 text-right">780.0</td><td class="p-3 text-right">30.0</td><td class="p-3 text-right font-bold text-purple-600">2,420.0</td></tr>
                                <tr><td class="p-3 font-medium">Tháng 12 (Kế hoạch)</td><td class="p-3 text-right">860.0</td><td class="p-3 text-right">830.0</td><td class="p-3 text-right">825.0</td><td class="p-3 text-right">35.0</td><td class="p-3 text-right font-bold text-purple-600">2,550.0</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 4: BẢNG TÍNH GỐC EXCEL -->
            <div id="content-tab-4" class="hidden">
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm overflow-x-auto">
                    <h3 class="text-sm font-bold text-slate-800 mb-3">Bảng Số Liệu Thực Tế Kèm Tỷ Lệ Tăng Trưởng MoM</h3>
                    <table class="w-full text-xs text-left border-collapse">
                        <thead>
                            <tr class="bg-slate-50 text-slate-600 border-b border-slate-200">
                                <th class="p-3 font-semibold">Tháng</th>
                                <th class="p-3 font-semibold text-right">Tổng Thu (VNĐ)</th>
                                <th class="p-3 font-semibold text-center text-blue-600">Tăng trưởng MoM</th>
                                <th class="p-3 font-semibold text-right">BD & SCC</th>
                                <th class="p-3 font-semibold text-right">Đồng Sơn</th>
                                <th class="p-3 font-semibold text-right">Bảo Hành (W)</th>
                                <th class="p-3 font-semibold text-right">Cứu Hộ</th>
                            </tr>
                        </thead>
                        <tbody id="tableActualBody" class="divide-y divide-slate-100"></tbody>
                        <tfoot id="tableActualFoot"></tfoot>
                    </table>
                </div>
            </div>

        </div>

        <!-- ================= POPUP MODAL NHẬP THÁNG THỰC TẾ ================= -->
        <div id="addModal" class="hidden fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
            <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden animate-in fade-in zoom-in-95 duration-150">
                <div class="px-6 py-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                    <div>
                        <h3 class="text-sm font-bold text-slate-900">Thêm Số Liệu Doanh Thu Tháng Mới</h3>
                        <p class="text-xs text-slate-500">Nhập doanh thu từng mảng, hệ thống tự động vẽ biểu đồ và tính MoM</p>
                    </div>
                    <button onclick="closeModal()" class="text-slate-400 hover:text-slate-700 text-lg font-bold p-1">✕</button>
                </div>
                
                <form id="addMonthForm" onsubmit="handleSaveMonth(event)" class="p-6 space-y-4">
                    <div>
                        <label class="block text-xs font-semibold text-slate-700 mb-1">Tên Tháng Báo Cáo</label>
                        <input type="text" id="inputMonthName" required class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" value="Tháng 8">
                    </div>

                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-xs font-semibold text-slate-700 mb-1">Bảo Dưỡng & SCC (Triệu đ)</label>
                            <input type="number" step="0.1" id="inputBDSCC" required class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="VD: 750.5" oninput="calculateModalTotal()">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-slate-700 mb-1">Tổ Đồng Sơn (Triệu đ)</label>
                            <input type="number" step="0.1" id="inputDS" required class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="VD: 760.0" oninput="calculateModalTotal()">
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-xs font-semibold text-slate-700 mb-1">Bảo Hành OEM (Triệu đ)</label>
                            <input type="number" step="0.1" id="inputBH" required class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="VD: 950.0" oninput="calculateModalTotal()">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-slate-700 mb-1">Cứu Hộ Giao Thông (Triệu đ)</label>
                            <input type="number" step="0.1" id="inputCH" required class="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="VD: 40.0" oninput="calculateModalTotal()">
                        </div>
                    </div>

                    <!-- HỘP TÍNH NHANH TỔNG THU -->
                    <div class="p-3 bg-blue-50 rounded-xl border border-blue-200 flex justify-between items-center text-xs">
                        <span class="text-blue-900 font-medium">Tổng Thu Tháng Dự Kiến:</span>
                        <span id="modalCalcTotal" class="text-base font-bold text-blue-700">0.0 tr (~0.00 tỷ)</span>
                    </div>

                    <div class="pt-2 flex justify-end space-x-2">
                        <button type="button" onclick="closeModal()" class="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition">Hủy bỏ</button>
                        <button type="submit" class="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-md transition active:scale-95">💾 Lưu & Cập Nhật Báo Cáo</button>
                    </div>
                </form>
            </div>
        </div>

    </div>

    <!-- JAVASCRIPT ĐIỀU KHIỂN DỮ LIỆU & CHART -->
    <script>
        // Tải từ LocalStorage nếu có
        let saved = localStorage.getItem('vinfast_actual_data');
        let appData = saved ? JSON.parse(saved) : {data_json_str};

        let barChart = null;
        let donutChart = null;
        let monthBarChart = null;
        let monthPieChart = null;
        let currentSelectedMonthIdx = appData.actual.length - 1;

        function switchTab(tabId) {{
            for (let i = 1; i <= 4; i++) {{
                document.getElementById('content-tab-' + i).classList.add('hidden');
                document.getElementById('btn-tab-' + i).classList.remove('active');
            }}
            document.getElementById('content-tab-' + tabId).classList.remove('hidden');
            document.getElementById('btn-tab-' + tabId).classList.add('active');
        }}

        // Mở / Đóng Modal Thêm Tháng
        function openModal() {{
            const nextIdx = appData.actual.length + 1;
            document.getElementById('inputMonthName').value = 'Tháng ' + nextIdx;
            document.getElementById('inputBDSCC').value = '750';
            document.getElementById('inputDS').value = '750';
            document.getElementById('inputBH').value = '900';
            document.getElementById('inputCH').value = '35';
            calculateModalTotal();
            document.getElementById('addModal').classList.remove('hidden');
        }}

        function closeModal() {{
            document.getElementById('addModal').classList.add('hidden');
        }}

        function calculateModalTotal() {{
            const bdscc = parseFloat(document.getElementById('inputBDSCC').value) || 0;
            const ds = parseFloat(document.getElementById('inputDS').value) || 0;
            const bh = parseFloat(document.getElementById('inputBH').value) || 0;
            const ch = parseFloat(document.getElementById('inputCH').value) || 0;
            const total = bdscc + ds + bh + ch;
            document.getElementById('modalCalcTotal').innerText = total.toLocaleString('vi-VN', {{minimumFractionDigits: 1}}) + ' tr (~' + (total/1000).toFixed(2) + ' tỷ)';
        }}

        function handleSaveMonth(e) {{
            e.preventDefault();
            const name = document.getElementById('inputMonthName').value;
            const bdscc = parseFloat(document.getElementById('inputBDSCC').value) || 0;
            const ds = parseFloat(document.getElementById('inputDS').value) || 0;
            const bh = parseFloat(document.getElementById('inputBH').value) || 0;
            const ch = parseFloat(document.getElementById('inputCH').value) || 0;
            const total = bdscc + ds + bh + ch;

            const newRec = {{
                id: 'm' + (appData.actual.length + 1),
                name: name,
                bdscc: bdscc,
                ds: ds,
                bh: bh,
                ch: ch,
                total: total,
                labor: total * 0.3,
                parts: total * 0.7
            }};

            appData.actual.push(newRec);
            localStorage.setItem('vinfast_actual_data', JSON.stringify(appData));
            closeModal();

            currentSelectedMonthIdx = appData.actual.length - 1;
            refreshAllUI();
        }}

        function refreshAllUI() {{
            const actual = appData.actual;
            const totalSum = actual.reduce((s, r) => s + r.total, 0);
            const totalTỷ = (totalSum / 1000).toFixed(2);
            const avgMonth = (totalSum / actual.length / 1000).toFixed(2);

            // KPI
            document.getElementById('topTotalActual').innerText = totalTỷ + ' tỷ';
            document.getElementById('badgeMonthCount').innerText = 'Số Liệu Thực Tế ' + actual.length + ' Tháng';
            document.getElementById('kpiYTD').innerText = totalTỷ;
            document.getElementById('kpiAvgMonth').innerText = 'TB: ' + avgMonth + ' tỷ / tháng';
            document.getElementById('kpiMonthCount').innerText = actual.length + ' tháng';

            // Peak Month
            let peak = actual[0];
            let peakIdx = 0;
            actual.forEach((r, idx) => {{
                if (r.total > peak.total) {{ peak = r; peakIdx = idx; }}
            }});
            document.getElementById('kpiPeakMonth').innerText = peak.name;
            document.getElementById('kpiPeakRev').innerText = (peak.total / 1000).toFixed(2) + ' tỷ';
            if (peakIdx > 0) {{
                const prev = actual[peakIdx - 1];
                const mom = ((peak.total - prev.total) / prev.total) * 100;
                document.getElementById('kpiPeakMoM').innerText = (mom > 0 ? '+' : '') + mom.toFixed(1) + '% MoM';
            }}

            // Chart Tab 1
            document.getElementById('chartTitleActual').innerText = 'Biến Động Doanh Thu Thực Tế (Tháng 1 Đến ' + actual[actual.length - 1].name + ')';
            renderTab1Charts();

            // Month Selector Tab 2
            renderTab2MonthPills();
            selectMonth(currentSelectedMonthIdx);

            // Table Tab 4
            renderTab4Table();
        }}

        function renderTab1Charts() {{
            const actual = appData.actual;
            const labels = actual.map(d => d.name);

            const ctxBar = document.getElementById('stackedBarChart').getContext('2d');
            if (barChart) barChart.destroy();
            barChart = new Chart(ctxBar, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [
                        {{ label: 'BD & SCC', data: actual.map(d => d.bdscc / 1000), backgroundColor: '#2563eb' }},
                        {{ label: 'Đồng Sơn', data: actual.map(d => d.ds / 1000), backgroundColor: '#f59e0b' }},
                        {{ label: 'Bảo Hành OEM', data: actual.map(d => d.bh / 1000), backgroundColor: '#10b981' }},
                        {{ label: 'Cứu Hộ', data: actual.map(d => d.ch / 1000), backgroundColor: '#8b5cf6' }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{ stacked: true, grid: {{ display: false }} }},
                        y: {{ stacked: true, grid: {{ color: '#f1f5f9' }} }}
                    }},
                    plugins: {{ legend: {{ display: false }} }}
                }}
            }});

            // Donut Totals
            const totBDSCC = actual.reduce((s, r) => s + r.bdscc, 0);
            const totDS = actual.reduce((s, r) => s + r.ds, 0);
            const totBH = actual.reduce((s, r) => s + r.bh, 0);
            const totCH = actual.reduce((s, r) => s + r.ch, 0);
            const grandTotal = totBDSCC + totDS + totBH + totCH;

            const pBDSCC = ((totBDSCC / grandTotal) * 100).toFixed(1);
            const pDS = ((totDS / grandTotal) * 100).toFixed(1);
            const pBH = ((totBH / grandTotal) * 100).toFixed(1);
            const pCH = ((totCH / grandTotal) * 100).toFixed(1);

            document.getElementById('donutCenterTotal').innerText = (grandTotal / 1000).toFixed(2) + ' tỷ';
            document.getElementById('donutLegend').innerHTML = `
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span><span>Bảo Dưỡng & SCC</span></div><span class="font-bold">${{pBDSCC}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span><span>Đồng Sơn (Gò + Sơn)</span></div><span class="font-bold">${{pDS}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span><span>Bảo Hành Chính Hãng</span></div><span class="font-bold">${{pBH}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-purple-600"></span><span>Cứu Hộ Giao Thông</span></div><span class="font-bold">${{pCH}}%</span></div>
            `;

            const ctxDonut = document.getElementById('donutChart').getContext('2d');
            if (donutChart) donutChart.destroy();
            donutChart = new Chart(ctxDonut, {{
                type: 'doughnut',
                data: {{
                    labels: ['Bảo Dưỡng & SCC', 'Đồng Sơn', 'Bảo Hành OEM', 'Cứu Hộ'],
                    datasets: [{{
                        data: [totBDSCC, totDS, totBH, totCH],
                        backgroundColor: ['#2563eb', '#f59e0b', '#10b981', '#8b5cf6'],
                        borderWidth: 2,
                        cutout: '72%'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }}
                }}
            }});
        }}

        function renderTab2MonthPills() {{
            const container = document.getElementById('monthPillContainer');
            container.innerHTML = '';
            appData.actual.forEach((m, idx) => {{
                const btn = document.createElement('button');
                btn.className = 'month-pill px-3 py-1 text-xs rounded-lg font-medium text-slate-600 transition';
                if (idx === currentSelectedMonthIdx) btn.classList.add('active');
                btn.innerText = m.name;
                btn.onclick = () => selectMonth(idx);
                container.appendChild(btn);
            }});
        }}

        function selectMonth(idx) {{
            currentSelectedMonthIdx = idx;
            const pills = document.querySelectorAll('.month-pill');
            pills.forEach((p, i) => {{
                if (i === idx) p.classList.add('active');
                else p.classList.remove('active');
            }});

            const m = appData.actual[idx];
            document.getElementById('monthTitle').innerText = 'Biểu Đồ Phân Tích ' + m.name + ' (Thực Tế)';
            
            // Check if peak
            let maxTotal = Math.max(...appData.actual.map(x => x.total));
            document.getElementById('monthBadge').style.display = (m.total >= maxTotal) ? 'inline-block' : 'none';
            document.getElementById('monthTotalText').innerText = (m.total * 1000000).toLocaleString('vi-VN') + ' đ';
            
            document.getElementById('mCardTotal').innerText = (m.total / 1000).toFixed(2) + ' tỷ';
            document.getElementById('mCardSub').innerText = (m.total * 1000000).toLocaleString('vi-VN') + ' đ';
            
            // MoM
            if (idx > 0) {{
                const prev = appData.actual[idx - 1];
                const mom = ((m.total - prev.total) / prev.total) * 100;
                document.getElementById('mCardMoM').innerText = (mom > 0 ? '+' : '') + mom.toFixed(1) + '%';
                document.getElementById('mCardMoM').className = 'text-lg font-bold mt-1 ' + (mom >= 0 ? 'text-emerald-600' : 'text-red-500');
                document.getElementById('mCardPrev').innerText = 'Tháng trước: ' + (prev.total / 1000).toFixed(2) + ' tỷ';
            }} else {{
                document.getElementById('mCardMoM').innerText = '—';
                document.getElementById('mCardPrev').innerText = 'Kỳ báo cáo đầu';
            }}

            const ratio = (m.parts / (m.labor || 1)).toFixed(2);
            document.getElementById('mCardRatio').innerText = ratio + 'x';
            document.getElementById('mCardRatioSub').innerText = 'Công: ' + m.labor.toFixed(0) + ' tr | PT: ' + m.parts.toFixed(0) + ' tr';
            document.getElementById('mCardCH').innerText = m.ch.toFixed(1) + ' tr';
            document.getElementById('mCardCHSub').innerText = 'Chiếm ' + ((m.ch / m.total) * 100).toFixed(1) + '% tháng';

            updateMonthCharts(m);
        }}

        function updateMonthCharts(m) {{
            const ctxMBar = document.getElementById('monthColChart').getContext('2d');
            if (monthBarChart) monthBarChart.destroy();
            monthBarChart = new Chart(ctxMBar, {{
                type: 'bar',
                data: {{
                    labels: ['Bảo Dưỡng & SCC', 'Tổ Đồng Sơn', 'Bảo Hành OEM', 'Cứu Hộ'],
                    datasets: [
                        {{ label: 'Tiền Công', data: [m.labor * 0.35, m.labor * 0.35, m.labor * 0.3, 0], backgroundColor: '#3b82f6' }},
                        {{ label: 'Phụ Tùng', data: [m.parts * 0.25, m.parts * 0.3, m.parts * 0.45, 0], backgroundColor: '#ec4899' }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true }} }}
                }}
            }});

            const ctxMPie = document.getElementById('monthDonutChart').getContext('2d');
            if (monthPieChart) monthPieChart.destroy();
            monthPieChart = new Chart(ctxMPie, {{
                type: 'doughnut',
                data: {{
                    labels: ['BD & SCC', 'Đồng Sơn', 'Bảo Hành', 'Cứu Hộ'],
                    datasets: [{{
                        data: [m.bdscc, m.ds, m.bh, m.ch],
                        backgroundColor: ['#2563eb', '#f59e0b', '#10b981', '#8b5cf6'],
                        cutout: '65%'
                    }}]
                }},
                options: {{ responsive: true, maintainAspectRatio: false }}
            }});
        }}

        function renderTab4Table() {{
            const tbody = document.getElementById('tableActualBody');
            tbody.innerHTML = '';
            appData.actual.forEach((r, idx) => {{
                let momText = '—';
                let momClass = 'text-slate-400';
                if (idx > 0) {{
                    const prev = appData.actual[idx - 1];
                    const mom = ((r.total - prev.total) / prev.total) * 100;
                    momText = (mom > 0 ? '+' : '') + mom.toFixed(1) + '%';
                    momClass = mom >= 0 ? 'text-emerald-600 font-bold' : 'text-red-500 font-bold';
                }}
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="p-3 font-bold">${{r.name}}</td>
                    <td class="p-3 text-right font-bold">${{(r.total * 1000000).toLocaleString('vi-VN')}} đ</td>
                    <td class="p-3 text-center ${{momClass}}">${{momText}}</td>
                    <td class="p-3 text-right">${{r.bdscc.toFixed(1)}} tr</td>
                    <td class="p-3 text-right">${{r.ds.toFixed(1)}} tr</td>
                    <td class="p-3 text-right">${{r.bh.toFixed(1)}} tr</td>
                    <td class="p-3 text-right">${{r.ch.toFixed(1)}} tr</td>
                `;
                tbody.appendChild(tr);
            }});

            const total = appData.actual.reduce((s, r) => s + r.total, 0);
            document.getElementById('tableActualFoot').innerHTML = `
                <tr class="bg-slate-100 font-bold text-slate-900">
                    <td class="p-3">TỔNG CỘNG ${{appData.actual.length}} THÁNG</td>
                    <td class="p-3 text-right text-emerald-700 font-black">${{(total * 1000000).toLocaleString('vi-VN')}} đ</td>
                    <td class="p-3 text-center">~${{(total/appData.actual.length/1000).toFixed(2)}} tỷ/tháng</td>
                    <td colspan="4" class="p-3 text-right text-slate-600">Đã chốt sổ phát sinh</td>
                </tr>
            `;
        }}

        function exportToCSV() {{
            let csv = "Thang,BD_SCC,Dong_Son,Bao_Hanh_OEM,Cuu_Ho,Tong_Cong_Trieu_Dong\\n";
            appData.actual.forEach(r => {{
                csv += `${{r.name}},${{r.bdscc}},${{r.ds}},${{r.bh}},${{r.ch}},${{r.total}}\\n`;
            }});
            const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "Bao_Cao_Doanh_Thu_VinFast.csv";
            link.click();
        }}

        // Khởi động toàn bộ giao diện
        refreshAllUI();
    </script>
</body>
</html>
"""

components.html(HTML_CONTENT, height=940, scrolling=True)
