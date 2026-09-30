import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import re

st.set_page_config(page_title="Báo Cáo Doanh Thu Dịch Vụ VinFast", page_icon="📊", layout="wide")

# Dữ liệu chuẩn 100% từng đồng của 7 tháng thực tế (13.85 tỷ)
RAW_ACTUAL_DATA = [
    {
        "id": "m1", "name": "Tháng 1",
        "congBD": 105552349, "congSCC": 132185890, "ptBDSCC": 392433779,
        "congGo": 33274750, "congSon": 67965091, "tongCongDS": 101239841, "ptDS": 420983471,
        "congBH": 177852500, "ptBH": 633141964, "tong": 1963389794,
        "cuuHo": 43466800, "total": 2006856594
    },
    {
        "id": "m2", "name": "Tháng 2",
        "congBD": 61107500, "congSCC": 95955155, "ptBDSCC": 204226871,
        "congGo": 23450000, "congSon": 47083611, "tongCongDS": 70533611, "ptDS": 261003264,
        "congBH": 126725000, "ptBH": 357115264, "tong": 1169999998,
        "cuuHo": 6292785, "total": 1176292783
    },
    {
        "id": "m3", "name": "Tháng 3",
        "congBD": 96957944, "congSCC": 163173190, "ptBDSCC": 771805214,
        "congGo": 80376666, "congSon": 119426804, "tongCongDS": 199803470, "ptDS": 486322661,
        "congBH": 56215000, "ptBH": 288563726, "tong": 2049507871,
        "cuuHo": 11399995, "total": 2060907866
    },
    {
        "id": "m4", "name": "Tháng 4",
        "congBD": 136176250, "congSCC": 109065486, "ptBDSCC": 511523675,
        "congGo": 55667195, "congSon": 128411661, "tongCongDS": 184078856, "ptDS": 607035927,
        "congBH": 126402500, "ptBH": 499518732, "tong": 2173801426,
        "cuuHo": 15040600, "total": 2188842026
    },
    {
        "id": "m5", "name": "Tháng 5",
        "congBD": 85586250, "congSCC": 54650487, "ptBDSCC": 179987803,
        "congGo": 68734375, "congSon": 152589500, "tongCongDS": 221323875, "ptDS": 444427264,
        "congBH": 170097500, "ptBH": 486790109, "tong": 1642863288,
        "cuuHo": 9252000, "total": 1652115288
    },
    {
        "id": "m6", "name": "Tháng 6",
        "congBD": 112101923, "congSCC": 49094157, "ptBDSCC": 213555761,
        "congGo": 82175519, "congSon": 136316033, "tongCongDS": 218491552, "ptDS": 465689781,
        "congBH": 152840000, "ptBH": 759515486, "tong": 1971288660,
        "cuuHo": 63179000, "total": 2034467660
    },
    {
        "id": "m7", "name": "Tháng 7",
        "congBD": 103387506, "congSCC": 223356704, "ptBDSCC": 411113918,
        "congGo": 96945250, "congSon": 153434000, "tongCongDS": 250379250, "ptDS": 488189498,
        "congBH": 224862500, "ptBH": 987791454, "tong": 2689080830,
        "cuuHo": 38256000, "total": 2727336830
    }
]

# Hàm tự động đọc Google Sheet tab Data_doanhthu
def fetch_from_gsheet(url):
    try:
        match_id = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
        if not match_id:
            return None
        sheet_id = match_id.group(1)
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=Data_doanhthu"
        df = pd.read_csv(csv_url)
        if "Thang" in df.columns:
            res = []
            for _, r in df.iterrows():
                cbd = float(r.get("CongBaoDuong", 0))
                cscc = float(r.get("CongSCC", 0))
                ptbdscc = float(r.get("PhuTungBDSCC", 0))
                cgo = float(r.get("CongGo", 0))
                cson = float(r.get("CongSon", 0))
                tongds = float(r.get("TongCongDS", cgo + cson))
                ptds = float(r.get("PtDongSon", 0))
                cbh = float(r.get("CongBaoHanh", 0))
                ptbh = float(r.get("PtBaoHanh", 0))
                cho = float(r.get("CuuHo", 0))
                tot = float(r.get("TongCongCuuHo", cbd + cscc + ptbdscc + tongds + ptds + cbh + ptbh + cho))
                tong_t = float(r.get("Tong", tot - cho))
                res.append({
                    "id": str(r.get("Thang")),
                    "name": str(r.get("Thang")),
                    "congBD": cbd, "congSCC": cscc, "ptBDSCC": ptbdscc,
                    "congGo": cgo, "congSon": cson, "tongCongDS": tongds, "ptDS": ptds,
                    "congBH": cbh, "ptBH": ptbh, "tong": tong_t, "cuuHo": cho, "total": tot
                })
            return res
    except Exception:
        pass
    return None

# Sidebar quản lý đồng bộ
with st.sidebar:
    st.header("🔗 Nguồn Dữ Liệu Google Sheet")
    st.info("💡 Bạn chỉ cần dán link Google Sheet một lần, hệ thống sẽ tự động lưu và kéo số liệu mỗi khi reboot.")
    saved_url = st.session_state.get("gsheet_url_saved", "")
    gsheet_input = st.text_input("Link Google Sheet (tab Data_doanhthu):", value=saved_url, placeholder="https://docs.google.com/spreadsheets/d/...")
    if gsheet_input:
        st.session_state["gsheet_url_saved"] = gsheet_input
        loaded_sheet = fetch_from_gsheet(gsheet_input)
        if loaded_sheet:
            st.success(f"✅ Đã kéo thành công {len(loaded_sheet)} tháng từ Google Sheet!")
            RAW_ACTUAL_DATA = loaded_sheet
        else:
            st.warning("Đang dùng dữ liệu chuẩn hệ thống. Đảm bảo sheet có tên tab là 'Data_doanhthu'.")

data_payload = {
    "actual": RAW_ACTUAL_DATA,
    "forecast": [
        {"name": "Tháng 8 (DK)", "total": 2130000000, "bdscc": 720000000, "ds": 705000000, "bh": 680000000, "ch": 25000000},
        {"name": "Tháng 9 (DK)", "total": 2225000000, "bdscc": 750000000, "ds": 730000000, "bh": 720000000, "ch": 25000000},
        {"name": "Tháng 10 (DK)", "total": 2320000000, "bdscc": 780000000, "ds": 760000000, "bh": 750000000, "ch": 30000000},
        {"name": "Tháng 11 (DK)", "total": 2420000000, "bdscc": 820000000, "ds": 790000000, "bh": 780000000, "ch": 30000000},
        {"name": "Tháng 12 (DK)", "total": 2550000000, "bdscc": 860000000, "ds": 830000000, "bh": 825000000, "ch": 35000000}
    ]
}

data_json_str = json.dumps(data_payload)

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
        .chart-toggle-btn.active {{ background-color: #ffffff; font-weight: 700; color: #1e293b; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        input::-webkit-outer-spin-button, input::-webkit-inner-spin-button {{ -webkit-appearance: none; margin: 0; }}
    </style>
</head>
<body class="text-slate-800">

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
                <button id="btn-tab-1" onclick="switchTab(1)" class="tab-btn active px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition">
                    📊 1. Số Liệu Thực Tế (7 Tháng)
                </button>
                <button id="btn-tab-2" onclick="switchTab(2)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition">
                    📅 2. Biểu Đồ Từng Tháng (Thực Tế)
                </button>
                <button id="btn-tab-3" onclick="switchTab(3)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-purple-200 bg-purple-50 text-purple-700 transition">
                    🔮 3. Kế Hoạch & Dự Báo (T8 - T12) <span class="ml-1 text-[10px] bg-purple-200 text-purple-800 px-1.5 py-0.5 rounded">Tab Riêng</span>
                </button>
                <button id="btn-tab-4" onclick="switchTab(4)" class="tab-btn px-4 py-2 text-xs rounded-xl border border-slate-200 bg-white text-slate-700 transition">
                    📑 4. Bảng Tính Gốc (Excel)
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
                        <span id="kpiPtRatio" class="text-emerald-600 font-medium">PT: 71.3%</span>
                        <span id="kpiLaborRatio" class="text-purple-600 font-medium">Công: 27.5%</span>
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
            
            <!-- TAB 1: SỐ LIỆU THỰC TẾ CÓ 3 NÚT CHUYỂN ĐỔI CHẾ ĐỘ BIỂU ĐỒ (ẢNH 1) -->
            <div id="content-tab-1">
                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div class="lg:col-span-2 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                            <h2 id="chartTitleActual" class="text-sm font-bold text-slate-800">Biến Động Doanh Thu Thực Tế (Tháng 1 Đến Tháng 7)</h2>
                            <!-- 3 NÚT CHUYỂN CHẾ ĐỘ BIỂU ĐỒ -->
                            <div class="flex items-center p-0.5 bg-slate-100 rounded-lg text-xs border border-slate-200">
                                <button id="btnModePillars" onclick="changeChartMode('pillars')" class="chart-toggle-btn active px-3 py-1 rounded-md text-slate-700 transition">
                                    4 Mảng Dịch Vụ
                                </button>
                                <button id="btnModeLaborParts" onclick="changeChartMode('labor-parts')" class="chart-toggle-btn px-3 py-1 rounded-md text-slate-600 transition">
                                    Công vs Phụ Tùng
                                </button>
                                <button id="btnModeLine" onclick="changeChartMode('line')" class="chart-toggle-btn px-3 py-1 rounded-md text-slate-600 transition">
                                    Đường Tổng Thu
                                </button>
                            </div>
                        </div>
                        <div class="h-80 w-full"><canvas id="mainChartTab1"></canvas></div>
                        <div id="legendModePillars" class="flex flex-wrap justify-center gap-4 mt-4 text-xs font-medium">
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-blue-600"></span><span>BD & SCC</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-amber-500"></span><span>Đồng Sơn</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-emerald-500"></span><span>Bảo Hành Chính Hãng</span></div>
                            <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-purple-600"></span><span>Cứu Hộ Giao Thông</span></div>
                        </div>
                    </div>

                    <!-- DONUT LŨY KẾ 4 MẢNG (HIỂN THỊ ĐẦY ĐỦ TIỀN VÀ %) -->
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                        <h2 class="text-sm font-bold text-slate-800 mb-2">Cơ Cấu Thực Tế 4 Mảng Lũy Kế</h2>
                        <div class="relative flex items-center justify-center h-52">
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

            <!-- TAB 2: BIỂU ĐỒ TỪNG THÁNG CÓ CỘT ĐÔI + BẢNG XẾP HẠNG 9 HẠNG MỤC (ẢNH 3, 4, 5) -->
            <div id="content-tab-2" class="hidden space-y-6">
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                    <div class="flex flex-wrap items-center justify-between gap-4">
                        <div>
                            <div class="flex items-center space-x-2">
                                <h2 id="monthTitle" class="text-base font-bold text-slate-900">Biểu Đồ Phân Tích Tháng 7 (Thực Tế)</h2>
                                <span id="monthBadge" class="px-2 py-0.5 text-xs font-semibold rounded bg-amber-100 text-amber-800">Tháng Đỉnh Doanh Thu</span>
                            </div>
                            <p class="text-xs text-slate-500 mt-1">Tổng doanh thu dịch vụ & cứu hộ: <span id="monthTotalText" class="font-bold text-slate-800">2.727.336.830 đ</span></p>
                        </div>
                        <div id="monthPillContainer" class="flex flex-wrap gap-1.5 bg-slate-100 p-1.5 rounded-xl border border-slate-200"></div>
                    </div>

                    <!-- 4 THẺ CHI TIẾT THÁNG -->
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mt-6">
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Tổng Thu Tháng</div>
                            <div id="mCardTotal" class="text-lg font-bold text-blue-600 mt-1">2.73 tỷ</div>
                            <div id="mCardSub" class="text-[10px] text-slate-400">2.727.336.830 đ</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">So Với Tháng Trước (MoM)</div>
                            <div id="mCardMoM" class="text-lg font-bold text-emerald-600 mt-1">+34.1%</div>
                            <div id="mCardPrev" class="text-[10px] text-slate-400">Tháng trước: 2.03 tỷ</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Tỷ Lệ PT / Công Tháng Này</div>
                            <div id="mCardRatio" class="text-lg font-bold text-purple-600 mt-1">2.35x</div>
                            <div id="mCardRatioSub" class="text-[10px] text-slate-400">Công: 802 tr | PT: 1.89 tỷ</div>
                        </div>
                        <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200">
                            <div class="text-[11px] text-slate-500 font-semibold">Cứu Hộ Giao Thông</div>
                            <div id="mCardCH" class="text-lg font-bold text-indigo-600 mt-1">38.3 tr</div>
                            <div id="mCardCHSub" class="text-[10px] text-slate-400">Chiếm 1.4%</div>
                        </div>
                    </div>

                    <!-- 2 BIỂU ĐỒ CHI TIẾT (CỘT ĐÔI CÔNG/PHỤ TÙNG vs TỔNG CẢ MẢNG - ẢNH 3) -->
                    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mt-6">
                        <div class="lg:col-span-2">
                            <div class="flex justify-between items-center mb-2">
                                <div class="text-xs font-bold text-slate-700">Doanh Thu 4 Trụ Cột Trong Tháng (Tách Tiền Công vs Phụ Tùng)</div>
                                <div class="text-[11px] text-slate-400 font-mono">Đơn vị: VNĐ</div>
                            </div>
                            <div class="h-64"><canvas id="monthDoubleColChart"></canvas></div>
                            <div class="flex justify-center items-center gap-5 mt-3 text-xs">
                                <div class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-[#ec4899]"></span><span>Phụ Tùng</span></div>
                                <div class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-[#3b82f6]"></span><span>Tiền Công</span></div>
                                <div class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-[#94a3b8]"></span><span>Tổng Cả Mảng</span></div>
                            </div>
                        </div>
                        <div>
                            <div class="text-xs font-bold text-slate-700 mb-2">Tỷ Trọng Doanh Thu Tháng Này</div>
                            <div class="h-44 flex items-center justify-center relative">
                                <canvas id="monthDonutChart"></canvas>
                                <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                                    <span id="donutMonthCenterName" class="text-[10px] text-slate-400">Tháng 7</span>
                                    <span id="donutMonthCenterTotal" class="text-xs font-bold text-slate-900">2.73 tỷ</span>
                                </div>
                            </div>
                            <!-- DANH MỤC TỶ TRỌNG CÓ ĐẦY ĐỦ TIỀN VÀ % (ẢNH 5) -->
                            <div id="monthDonutList" class="space-y-1.5 pt-3 border-t border-slate-100 text-xs mt-2"></div>
                        </div>
                    </div>
                </div>

                <!-- BẢNG XẾP HẠNG 9 HẠNG MỤC THỰC TẾ TRONG THÁNG TỪ CAO XUỐNG THẤP (ẢNH 4) -->
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                    <h3 id="rankTitle" class="text-xs font-bold text-slate-800 mb-3">
                        Bảng Xếp Hạng Hạng Mục Thực Tế Trong Tháng 7 (Từ Cao Xuống Thấp)
                    </h3>
                    <div id="rankGridContainer" class="grid grid-cols-1 sm:grid-cols-3 gap-3"></div>
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
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm overflow-x-auto">
                    <h3 class="text-sm font-bold text-slate-800 mb-3">Dự Phóng Doanh Thu Các Tháng Tiếp Theo</h3>
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

            <!-- TAB 4: BẢNG TÍNH GỐC EXCEL ĐẦY ĐỦ 13 CỘT (ẢNH 7) -->
            <div id="content-tab-4" class="hidden">
                <div class="bg-white border border-slate-300 rounded-xl shadow-sm overflow-hidden">
                    <div class="p-3.5 bg-slate-100 border-b border-slate-300 flex items-center justify-between text-xs">
                        <span class="font-bold text-slate-800">
                            Bảng Tính Doanh Thu Thực Tế (Chỉ gồm các tháng phát sinh thực tế)
                        </span>
                        <span class="text-slate-500 font-mono">Đơn vị: VNĐ</span>
                    </div>

                    <div class="overflow-x-auto">
                        <table class="w-full border-collapse text-[11px] text-slate-800">
                            <thead>
                                <tr class="bg-slate-200 border-b border-slate-300 divide-x divide-slate-300 text-center font-bold">
                                    <th class="py-2.5 px-3 text-left bg-slate-300 sticky left-0 z-10 w-24">Tháng</th>
                                    <th class="py-2 px-2 bg-emerald-100 text-emerald-950 min-w-[95px]">Công BD</th>
                                    <th class="py-2 px-2 min-w-[95px]">Công SCC</th>
                                    <th class="py-2 px-2 min-w-[105px]">PT BD+SCC</th>
                                    <th class="py-2 px-2 min-w-[90px]">Công Gò</th>
                                    <th class="py-2 px-2 min-w-[95px]">Công Sơn</th>
                                    <th class="py-2 px-2 bg-blue-100 text-blue-950 min-w-[105px]">Tổng Công ĐS*</th>
                                    <th class="py-2 px-2 min-w-[105px]">PT Đồng Sơn</th>
                                    <th class="py-2 px-2 min-w-[100px]">Công BH</th>
                                    <th class="py-2 px-2 min-w-[105px]">PT BH</th>
                                    <th class="py-2 px-2 bg-slate-300 font-extrabold min-w-[115px]">Tổng*</th>
                                    <th class="py-2 px-2 min-w-[90px]">Cứu hộ</th>
                                    <th class="py-2 px-2 bg-amber-200 text-amber-950 font-extrabold min-w-[130px]">
                                        Tổng Cộng Cứu Hộ*
                                    </th>
                                </tr>
                            </thead>
                            <tbody id="table13ColsBody" class="divide-y divide-slate-200 font-mono"></tbody>
                            <tfoot id="table13ColsFoot" class="font-mono font-bold"></tfoot>
                        </table>
                    </div>
                </div>
            </div>

        </div>

        <!-- ================= POPUP MODAL NHẬP THÁNG ĐẦY ĐỦ 3 KHỐI (ẢNH 1 CỦA USER) ================= -->
        <div id="addModal" class="hidden fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto text-slate-200 animate-in fade-in zoom-in-95 duration-150">
                
                <div class="flex items-center justify-between p-5 border-b border-slate-800">
                    <div>
                        <h3 class="text-base font-bold text-white">Thêm Số Liệu Doanh Thu Tháng Mới</h3>
                        <p class="text-xs text-slate-400 mt-0.5">Nhập các trường tiền công và phụ tùng, hệ thống tự động cộng dồn doanh số</p>
                    </div>
                    <button onclick="closeModal()" class="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition">✕</button>
                </div>

                <form onsubmit="handleSaveMonth(event)" class="p-5 space-y-4">
                    <div class="flex items-center justify-between gap-4">
                        <div class="w-1/2">
                            <label class="block text-xs font-semibold text-slate-300 mb-1">Tên Tháng</label>
                            <input type="text" id="inThang" required class="w-full px-3 py-2 text-xs bg-slate-800/80 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="Tháng 8">
                        </div>
                        <div class="pt-5">
                            <button type="button" onclick="copyFromLastMonth()" class="text-xs font-medium text-blue-400 hover:text-blue-300 flex items-center space-x-1.5 transition">
                                <span>✨</span>
                                <span id="copyBtnLabel">Sao chép số từ Tháng 7</span>
                            </button>
                        </div>
                    </div>

                    <!-- 1. BẢO DƯỠNG & SỬA CHỮA CHUNG -->
                    <div class="p-4 bg-slate-800/40 rounded-xl border border-slate-700/60 space-y-3">
                        <div class="text-xs font-bold text-blue-400 tracking-wide uppercase">1. BẢO DƯỠNG & SỬA CHỮA CHUNG</div>
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Công Bảo Dưỡng</label>
                                <input type="number" id="inCongBD" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="110000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Công SCC</label>
                                <input type="number" id="inCongSCC" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="150000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Phụ tùng BD+SCC</label>
                                <input type="number" id="inPtBDSCC" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="350000000" oninput="calcFormTotals()">
                            </div>
                        </div>
                    </div>

                    <!-- 2. TỔ ĐỒNG SƠN (BODY & PAINT) -->
                    <div class="p-4 bg-slate-800/40 rounded-xl border border-slate-700/60 space-y-3">
                        <div class="text-xs font-bold text-amber-400 tracking-wide uppercase">2. TỔ ĐỒNG SƠN (BODY & PAINT)</div>
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Công Gò</label>
                                <input type="number" id="inCongGo" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="70000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Công Sơn</label>
                                <input type="number" id="inCongSon" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="130000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">PT Đồng Sơn</label>
                                <input type="number" id="inPtDS" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="450000000" oninput="calcFormTotals()">
                            </div>
                        </div>
                    </div>

                    <!-- 3. BẢO HÀNH & CỨU HỘ -->
                    <div class="p-4 bg-slate-800/40 rounded-xl border border-slate-700/60 space-y-3">
                        <div class="text-xs font-bold text-emerald-400 tracking-wide uppercase">3. BẢO HÀNH & CỨU HỘ</div>
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Công Bảo Hành</label>
                                <input type="number" id="inCongBH" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="160000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Phụ Tùng Bảo Hành</label>
                                <input type="number" id="inPtBH" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="500000000" oninput="calcFormTotals()">
                            </div>
                            <div>
                                <label class="block text-[11px] text-slate-400 mb-1">Cứu Hộ 24/7</label>
                                <input type="number" id="inCuuHo" class="w-full px-3 py-1.5 text-xs bg-slate-800 border border-slate-700 rounded-lg text-white outline-none focus:border-blue-500" value="25000000" oninput="calcFormTotals()">
                            </div>
                        </div>
                    </div>

                    <div class="p-3.5 bg-blue-950/40 rounded-xl border border-blue-900/60 flex flex-wrap justify-between items-center text-xs">
                        <span class="text-slate-300">Tổng công ĐS (Gò+Sơn): <b id="formTotalDS" class="text-white font-bold">200.000.000 đ</b></span>
                        <span class="text-slate-300">Tổng Doanh Thu Tháng: <b id="formGrandTotal" class="text-blue-400 font-extrabold text-sm">1.945.000.000 đ</b></span>
                    </div>

                    <div class="pt-3 flex justify-end space-x-3">
                        <button type="button" onclick="closeModal()" class="px-4 py-2 text-xs font-semibold text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-xl transition">Hủy Bỏ</button>
                        <button type="submit" class="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-500 rounded-xl shadow-lg transition active:scale-95 flex items-center space-x-1.5">
                            <span>+</span>
                            <span>Thêm Vào Bảng Tính</span>
                        </button>
                    </div>
                </form>
            </div>
        </div>

    </div>

    <!-- SCRIPT ĐIỀU KHIỂN ĐẦY ĐỦ TÍNH NĂNG -->
    <script>
        let appData = {data_json_str};

        let currentChartMode = 'pillars';
        let mainChartInstance = null;
        let donutChartInstance = null;
        let monthDoubleColInstance = null;
        let monthDonutInstance = null;
        let currentSelectedMonthIdx = appData.actual.length - 1;

        function switchTab(tabId) {{
            for (let i = 1; i <= 4; i++) {{
                document.getElementById('content-tab-' + i).classList.add('hidden');
                document.getElementById('btn-tab-' + i).classList.remove('active');
            }}
            document.getElementById('content-tab-' + tabId).classList.remove('hidden');
            document.getElementById('btn-tab-' + tabId).classList.add('active');
        }}

        // ĐỔI CHẾ ĐỘ BIỂU ĐỒ TAB 1 (4 MẢNG / CÔNG VS PHỤ TÙNG / ĐƯỜNG TỔNG THU)
        function changeChartMode(mode) {{
            currentChartMode = mode;
            document.getElementById('btnModePillars').classList.remove('active');
            document.getElementById('btnModeLaborParts').classList.remove('active');
            document.getElementById('btnModeLine').classList.remove('active');
            
            if (mode === 'pillars') document.getElementById('btnModePillars').classList.add('active');
            if (mode === 'labor-parts') document.getElementById('btnModeLaborParts').classList.add('active');
            if (mode === 'line') document.getElementById('btnModeLine').classList.add('active');

            renderMainChartTab1();
        }}

        function openModal() {{
            const nextIdx = appData.actual.length + 1;
            document.getElementById('inThang').value = 'Tháng ' + nextIdx;
            const last = appData.actual[appData.actual.length - 1];
            document.getElementById('copyBtnLabel').innerText = 'Sao chép số từ ' + last.name;
            calcFormTotals();
            document.getElementById('addModal').classList.remove('hidden');
        }}

        function closeModal() {{
            document.getElementById('addModal').classList.add('hidden');
        }}

        function copyFromLastMonth() {{
            const last = appData.actual[appData.actual.length - 1];
            if (!last) return;
            document.getElementById('inCongBD').value = last.congBD;
            document.getElementById('inCongSCC').value = last.congSCC;
            document.getElementById('inPtBDSCC').value = last.ptBDSCC;
            document.getElementById('inCongGo').value = last.congGo;
            document.getElementById('inCongSon').value = last.congSon;
            document.getElementById('inPtDS').value = last.ptDS;
            document.getElementById('inCongBH').value = last.congBH;
            document.getElementById('inPtBH').value = last.ptBH;
            document.getElementById('inCuuHo').value = last.cuuHo;
            calcFormTotals();
        }}

        function calcFormTotals() {{
            const cgo = parseFloat(document.getElementById('inCongGo').value) || 0;
            const cson = parseFloat(document.getElementById('inCongSon').value) || 0;
            const tongDS = cgo + cson;
            document.getElementById('formTotalDS').innerText = tongDS.toLocaleString('vi-VN') + ' đ';

            const cbd = parseFloat(document.getElementById('inCongBD').value) || 0;
            const cscc = parseFloat(document.getElementById('inCongSCC').value) || 0;
            const ptbdscc = parseFloat(document.getElementById('inPtBDSCC').value) || 0;
            const ptds = parseFloat(document.getElementById('inPtDS').value) || 0;
            const cbh = parseFloat(document.getElementById('inCongBH').value) || 0;
            const ptbh = parseFloat(document.getElementById('inPtBH').value) || 0;
            const cho = parseFloat(document.getElementById('inCuuHo').value) || 0;

            const grand = cbd + cscc + ptbdscc + tongDS + ptds + cbh + ptbh + cho;
            document.getElementById('formGrandTotal').innerText = grand.toLocaleString('vi-VN') + ' đ';
        }}

        function handleSaveMonth(e) {{
            e.preventDefault();
            const name = document.getElementById('inThang').value;
            const cbd = parseFloat(document.getElementById('inCongBD').value) || 0;
            const cscc = parseFloat(document.getElementById('inCongSCC').value) || 0;
            const ptbdscc = parseFloat(document.getElementById('inPtBDSCC').value) || 0;
            const cgo = parseFloat(document.getElementById('inCongGo').value) || 0;
            const cson = parseFloat(document.getElementById('inCongSon').value) || 0;
            const tongds = cgo + cson;
            const ptds = parseFloat(document.getElementById('inPtDS').value) || 0;
            const cbh = parseFloat(document.getElementById('inCongBH').value) || 0;
            const ptbh = parseFloat(document.getElementById('inPtBH').value) || 0;
            const cho = parseFloat(document.getElementById('inCuuHo').value) || 0;
            const total = cbd + cscc + ptbdscc + tongds + ptds + cbh + ptbh + cho;

            const newRec = {{
                id: 'm' + (appData.actual.length + 1),
                name: name,
                congBD: cbd, congSCC: cscc, ptBDSCC: ptbdscc,
                congGo: cgo, congSon: cson, tongCongDS: tongds, ptDS: ptds,
                congBH: cbh, ptBH: ptbh, tong: total - cho, cuuHo: cho, total: total
            }};

            appData.actual.push(newRec);
            closeModal();

            currentSelectedMonthIdx = appData.actual.length - 1;
            refreshAllUI();
        }}

        function refreshAllUI() {{
            const actual = appData.actual;
            const totalSum = actual.reduce((s, r) => s + r.total, 0);
            const totalTy = (totalSum / 1000000000).toFixed(2);
            const avgMonth = (totalSum / actual.length / 1000000000).toFixed(2);

            document.getElementById('topTotalActual').innerText = totalTy + ' tỷ';
            document.getElementById('badgeMonthCount').innerText = 'Số Liệu Thực Tế ' + actual.length + ' Tháng';
            document.getElementById('kpiYTD').innerText = totalTy;
            document.getElementById('kpiAvgMonth').innerText = 'TB: ' + avgMonth + ' tỷ / tháng';
            document.getElementById('kpiMonthCount').innerText = actual.length + ' tháng';

            let peak = actual[0];
            let peakIdx = 0;
            actual.forEach((r, idx) => {{
                if (r.total > peak.total) {{ peak = r; peakIdx = idx; }}
            }});
            document.getElementById('kpiPeakMonth').innerText = peak.name;
            document.getElementById('kpiPeakRev').innerText = (peak.total / 1000000000).toFixed(2) + ' tỷ';
            if (peakIdx > 0) {{
                const prev = actual[peakIdx - 1];
                const mom = ((peak.total - prev.total) / prev.total) * 100;
                document.getElementById('kpiPeakMoM').innerText = (mom > 0 ? '+' : '') + mom.toFixed(1) + '% MoM';
            }}

            const totalLabor = actual.reduce((s, r) => s + r.congBD + r.congSCC + r.tongCongDS + r.congBH, 0);
            const totalParts = actual.reduce((s, r) => s + r.ptBDSCC + r.ptDS + r.ptBH, 0);
            const ratio = (totalParts / (totalLabor || 1)).toFixed(2);
            document.getElementById('kpiRatio').innerText = ratio + 'x';
            document.getElementById('kpiPtRatio').innerText = 'PT: ' + ((totalParts / totalSum) * 100).toFixed(1) + '%';
            document.getElementById('kpiLaborRatio').innerText = 'Công: ' + ((totalLabor / totalSum) * 100).toFixed(1) + '%';

            const totalBH = actual.reduce((s, r) => s + r.congBH + r.ptBH, 0);
            document.getElementById('kpiBH').innerText = (totalBH / 1000000000).toFixed(2);
            document.getElementById('kpiBHRate').innerText = 'Chiếm ' + ((totalBH / totalSum) * 100).toFixed(1) + '% toàn xưởng';

            document.getElementById('chartTitleActual').innerText = 'Biến Động Doanh Thu Thực Tế (Tháng 1 Đến ' + actual[actual.length - 1].name + ')';
            renderMainChartTab1();
            renderDonutTab1();
            renderTab2MonthPills();
            selectMonth(currentSelectedMonthIdx);
            renderTab4Excel13Cols();
        }}

        // VẼ BIỂU ĐỒ TAB 1 THEO 3 CHẾ ĐỘ
        function renderMainChartTab1() {{
            const actual = appData.actual;
            const labels = actual.map(d => d.name);
            const ctx = document.getElementById('mainChartTab1').getContext('2d');
            if (mainChartInstance) mainChartInstance.destroy();

            const legendBox = document.getElementById('legendModePillars');

            if (currentChartMode === 'pillars') {{
                legendBox.classList.remove('hidden');
                legendBox.innerHTML = `
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-blue-600"></span><span>BD & SCC</span></div>
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-amber-500"></span><span>Đồng Sơn</span></div>
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-emerald-500"></span><span>Bảo Hành Chính Hãng</span></div>
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-purple-600"></span><span>Cứu Hộ Giao Thông</span></div>
                `;
                mainChartInstance = new Chart(ctx, {{
                    type: 'bar',
                    data: {{
                        labels: labels,
                        datasets: [
                            {{ label: 'BD & SCC', data: actual.map(d => (d.congBD + d.congSCC + d.ptBDSCC) / 1000000000), backgroundColor: '#2563eb' }},
                            {{ label: 'Đồng Sơn', data: actual.map(d => (d.tongCongDS + d.ptDS) / 1000000000), backgroundColor: '#f59e0b' }},
                            {{ label: 'Bảo Hành', data: actual.map(d => (d.congBH + d.ptBH) / 1000000000), backgroundColor: '#10b981' }},
                            {{ label: 'Cứu Hộ', data: actual.map(d => d.cuuHo / 1000000000), backgroundColor: '#8b5cf6' }}
                        ]
                    }},
                    options: {{
                        responsive: true, maintainAspectRatio: false,
                        scales: {{ x: {{ stacked: true, grid: {{ display: false }} }}, y: {{ stacked: true, grid: {{ color: '#f1f5f9' }} }} }},
                        plugins: {{ legend: {{ display: false }} }}
                    }}
                }});
            }} else if (currentChartMode === 'labor-parts') {{
                legendBox.classList.remove('hidden');
                legendBox.innerHTML = `
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-[#3b82f6]"></span><span>Tiền Công (Labor)</span></div>
                    <div class="flex items-center space-x-1.5"><span class="w-3 h-3 rounded bg-[#ec4899]"></span><span>Phụ Tùng (Parts)</span></div>
                `;
                mainChartInstance = new Chart(ctx, {{
                    type: 'bar',
                    data: {{
                        labels: labels,
                        datasets: [
                            {{ label: 'Tiền Công', data: actual.map(d => (d.congBD + d.congSCC + d.tongCongDS + d.congBH) / 1000000000), backgroundColor: '#3b82f6' }},
                            {{ label: 'Phụ Tùng', data: actual.map(d => (d.ptBDSCC + d.ptDS + d.ptBH) / 1000000000), backgroundColor: '#ec4899' }}
                        ]
                    }},
                    options: {{
                        responsive: true, maintainAspectRatio: false,
                        scales: {{ x: {{ stacked: false, grid: {{ display: false }} }}, y: {{ grid: {{ color: '#f1f5f9' }} }} }},
                        plugins: {{ legend: {{ display: false }} }}
                    }}
                }});
            }} else {{
                // ĐƯỜNG TỔNG THU (ẢNH 1)
                legendBox.classList.remove('hidden');
                legendBox.innerHTML = `
                    <div class="flex items-center space-x-1.5"><span class="w-3.5 h-1.5 rounded bg-blue-600"></span><span class="text-blue-600 font-bold">Tổng Doanh Thu</span></div>
                `;
                mainChartInstance = new Chart(ctx, {{
                    type: 'line',
                    data: {{
                        labels: labels,
                        datasets: [{{
                            label: 'Tổng Doanh Thu',
                            data: actual.map(d => d.total / 1000000000),
                            borderColor: '#2563eb',
                            backgroundColor: '#2563eb',
                            borderWidth: 3,
                            tension: 0.35,
                            pointRadius: 5,
                            pointBackgroundColor: '#2563eb'
                        }}]
                    }},
                    options: {{
                        responsive: true, maintainAspectRatio: false,
                        scales: {{ x: {{ grid: {{ display: false }} }}, y: {{ grid: {{ color: '#f1f5f9' }} }} }},
                        plugins: {{ legend: {{ display: false }} }}
                    }}
                }});
            }}
        }}

        // DONUT LŨY KẾ TAB 1
        function renderDonutTab1() {{
            const actual = appData.actual;
            const totBDSCC = actual.reduce((s, r) => s + (r.congBD + r.congSCC + r.ptBDSCC), 0);
            const totDS = actual.reduce((s, r) => s + (r.tongCongDS + r.ptDS), 0);
            const totBH = actual.reduce((s, r) => s + (r.congBH + r.ptBH), 0);
            const totCH = actual.reduce((s, r) => s + r.cuuHo, 0);
            const grandTotal = totBDSCC + totDS + totBH + totCH;

            document.getElementById('donutCenterTotal').innerText = (grandTotal / 1000000000).toFixed(2) + ' tỷ';
            document.getElementById('donutLegend').innerHTML = `
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span><span>Bảo Dưỡng & SCC</span></div><span class="font-bold font-mono">${{((totBDSCC/grandTotal)*100).toFixed(1)}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span><span>Đồng Sơn (Gò + Sơn)</span></div><span class="font-bold font-mono">${{((totDS/grandTotal)*100).toFixed(1)}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span><span>Bảo Hành Chính Hãng</span></div><span class="font-bold font-mono">${{((totBH/grandTotal)*100).toFixed(1)}}%</span></div>
                <div class="flex justify-between items-center"><div class="flex items-center space-x-2"><span class="w-2.5 h-2.5 rounded-full bg-purple-600"></span><span>Cứu Hộ Giao Thông</span></div><span class="font-bold font-mono">${{((totCH/grandTotal)*100).toFixed(1)}}%</span></div>
            `;

            const ctxDonut = document.getElementById('donutChart').getContext('2d');
            if (donutChartInstance) donutChartInstance.destroy();
            donutChartInstance = new Chart(ctxDonut, {{
                type: 'doughnut',
                data: {{
                    labels: ['Bảo Dưỡng & SCC', 'Đồng Sơn', 'Bảo Hành OEM', 'Cứu Hộ'],
                    datasets: [{{
                        data: [totBDSCC, totDS, totBH, totCH],
                        backgroundColor: ['#2563eb', '#f59e0b', '#10b981', '#8b5cf6'],
                        cutout: '72%'
                    }}]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
            }});
        }}

        // TAB 2: CHỌN THÁNG
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
            
            let maxTotal = Math.max(...appData.actual.map(x => x.total));
            document.getElementById('monthBadge').style.display = (m.total >= maxTotal) ? 'inline-block' : 'none';
            document.getElementById('monthTotalText').innerText = m.total.toLocaleString('vi-VN') + ' đ';
            
            document.getElementById('mCardTotal').innerText = (m.total / 1000000000).toFixed(2) + ' tỷ';
            document.getElementById('mCardSub').innerText = m.total.toLocaleString('vi-VN') + ' đ';
            
            if (idx > 0) {{
                const prev = appData.actual[idx - 1];
                const mom = ((m.total - prev.total) / prev.total) * 100;
                document.getElementById('mCardMoM').innerText = (mom > 0 ? '+' : '') + mom.toFixed(1) + '%';
                document.getElementById('mCardMoM').className = 'text-lg font-bold mt-1 ' + (mom >= 0 ? 'text-emerald-600' : 'text-red-500');
                document.getElementById('mCardPrev').innerText = 'Tháng trước: ' + (prev.total / 1000000000).toFixed(2) + ' tỷ';
            }} else {{
                document.getElementById('mCardMoM').innerText = '—';
                document.getElementById('mCardPrev').innerText = 'Kỳ báo cáo đầu';
            }}

            const mLabor = m.congBD + m.congSCC + m.tongCongDS + m.congBH;
            const mParts = m.ptBDSCC + m.ptDS + m.ptBH;
            const ratio = (mParts / (mLabor || 1)).toFixed(2);
            document.getElementById('mCardRatio').innerText = ratio + 'x';
            document.getElementById('mCardRatioSub').innerText = 'Công: ' + (mLabor/1000000).toFixed(0) + ' tr | PT: ' + (mParts/1000000000).toFixed(2) + ' tỷ';
            document.getElementById('mCardCH').innerText = (m.cuuHo / 1000000).toFixed(1) + ' tr';
            document.getElementById('mCardCHSub').innerText = 'Chiếm ' + ((m.cuuHo / m.total) * 100).toFixed(1) + '%';

            updateMonthChartsAndRankings(m, mLabor, mParts);
        }}

        // VẼ BIỂU ĐỒ CỘT ĐÔI + DONUT + BẢNG XẾP HẠNG 9 HẠNG MỤC CỦA THÁNG
        function updateMonthChartsAndRankings(m, mLabor, mParts) {{
            // 1. Biểu đồ 4 trụ cột có cột đôi (Tiền Công + PT đứng cạnh Tổng Cả Mảng - Ảnh 3)
            const valBDSCC = m.congBD + m.congSCC + m.ptBDSCC;
            const valDS = m.tongCongDS + m.ptDS;
            const valBH = m.congBH + m.ptBH;
            const valCH = m.cuuHo;

            const ctxDCol = document.getElementById('monthDoubleColChart').getContext('2d');
            if (monthDoubleColInstance) monthDoubleColInstance.destroy();
            monthDoubleColInstance = new Chart(ctxDCol, {{
                type: 'bar',
                data: {{
                    labels: ['Bảo Dưỡng & SCC', 'Tổ Đồng Sơn', 'Bảo Hành (OEM)', 'Cứu Hộ Giao Thông'],
                    datasets: [
                        {{ label: 'Tiền Công', data: [m.congBD + m.congSCC, m.tongCongDS, m.congBH, 0], backgroundColor: '#3b82f6', stack: 'stack1' }},
                        {{ label: 'Phụ Tùng', data: [m.ptBDSCC, m.ptDS, m.ptBH, 0], backgroundColor: '#ec4899', stack: 'stack1' }},
                        {{ label: 'Tổng Cả Mảng', data: [valBDSCC, valDS, valBH, valCH], backgroundColor: '#94a3b8', stack: 'stack2' }}
                    ]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    scales: {{
                        x: {{ grid: {{ display: false }} }},
                        y: {{ ticks: {{ callback: v => (v/1000000) + 'tr' }}, grid: {{ color: '#f1f5f9' }} }}
                    }},
                    plugins: {{ legend: {{ display: false }} }}
                }}
            }});

            // 2. Donut của tháng kèm đầy đủ số tiền và % (Ảnh 5)
            document.getElementById('donutMonthCenterName').innerText = m.name;
            document.getElementById('donutMonthCenterTotal').innerText = (m.total / 1000000000).toFixed(2) + ' tỷ';

            const ctxMPie = document.getElementById('monthDonutChart').getContext('2d');
            if (monthDonutInstance) monthDonutInstance.destroy();
            monthDonutInstance = new Chart(ctxMPie, {{
                type: 'doughnut',
                data: {{
                    labels: ['Bảo Dưỡng & SCC', 'Tổ Đồng Sơn', 'Bảo Hành (OEM)', 'Cứu Hộ Giao Thông'],
                    datasets: [{{
                        data: [valBDSCC, valDS, valBH, valCH],
                        backgroundColor: ['#2563eb', '#f59e0b', '#10b981', '#8b5cf6'],
                        cutout: '65%'
                    }}]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
            }});

            const fmtCompact = v => v >= 1000000000 ? (v/1000000000).toFixed(2) + ' tỷ' : (v/1000000).toFixed(1) + ' tr';
            document.getElementById('monthDonutList').innerHTML = `
                <div class="flex items-center justify-between text-[11px]"><div class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-blue-600"></span><span class="text-slate-700">Bảo Dưỡng & SCC</span></div><div class="flex items-center gap-2 font-mono"><span class="text-slate-500">${{fmtCompact(valBDSCC)}}</span><span class="font-bold text-slate-900 w-12 text-right">${{((valBDSCC/m.total)*100).toFixed(1)}}%</span></div></div>
                <div class="flex items-center justify-between text-[11px]"><div class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-amber-500"></span><span class="text-slate-700">Tổ Đồng Sơn</span></div><div class="flex items-center gap-2 font-mono"><span class="text-slate-500">${{fmtCompact(valDS)}}</span><span class="font-bold text-slate-900 w-12 text-right">${{((valDS/m.total)*100).toFixed(1)}}%</span></div></div>
                <div class="flex items-center justify-between text-[11px]"><div class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-emerald-500"></span><span class="text-slate-700">Bảo Hành (OEM)</span></div><div class="flex items-center gap-2 font-mono"><span class="text-slate-500">${{fmtCompact(valBH)}}</span><span class="font-bold text-slate-900 w-12 text-right">${{((valBH/m.total)*100).toFixed(1)}}%</span></div></div>
                <div class="flex items-center justify-between text-[11px]"><div class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-purple-600"></span><span class="text-slate-700">Cứu Hộ Giao Thông</span></div><div class="flex items-center gap-2 font-mono"><span class="text-slate-500">${{fmtCompact(valCH)}}</span><span class="font-bold text-slate-900 w-12 text-right">${{((valCH/m.total)*100).toFixed(1)}}%</span></div></div>
            `;

            // 3. BẢNG XẾP HẠNG 9 HẠNG MỤC TỪ CAO XUỐNG THẤP (ẢNH 4 CỦA USER)
            document.getElementById('rankTitle').innerText = 'Bảng Xếp Hạng Hạng Mục Thực Tế Trong ' + m.name + ' (Từ Cao Xuống Thấp)';
            const items = [
                {{ name: 'PT Bảo hành', value: m.ptBH, group: 'Phụ tùng' }},
                {{ name: 'PT Đồng Sơn', value: m.ptDS, group: 'Phụ tùng' }},
                {{ name: 'PT BD+SCC', value: m.ptBDSCC, group: 'Phụ tùng' }},
                {{ name: 'Công Bảo hành', value: m.congBH, group: 'Công thợ' }},
                {{ name: 'Công SCC', value: m.congSCC, group: 'Công thợ' }},
                {{ name: 'Công Sơn', value: m.congSon, group: 'Công thợ' }},
                {{ name: 'Công Bảo dưỡng', value: m.congBD, group: 'Công thợ' }},
                {{ name: 'Công Gò', value: m.congGo, group: 'Công thợ' }},
                {{ name: 'Cứu hộ', value: m.cuuHo, group: 'Khác' }}
            ].sort((a, b) => b.value - a.value);

            const rankContainer = document.getElementById('rankGridContainer');
            rankContainer.innerHTML = '';
            items.forEach((it, idx) => {{
                const pct = ((it.value / m.total) * 100).toFixed(1);
                const el = document.createElement('div');
                el.className = 'p-3 rounded-xl border border-slate-100 bg-slate-50/70 flex items-center justify-between text-xs';
                el.innerHTML = `
                    <div class="flex items-center gap-2.5">
                        <span class="w-5 h-5 rounded-full bg-slate-200 text-slate-700 font-bold text-[10px] flex items-center justify-center">${{idx + 1}}</span>
                        <div>
                            <span class="font-bold text-slate-800 block">${{it.name}}</span>
                            <span class="text-[10px] text-slate-400">${{it.group}}</span>
                        </div>
                    </div>
                    <div class="text-right font-mono">
                        <span class="font-bold text-slate-900 block">${{(it.value/1000000).toFixed(1)}} tr</span>
                        <span class="text-[10px] text-slate-500">${{pct}}%</span>
                    </div>
                `;
                rankContainer.appendChild(el);
            }});
        }}

        // TAB 4: BẢNG TÍNH GỐC EXCEL 13 CỘT ĐẦY ĐỦ TỪNG ĐỒNG (ẢNH 7 CỦA USER)
        function renderTab4Excel13Cols() {{
            const tbody = document.getElementById('table13ColsBody');
            tbody.innerHTML = '';
            const actual = appData.actual;

            actual.forEach(r => {{
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 divide-x divide-slate-200';
                tr.innerHTML = `
                    <td class="py-2.5 px-3 font-bold bg-slate-50 sticky left-0 z-10">${{r.name}}</td>
                    <td class="py-2 px-2 text-right bg-emerald-50/50">${{r.congBD.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.congSCC.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.ptBDSCC.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.congGo.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.congSon.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-blue-50/70 font-semibold text-blue-900">${{r.tongCongDS.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.ptDS.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.congBH.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.ptBH.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-slate-100 font-bold">${{r.tong.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{r.cuuHo.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-amber-100 text-amber-950 font-black">${{r.total.toLocaleString('vi-VN')}} đ</td>
                `;
                tbody.appendChild(tr);
            }});

            // Dòng tổng cộng 13 cột chuẩn xác
            const totCongBD = actual.reduce((s, r) => s + r.congBD, 0);
            const totCongSCC = actual.reduce((s, r) => s + r.congSCC, 0);
            const totPtBDSCC = actual.reduce((s, r) => s + r.ptBDSCC, 0);
            const totCongGo = actual.reduce((s, r) => s + r.congGo, 0);
            const totCongSon = actual.reduce((s, r) => s + r.congSon, 0);
            const totTongCongDS = actual.reduce((s, r) => s + r.tongCongDS, 0);
            const totPtDS = actual.reduce((s, r) => s + r.ptDS, 0);
            const totCongBH = actual.reduce((s, r) => s + r.congBH, 0);
            const totPtBH = actual.reduce((s, r) => s + r.ptBH, 0);
            const totTong = actual.reduce((s, r) => s + r.tong, 0);
            const totCuuHo = actual.reduce((s, r) => s + r.cuuHo, 0);
            const totGrand = actual.reduce((s, r) => s + r.total, 0);

            document.getElementById('table13ColsFoot').innerHTML = `
                <tr class="bg-slate-200 border-t-2 border-slate-400 divide-x divide-slate-300">
                    <td class="py-2.5 px-3 bg-slate-300 font-black sticky left-0 z-10">Tổng Thực Tế</td>
                    <td class="py-2 px-2 text-right text-emerald-950 font-bold">${{totCongBD.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totCongSCC.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totPtBDSCC.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totCongGo.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totCongSon.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-blue-100 text-blue-950 font-bold">${{totTongCongDS.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totPtDS.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totCongBH.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totPtBH.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-slate-300 font-black">${{totTong.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right">${{totCuuHo.toLocaleString('vi-VN')}} đ</td>
                    <td class="py-2 px-2 text-right bg-amber-200 text-amber-950 font-black">${{totGrand.toLocaleString('vi-VN')}} đ</td>
                </tr>
            `;
        }}

        function exportToCSV() {{
            let csv = "Thang,CongBaoDuong,CongSCC,PhuTungBDSCC,CongGo,CongSon,TongCongDS,PtDongSon,CongBaoHanh,PtBaoHanh,Tong,CuuHo,TongCongCuuHo\\n";
            appData.actual.forEach(r => {{
                csv += `${{r.name}},${{r.congBD}},${{r.congSCC}},${{r.ptBDSCC}},${{r.congGo}},${{r.congSon}},${{r.tongCongDS}},${{r.ptDS}},${{r.congBH}},${{r.ptBH}},${{r.tong}},${{r.cuuHo}},${{r.total}}\\n`;
            }});
            const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "Data_doanhthu_VinFast_Chi_Tiet.csv";
            link.click();
        }}

        refreshAllUI();
    </script>
</body>
</html>
"""

components.html(HTML_CONTENT, height=980, scrolling=True)
