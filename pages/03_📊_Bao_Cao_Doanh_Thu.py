import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Cấu hình giao diện Streamlit rộng rãi
st.set_page_config(
    page_title="Báo Cáo Doanh Thu Xưởng Dịch Vụ",
    page_icon="📊",
    layout="wide"
)

# Đường dẫn file lưu trữ số liệu
DATA_CSV_PATH = "data_doanh_thu_thang.csv"

# Hàm định dạng tiền tệ VNĐ chuẩn
def format_vnd_compact(val):
    if abs(val) >= 1_000_000_000:
        return f"{val / 1_000_000_000:.2f} tỷ"
    elif abs(val) >= 1_000_000:
        return f"{val / 1_000_000:.1f} tr"
    return f"{val:,.0f} đ"

def format_vnd_full(val):
    return f"{val:,.0f} đ".replace(",", ".")

# Tải dữ liệu 7 tháng chuẩn từ bảng Excel
@st.cache_data
def get_initial_data():
    if os.path.exists(DATA_CSV_PATH):
        try:
            return pd.read_csv(DATA_CSV_PATH)
        except Exception:
            pass
            
    raw = [
        {"Tháng": "Tháng 1", "Công Bảo dưỡng": 105552349, "Công SCC": 132185890, "Phụ tùng BD+SCC": 392433779, "Công Gò": 33274750, "Công Sơn": 67965091, "Tổng Công ĐS": 101239841, "PT Đồng Sơn": 420983471, "Công Bảo hành": 177852500, "PT Bảo hành": 633141964, "Tổng": 1963389794, "Cứu hộ": 43466800, "Tổng ( Cộng Cứu hộ)": 2006856594},
        {"Tháng": "Tháng 2", "Công Bảo dưỡng": 61107500, "Công SCC": 95955155, "Phụ tùng BD+SCC": 204226871, "Công Gò": 23450000, "Công Sơn": 47083611, "Tổng Công ĐS": 70533611, "PT Đồng Sơn": 261003264, "Công Bảo hành": 126725000, "PT Bảo hành": 357115264, "Tổng": 1169999998, "Cứu hộ": 6292785, "Tổng ( Cộng Cứu hộ)": 1176292783},
        {"Tháng": "Tháng 3", "Công Bảo dưỡng": 96957944, "Công SCC": 163173190, "Phụ tùng BD+SCC": 771805214, "Công Gò": 80376666, "Công Sơn": 119426804, "Tổng Công ĐS": 199803470, "PT Đồng Sơn": 486322661, "Công Bảo hành": 56215000, "PT Bảo hành": 288563726, "Tổng": 2049507871, "Cứu hộ": 11399995, "Tổng ( Cộng Cứu hộ)": 2060907866},
        {"Tháng": "Tháng 4", "Công Bảo dưỡng": 136176250, "Công SCC": 109065486, "Phụ tùng BD+SCC": 511523675, "Công Gò": 55667195, "Công Sơn": 128411661, "Tổng Công ĐS": 184078856, "PT Đồng Sơn": 607035927, "Công Bảo hành": 126402500, "PT Bảo hành": 499518732, "Tổng": 2173801426, "Cứu hộ": 15040600, "Tổng ( Cộng Cứu hộ)": 2188842026},
        {"Tháng": "Tháng 5", "Công Bảo dưỡng": 85586250, "Công SCC": 54650487, "Phụ tùng BD+SCC": 179987803, "Công Gò": 68734375, "Công Sơn": 152589500, "Tổng Công ĐS": 221323875, "PT Đồng Sơn": 444427264, "Công Bảo hành": 170097500, "PT Bảo hành": 486790109, "Tổng": 1642863288, "Cứu hộ": 9252000, "Tổng ( Cộng Cứu hộ)": 1652115288},
        {"Tháng": "Tháng 6", "Công Bảo dưỡng": 112101923, "Công SCC": 49094157, "Phụ tùng BD+SCC": 213555761, "Công Gò": 82175519, "Công Sơn": 136316033, "Tổng Công ĐS": 218491552, "PT Đồng Sơn": 465689781, "Công Bảo hành": 152840000, "PT Bảo hành": 759515486, "Tổng": 1971288660, "Cứu hộ": 63179000, "Tổng ( Cộng Cứu hộ)": 2034467660},
        {"Tháng": "Tháng 7", "Công Bảo dưỡng": 103387506, "Công SCC": 223356704, "Phụ tùng BD+SCC": 411113918, "Công Gò": 96945250, "Công Sơn": 153434000, "Tổng Công ĐS": 250379250, "PT Đồng Sơn": 488189498, "Công Bảo hành": 224862500, "PT Bảo hành": 987791454, "Tổng": 2689080830, "Cứu hộ": 38256000, "Tổng ( Cộng Cứu hộ)": 2727336830},
    ]
    df_init = pd.DataFrame(raw)
    df_init.to_csv(DATA_CSV_PATH, index=False)
    return df_init

df = get_initial_data()

# Header giao diện
st.markdown("""
<div style="background: linear-gradient(90deg, #1e293b, #0f172a); padding: 18px 22px; border-radius: 12px; margin-bottom: 20px; color: white; display: flex; justify-content: space-between; align-items: center;">
    <div>
        <h2 style="margin:0; font-size: 20px; color: white;">📊 Báo Cáo Phân Tích Doanh Thu Xưởng Dịch Vụ</h2>
        <p style="margin: 4px 0 0 0; font-size: 13px; color: #94a3b8;">
            Bảo dưỡng định kỳ, Sửa chữa chung, Đồng Sơn, Bảo hành & Cứu hộ giao thông
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# Tính toán các chỉ số YTD
ytd_revenue = df["Tổng ( Cộng Cứu hộ)"].sum()
avg_monthly_rev = ytd_revenue / len(df)
peak_row = df.loc[df["Tổng ( Cộng Cứu hộ)"].idxmax()]
total_labor = (df["Công Bảo dưỡng"] + df["Công SCC"] + df["Tổng Công ĐS"] + df["Công Bảo hành"]).sum()
total_parts = (df["Phụ tùng BD+SCC"] + df["PT Đồng Sơn"] + df["PT Bảo hành"]).sum()
ratio_pt_cong = total_parts / total_labor if total_labor > 0 else 0
total_warranty = (df["Công Bảo hành"] + df["PT Bảo hành"]).sum()

# TÁCH RÕ 4 TAB CHUẨN UX
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 1. Số Liệu Thực Tế (7 Tháng)",
    "📅 2. Biểu Đồ Từng Tháng (Thực Tế)",
    "🔮 3. Kế Hoạch & Dự Báo (T8 - T12) [Tab Riêng]",
    "📋 4. Bảng Tính Gốc (Excel)"
])

# ================= TAB 1: SỐ LIỆU THỰC TẾ =================
with tab1:
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Doanh Thu Thực Tế (YTD)", format_vnd_compact(ytd_revenue), f"TB: {format_vnd_compact(avg_monthly_rev)}/tháng")
    with c2:
        st.metric("Tháng Đỉnh Doanh Thu", peak_row["Tháng"], format_vnd_compact(peak_row["Tổng ( Cộng Cứu hộ)"]))
    with c3:
        st.metric("Tỷ Lệ Phụ Tùng / Công", f"{ratio_pt_cong:.2f}x", f"PT: {(total_parts/ytd_revenue)*100:.1f}% | Công: {(total_labor/ytd_revenue)*100:.1f}%")
    with c4:
        st.metric("Bảo Hành Nhà Máy (W)", format_vnd_compact(total_warranty), f"Chiếm {(total_warranty/ytd_revenue)*100:.1f}% toàn xưởng")

    st.markdown("---")

    col_chart_left, col_chart_right = st.columns([2, 1])
    with col_chart_left:
        st.subheader("Biến Động Doanh Thu Thực Tế (Tháng 1 Đến Tháng 7)")
        chart_data = pd.DataFrame({
            "Tháng": df["Tháng"],
            "BD & SCC": df["Công Bảo dưỡng"] + df["Công SCC"] + df["Phụ tùng BD+SCC"],
            "Đồng Sơn": df["Tổng Công ĐS"] + df["PT Đồng Sơn"],
            "Bảo Hành Chính Hãng": df["Công Bảo hành"] + df["PT Bảo hành"],
            "Cứu Hộ Giao Thông": df["Cứu hộ"],
        })
        fig_bar = px.bar(
            chart_data,
            x="Tháng",
            y=["BD & SCC", "Đồng Sơn", "Bảo Hành Chính Hãng", "Cứu Hộ Giao Thông"],
            barmode="stack",
            color_discrete_map={
                "BD & SCC": "#2563eb",
                "Đồng Sơn": "#f59e0b",
                "Bảo Hành Chính Hãng": "#10b981",
                "Cứu Hộ Giao Thông": "#8b5cf6"
            }
        )
        fig_bar.update_layout(legend_title="Trụ Cột Dịch Vụ", margin=dict(l=0, r=0, t=10, b=0), height=360)
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_chart_right:
        st.subheader("Cơ Cấu Thực Tế 4 Mảng Lũy Kế")
        pie_df = pd.DataFrame({
            "Mảng": ["Bảo Dưỡng & SCC", "Đồng Sơn (Gò + Sơn)", "Bảo Hành Chính Hãng", "Cứu Hộ Giao Thông"],
            "Doanh Thu": [
                (df["Công Bảo dưỡng"] + df["Công SCC"] + df["Phụ tùng BD+SCC"]).sum(),
                (df["Tổng Công ĐS"] + df["PT Đồng Sơn"]).sum(),
                (df["Công Bảo hành"] + df["PT Bảo hành"]).sum(),
                df["Cứu hộ"].sum()
            ]
        })
        fig_pie = px.pie(
            pie_df,
            names="Mảng",
            values="Doanh Thu",
            hole=0.6,
            color="Mảng",
            color_discrete_map={
                "Bảo Dưỡng & SCC": "#2563eb",
                "Đồng Sơn (Gò + Sơn)": "#f59e0b",
                "Bảo Hành Chính Hãng": "#10b981",
                "Cứu Hộ Giao Thông": "#8b5cf6"
            }
        )
        fig_pie.update_layout(margin=dict(l=0, r=0, t=10, b=0), height=360)
        st.plotly_chart(fig_pie, use_container_width=True)


# ================= TAB 2: BIỂU ĐỒ TỪNG THÁNG =================
with tab2:
    selected_m = st.selectbox("👉 Chọn Tháng Muốn Soi Sâu:", df["Tháng"].tolist(), index=len(df)-1)
    m_row = df[df["Tháng"] == selected_m].iloc[0]

    cur_rev = m_row["Tổng ( Cộng Cứu hộ)"]
    cur_labor = m_row["Công Bảo dưỡng"] + m_row["Công SCC"] + m_row["Tổng Công ĐS"] + m_row["Công Bảo hành"]
    cur_parts = m_row["Phụ tùng BD+SCC"] + m_row["PT Đồng Sơn"] + m_row["PT Bảo hành"]

    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Doanh Thu Tháng", format_vnd_compact(cur_rev), format_vnd_full(cur_rev))
    col_m2.metric("Tiền Công Thợ", format_vnd_compact(cur_labor), f"{(cur_labor/cur_rev)*100:.1f}%")
    col_m3.metric("Phụ Tùng", format_vnd_compact(cur_parts), f"{(cur_parts/cur_rev)*100:.1f}%")

    c_m_chart, c_m_table = st.columns([1, 1])
    with c_m_chart:
        st.write(f"**Doanh Thu 4 Mảng Trong {selected_m}**")
        fig_m_bar = px.bar(
            x=["BD & SCC", "Tổ Đồng Sơn", "Bảo Hành (OEM)", "Cứu Hộ"],
            y=[
                m_row["Công Bảo dưỡng"] + m_row["Công SCC"] + m_row["Phụ tùng BD+SCC"],
                m_row["Tổng Công ĐS"] + m_row["PT Đồng Sơn"],
                m_row["Công Bảo hành"] + m_row["PT Bảo hành"],
                m_row["Cứu hộ"]
            ],
            color=["BD & SCC", "Tổ Đồng Sơn", "Bảo Hành (OEM)", "Cứu Hộ"],
            color_discrete_map={
                "BD & SCC": "#2563eb",
                "Tổ Đồng Sơn": "#f59e0b",
                "Bảo Hành (OEM)": "#10b981",
                "Cứu Hộ": "#8b5cf6"
            }
        )
        fig_m_bar.update_layout(showlegend=False, yaxis_title="VNĐ", height=320)
        st.plotly_chart(fig_m_bar, use_container_width=True)

    with c_m_table:
        st.write(f"**Bảng Xếp Hạng Hạng Mục Thực Tế {selected_m}**")
        details = {
            "Công Bảo dưỡng": m_row["Công Bảo dưỡng"],
            "Công SCC": m_row["Công SCC"],
            "PT BD+SCC": m_row["Phụ tùng BD+SCC"],
            "Công Gò": m_row["Công Gò"],
            "Công Sơn": m_row["Công Sơn"],
            "PT Đồng Sơn": m_row["PT Đồng Sơn"],
            "Công Bảo hành": m_row["Công Bảo hành"],
            "PT Bảo hành": m_row["PT Bảo hành"],
            "Cứu hộ": m_row["Cứu hộ"]
        }
        sorted_details = sorted(details.items(), key=lambda x: x[1], reverse=True)
        r_df = pd.DataFrame(sorted_details, columns=["Hạng mục", "Doanh thu"])
        r_df["Doanh thu hiển thị"] = r_df["Doanh thu"].apply(format_vnd_full)
        r_df["Tỷ trọng (%)"] = ((r_df["Doanh thu"] / cur_rev) * 100).round(1).astype(str) + "%"
        st.dataframe(r_df[["Hạng mục", "Doanh thu hiển thị", "Tỷ trọng (%)"]], use_container_width=True, hide_index=True)


# ================= TAB 3: KẾ HOẠCH & DỰ BÁO (RIÊNG BIỆT) =================
with tab3:
    st.info("🔮 **Khu vực Kế hoạch độc lập**: Theo dõi tiến độ về đích cả năm mà không làm sai lệch số kế toán.")
    
    annual_target = st.selectbox(
        "Đặt Mục Tiêu Doanh Thu Cả Năm (VND):",
        [20_000_000_000, 22_000_000_000, 24_000_000_000, 26_000_000_000, 28_000_000_000],
        index=2
    )

    remaining_target = max(0, annual_target - ytd_revenue)
    remaining_months = max(1, 12 - len(df))
    monthly_runrate = remaining_target / remaining_months

    fc1, fc2, fc3 = st.columns(3)
    fc1.metric("Thực Tế Đã Đạt (YTD)", format_vnd_compact(ytd_revenue), f"{(ytd_revenue/annual_target)*100:.1f}% Mục tiêu")
    fc2.metric(f"Mục Tiêu Còn Lại ({remaining_months} tháng)", format_vnd_compact(remaining_target))
    fc3.metric("Mỗi Tháng Cần Đạt", format_vnd_compact(monthly_runrate), "Bình quân về đích")

    # Kịch bản dự kiến 5 tháng cuối năm
    fc_plan = pd.DataFrame({
        "Tháng": ["Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12"],
        "Kế Hoạch": [2_130_000_000, 2_265_000_000, 2_435_000_000, 2_610_000_000, 2_955_000_000]
    })

    full_line = pd.DataFrame({
        "Tháng": df["Tháng"].tolist() + fc_plan["Tháng"].tolist(),
        "Loại": ["Thực Tế (Đã chốt)"] * len(df) + ["Kế Hoạch Dự Kiến"] * len(fc_plan),
        "Doanh Thu": df["Tổng ( Cộng Cứu hộ)"].tolist() + fc_plan["Kế Hoạch"].tolist()
    })

    fig_full = px.bar(
        full_line,
        x="Tháng",
        y="Doanh Thu",
        color="Loại",
        color_discrete_map={"Thực Tế (Đã chốt)": "#2563eb", "Kế Hoạch Dự Kiến": "#a855f7"},
        title="Biểu đồ 12 Tháng: Thực Tế (Xanh) vs Kế Hoạch Dự Kiến (Tím)"
    )
    fig_full.add_hline(y=annual_target/12, line_dash="dash", line_color="#f59e0b", annotation_text="Định mức bình quân tháng")
    st.plotly_chart(fig_full, use_container_width=True)


# ================= TAB 4: BẢNG TÍNH GỐC & CẬP NHẬT THÁNG MỚI =================
with tab4:
    st.subheader("📋 Bảng Tính Doanh Thu Thực Tế")
    disp_df = df.copy()
    for col in [c for c in df.columns if c != "Tháng"]:
        disp_df[col] = disp_df[col].apply(format_vnd_full)
    st.dataframe(disp_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("➕ Cập Nhật Thêm Số Liệu Tháng Mới (Qua mỗi tháng)")
    st.caption("Các cột tổng cộng sẽ được tự động tính toán chính xác 100% theo công thức xưởng.")

    with st.form("form_add_month"):
        f_thang = st.text_input("Tên Tháng Mới:", value=f"Tháng {len(df)+1}")
        col_a, col_b, col_c = st.columns(3)

        with col_a:
            st.markdown("**1. Bảo Dưỡng & SCC**")
            in_cong_bd = st.number_input("Công Bảo dưỡng", min_value=0, value=110000000, step=1000000)
            in_cong_scc = st.number_input("Công SCC", min_value=0, value=150000000, step=1000000)
            in_pt_bdscc = st.number_input("Phụ tùng BD+SCC", min_value=0, value=450000000, step=1000000)

        with col_b:
            st.markdown("**2. Tổ Đồng Sơn**")
            in_cong_go = st.number_input("Công Gò", min_value=0, value=85000000, step=1000000)
            in_cong_son = st.number_input("Công Sơn", min_value=0, value=140000000, step=1000000)
            in_pt_ds = st.number_input("PT Đồng Sơn", min_value=0, value=480000000, step=1000000)

        with col_c:
            st.markdown("**3. Bảo Hành & Cứu Hộ**")
            in_cong_bh = st.number_input("Công Bảo hành", min_value=0, value=160000000, step=1000000)
            in_pt_bh = st.number_input("PT Bảo hành", min_value=0, value=520000000, step=1000000)
            in_cuu_ho = st.number_input("Cứu hộ", min_value=0, value=25000000, step=1000000)

        btn_submit = st.form_submit_button("💾 Lưu Số Liệu Tháng Này")
        if btn_submit:
            calc_tong_ds = in_cong_go + in_cong_son
            calc_tong = in_cong_bd + in_cong_scc + in_pt_bdscc + calc_tong_ds + in_pt_ds + in_cong_bh + in_pt_bh
            calc_tong_cuu_ho = calc_tong + in_cuu_ho

            new_record = {
                "Tháng": f_thang,
                "Công Bảo dưỡng": in_cong_bd,
                "Công SCC": in_cong_scc,
                "Phụ tùng BD+SCC": in_pt_bdscc,
                "Công Gò": in_cong_go,
                "Công Sơn": in_cong_son,
                "Tổng Công ĐS": calc_tong_ds,
                "PT Đồng Sơn": in_pt_ds,
                "Công Bảo hành": in_cong_bh,
                "PT Bảo hành": in_pt_bh,
                "Tổng": calc_tong,
                "Cứu hộ": in_cuu_ho,
                "Tổng ( Cộng Cứu hộ)": calc_tong_cuu_ho
            }

            new_df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
            new_df.to_csv(DATA_CSV_PATH, index=False)
            st.cache_data.clear()
            st.success(f" Đã lưu thành công dữ liệu cho {f_thang}!")
            st.rerun()
