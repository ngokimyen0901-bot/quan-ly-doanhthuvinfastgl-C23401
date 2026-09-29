# TAB 6: GỒM CẢ BỎ QUA THỦ CÔNG LẪN CÁC MỤC LẺ XƯỞNG TỰ CHỊU / HÃNG TỪ CHỐI
    with t_b_ignore:
        st.subheader("⚪ Danh Sách Hạng Mục Bỏ Qua & Hãng Không Duyệt (Xưởng Tự Chịu)")
        
        # Danh sách hiển thị có đầy đủ mã LSC gốc, biển số, cố vấn
        df_lsc_map = df_bh_m[[col_b_ct_lsc, 'lsc_norm', 'Biển số', 'Cố vấn dịch vụ']].drop_duplicates(subset=['lsc_norm']).copy()
        df_lsc_map['display_opt'] = df_lsc_map.apply(
            lambda r: f"{r[col_b_ct_lsc]} | {r['Biển số']} ({r['Cố vấn dịch vụ']})", axis=1
        )
        map_display_to_norm = dict(zip(df_lsc_map['display_opt'], df_lsc_map['lsc_norm']))
        map_norm_to_display = dict(zip(df_lsc_map['lsc_norm'], df_lsc_map['display_opt']))
        
        # Các lựa chọn mặc định đã chọn trước đó
        current_defaults = [map_norm_to_display[k] for k in st.session_state.ds_bo_qua if k in map_norm_to_display]

        c_big1, c_big2 = st.columns([7.5, 2.5])
        with c_big1:
            sel_b_display = st.multiselect(
                "Tìm & chọn các LSC muốn gán nhãn [BỎ QUA] (Gõ mã LSC, biển số hoặc tên CVDV):",
                options=df_lsc_map['display_opt'].tolist(),
                default=current_defaults,
                key="ms_b_ignore_display",
                placeholder="Gõ tìm kiếm VD: C23401-WO-260922-0011 hoặc 81A52924..."
            )
        with c_big2:
            st.write("")
            st.write("")
            if st.button("💾 Lưu Trạng Thái Bỏ Qua", type="primary", key="btn_b_ignore", use_container_width=True):
                # Lưu danh sách vào session_state
                chosen_norms = {map_display_to_norm[opt] for opt in sel_b_display if opt in map_display_to_norm}
                st.session_state.ds_bo_qua = chosen_norms
                
                # Cập nhật trực tiếp nhãn sang Google Sheets để lưu vĩnh viễn
                with st.spinner("⏳ Đang đồng bộ trạng thái Bỏ Qua lên Google Sheets..."):
                    for idx_r, r_val in df_bh_m.iterrows():
                        if r_val.get('lsc_norm') in chosen_norms:
                            df_bh_m.at[idx_r, 'Nhãn Trạng Thái'] = 'BỎ QUA (Không theo dõi)'
                    try:
                        conn.update(worksheet="ChiTiet_BaoHanh", data=df_bh_m)
                        st.toast("✅ Đã lưu trạng thái Bỏ Qua vĩnh viễn lên Google Sheets!", icon="🚀")
                    except Exception as e_save_ig:
                        st.warning(f"Lưu tạm bộ nhớ (chưa ghi Sheets): {e_save_ig}")
                st.success("Đã cập nhật danh sách bỏ qua thành công!")
                st.rerun()

        # Hiển thị tất cả các mục bị bỏ qua / không claim / hãng từ chối
        df_b_ignored = df_bh_m[
            df_bh_m['lsc_norm'].isin(st.session_state.ds_bo_qua) | 
            df_bh_m['Nhãn Trạng Thái'].apply(is_bo_qua_hoac_khong_duyet)
        ].copy()

        if not df_b_ignored.empty:
            df_b_ig_sum = df_b_ignored.groupby(['lsc_norm', col_b_ct_lsc]).agg({
                'Biển số': 'first',
                'Cố vấn dịch vụ': 'first',
                col_b_ct_mavt: 'count',
                'Tiền xưởng (DMS)': 'sum',
                'Nhãn Trạng Thái': lambda x: " | ".join(sorted(set(x)))
            }).reset_index().rename(columns={
                col_b_ct_mavt: 'Số mục bỏ qua',
                'Tiền xưởng (DMS)': 'Tiền xưởng',
                'Nhãn Trạng Thái': 'Lý do / Trạng thái'
            })
            df_b_ig_sum.insert(0, 'STT', range(1, len(df_b_ig_sum) + 1))
            df_b_ig_sum['Ghi chú kế toán'] = "Xưởng tự chịu chi phí (Quá hạn / Không đòi Nhà máy)"
            cfg_b_ig = {
                "STT": st.column_config.NumberColumn("STT", width="small"),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                "Số mục bỏ qua": st.column_config.NumberColumn("Số mục", width="small"),
                "Tiền xưởng": st.column_config.NumberColumn(format="%,d đ", width="medium"),
                "Lý do / Trạng thái": st.column_config.TextColumn("Lý do / Trạng thái", width="large"),
                "Ghi chú kế toán": st.column_config.TextColumn("Ghi chú kế toán", width="medium")
            }
            st.dataframe(df_b_ig_sum[['STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', 'Số mục bỏ qua', 'Tiền xưởng', 'Lý do / Trạng thái', 'Ghi chú kế toán']], use_container_width=True, hide_index=True, column_config=cfg_b_ig)
        else:
            st.info("Chưa có lệnh nào được gán nhãn Bỏ qua hoặc Không duyệt.")
