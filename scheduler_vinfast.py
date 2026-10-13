"""
scheduler_vinfast.py - Tiến trình chạy ngầm định kỳ hoặc kích hoạt khi bấm nút trên web
"""

import os
import sys
import pandas as pd
from datetime import datetime
import bot_vinfast

def chay_tien_trinh_dong_bo():
    now_str = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    print(f"[*] Bắt đầu đồng bộ tự động VinFast: {now_str}")
    
    # 1. Kéo file từ VinFast
    res_vinfast = bot_vinfast.tu_dong_xuat_file_vinfast(headless=True)
    if not res_vinfast.get("success"):
        print(f"[!] Thất bại: {res_vinfast.get('message')}")
        return

    file_db_path = res_vinfast.get("file_path")
    if not file_db_path or not os.path.exists(file_db_path):
        return

    # 2. Đọc file và nạp lên Google Sheets
    try:
        import app as main_app
        df_master = main_app.load_data_from_gsheets()
        
        with open(file_db_path, "rb") as f:
            df_raw = main_app.doc_file_db(f)

        df_inc = pd.DataFrame()
        for c in main_app.TAT_CA_COT:
            df_inc[c] = df_raw[c] if c in df_raw.columns else ""

        for c in main_app.COT_TIEN:
            df_inc[c] = pd.to_numeric(main_app.clean_tien_series(df_inc[c]), errors='coerce').fillna(0.0)

        df_inc['Số lệnh sửa chữa'] = df_inc['Số lệnh sửa chữa'].astype(str).apply(main_app.clean_lsc_giu_gach)

        mask_co_ck = (df_inc['Tổng trước chiết khấu'] > 0) & (df_inc['Chiết khấu đại lý'] > 0)
        df_inc.loc[mask_co_ck, 'Tổng sau chiết khấu'] = df_inc.loc[mask_co_ck, 'Tổng trước chiết khấu'] - df_inc.loc[mask_co_ck, 'Chiết khấu đại lý']

        def phan_loai_theo_tien(r):
            bh_hang = float(r.get('BH hãng thanh toán', 0) or 0)
            bh = float(r.get('BH thanh toán', 0) or 0)
            noi_bo = float(r.get('Nội bộ thanh toán', 0) or 0)
            kh = float(r.get('KH thanh toán', 0) or 0)
            if bh_hang > 0: return "Bảo Hành Hãng"
            elif bh > 0: return "Bảo Hiểm"
            elif noi_bo > 0: return "Nội Bộ / PDI"
            elif kh > 0: return "KH Thông Thường"
            return "KH Thông Thường"

        df_inc['Phân loại KH'] = df_inc.apply(phan_loai_theo_tien, axis=1)
        df_inc.loc[(df_inc['BH hãng thanh toán'] > 0) & (df_inc['Phê duyệt bảo hành'].isna() | (df_inc['Phê duyệt bảo hành'] == '')), 'Phê duyệt bảo hành'] = "Chờ duyệt"

        df_inc = main_app.dong_bo_hoa_don(df_inc)
        df_inc = main_app.chuan_hoa_kieu_du_lieu(df_inc)
        df_inc = main_app.loc_chuan_tu_29_thang_8(df_inc)

        df_master = df_master.reset_index(drop=True)
        master_norm_map = {main_app.norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
        new_records = []
        status_updated_cnt = 0

        for _, row in df_inc.iterrows():
            lsc = row['Số lệnh sửa chữa']
            k_norm = main_app.norm_lsc_key(lsc)
            if not k_norm: continue

            if k_norm not in master_norm_map:
                new_records.append(row)
                master_norm_map[k_norm] = len(df_master) + len(new_records) - 1
            else:
                m_idx = master_norm_map[k_norm]
                old_status = str(df_master.at[m_idx, 'Trạng thái'])
                new_status = str(row.get('Trạng thái', ''))
                if new_status and new_status not in ['nan', 'None'] and old_status != new_status:
                    df_master.at[m_idx, 'Trạng thái'] = new_status
                    status_updated_cnt += 1

                for col in main_app.COT_TIEN:
                    if col in df_master.columns and col in row:
                        df_master.at[m_idx, col] = float(row[col])

                for col in df_inc.columns:
                    if col in main_app.COT_HOA_DON or col in main_app.COT_TIEN or col == 'Số lệnh sửa chữa':
                        continue
                    if col in df_master.columns:
                        val_new = str(row[col]).strip()
                        if val_new and val_new not in ['nan', 'None']:
                            df_master.at[m_idx, col] = row[col]

        if new_records:
            df_master = pd.concat([df_master, pd.DataFrame(new_records)], ignore_index=True)

        main_app.save_data_to_gsheets(df_master)
        print(f"[+] ĐÃ ĐỒNG BỘ THÀNH CÔNG LÊN GOOGLE SHEETS! Thêm mới: {len(new_records)}, cập nhật: {status_updated_cnt}")

    except Exception as e:
        print(f"[-] Lỗi đồng bộ: {e}")

if __name__ == "__main__":
    chay_tien_trinh_dong_bo()
