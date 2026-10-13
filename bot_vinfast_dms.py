import os
import re
import time
import json
import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile_vinfast")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# URL mục tiêu chốt mốc dữ liệu
NGAY_BAT_DAU = "2026-08-29"

def lay_url_theo_ngay():
    hom_nay = datetime.datetime.now().strftime("%Y-%m-%d")
    return f"https://dealer-portal.vinfast.com/dealer/aftersales/work-orders#woDatePreset=CUSTOM&woFrom={NGAY_BAT_DAU}&woLimit=20&woPage=1&woTo={hom_nay}"

def ghi_nhat_ky(new_cnt, upd_cnt, tot_cnt, trang_thai="Thành công"):
    thoi_gian_hien_tai = datetime.datetime.now().strftime("%H:%M - %d/%m/%Y")
    data_sync = {
        "thoi_gian": thoi_gian_hien_tai,
        "trang_thai": trang_thai,
        "so_lenh_moi": new_cnt,
        "so_lenh_cap_nhat": upd_cnt,
        "tong_so_lenh": tot_cnt
    }
    file_log = os.path.join(BASE_DIR, "last_sync.json")
    with open(file_log, "w", encoding="utf-8") as f:
        json.dump(data_sync, f, ensure_ascii=False, indent=4)
    print(f"✅ Đã ghi nhận lịch sử đồng bộ lúc: {thoi_gian_hien_tai}")

def nap_du_lieu_vao_master(filepath_excel):
    """Sử dụng logic của app.py để nạp dữ liệu và đẩy lên Google Sheets"""
    try:
        from app import load_data_from_gsheets, save_data_to_gsheets, doc_file_db, clean_lsc_giu_gach, norm_lsc_key, clean_tien_series, COT_TIEN, TAT_CA_COT
        
        print("☁️ Đang đọc MasterData từ Google Sheets...")
        df_master = load_data_from_gsheets()
        
        print(f"📂 Đang phân tích file {filepath_excel}...")
        df_raw = doc_file_db(filepath_excel)
        
        df_inc = pd.DataFrame()
        for c in TAT_CA_COT:
            df_inc[c] = df_raw[c] if c in df_raw.columns else ""

        for c in COT_TIEN:
            df_inc[c] = pd.to_numeric(clean_tien_series(df_inc[c]), errors='coerce').fillna(0.0)

        df_inc['Số lệnh sửa chữa'] = df_inc['Số lệnh sửa chữa'].astype(str).apply(clean_lsc_giu_gach)

        # Tự động phân loại luồng
        def phan_loai_theo_tien(r):
            if float(r.get('BH hãng thanh toán', 0) or 0) > 0: return "Bảo Hành Hãng"
            if float(r.get('BH thanh toán', 0) or 0) > 0: return "Bảo Hiểm"
            if float(r.get('Nội bộ thanh toán', 0) or 0) > 0: return "Nội Bộ / PDI"
            return "KH Thông Thường"

        df_inc['Phân loại KH'] = df_inc.apply(phan_loai_theo_tien, axis=1)

        COT_BAO_VE = ['Số hóa đơn', 'Số HĐ', 'Ngày xuất hóa đơn', 'Ngày xuất hoá đơn', 'Ngày HĐ', 'Tiền HĐ', 'Giá trị xuất hóa đơn', 'Ghi chú']
        master_norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
        
        new_records = []
        status_updated_cnt = 0

        for _, row in df_inc.iterrows():
            k_norm = norm_lsc_key(row['Số lệnh sửa chữa'])
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

                for col in COT_TIEN:
                    if col in df_master.columns and col in row:
                        df_master.at[m_idx, col] = float(row[col])

                for col in df_inc.columns:
                    if col in COT_BAO_VE or col in COT_TIEN or col == 'Số lệnh sửa chữa': continue
                    if col in df_master.columns:
                        v = str(row[col]).strip()
                        if v and v not in ['nan', 'None']:
                            df_master.at[m_idx, col] = row[col]

        if new_records:
            df_master = pd.concat([df_master, pd.DataFrame(new_records)], ignore_index=True)

        print("🚀 Đang đẩy dữ liệu lên Google Sheets...")
        save_data_to_gsheets(df_master)
        return len(new_records), status_updated_cnt, len(df_master)
    except Exception as e:
        print(f"❌ Lỗi xử lý dữ liệu: {e}")
        return -1, -1, -1

def chay_bot():
    url_target = lay_url_theo_ngay()
    print("=" * 60)
    print("🚀 BẮT ĐẦU BOT LẤY DỮ LIỆU TỰ ĐỘNG VINFAST DMS")
    print(f"🔗 URL: {url_target}")
    print("=" * 60)

    filepath_final = None

    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            args=["--start-maximized"],
            downloads_path=DOWNLOAD_DIR
        )
        page = browser.new_page()
        
        # 1. Truy cập trực tiếp trang Lệnh sửa chữa kèm query ngày
        page.goto(url_target, timeout=60000)
        time.sleep(4)

        # 2. Điều hướng qua Menu nếu trang bị đẩy về Dashboard
        if "work-orders" not in page.url:
            print("🧭 Đang điều hướng từ Menu vào 'Báo giá & sửa chữa -> Lệnh sửa chữa'...")
            try:
                menu_hm = page.locator("text='Hậu mãi'")
                if menu_hm.is_visible():
                    menu_hm.click()
                    time.sleep(1)

                nhom_bg = page.locator("text='Báo giá & sửa chữa'")
                if nhom_bg.is_visible():
                    nhom_bg.click()
                    time.sleep(1)

                item_lsc = page.locator("text='Lệnh sửa chữa'").first
                item_lsc.click()
                time.sleep(3)
            except Exception as e:
                print(f"Lưu ý điều hướng menu: {e}")

        # Đảm bảo trang đã nạp xong
        page.wait_for_load_state("networkidle")
        time.sleep(2)

        # 3. Thao tác Xuất dữ liệu và Tải file
        try:
            print("📥 Đang tìm và bấm nút 'Xuất dữ liệu'...")
            btn_export = page.locator("button:has-text('Xuất dữ liệu')").first
            btn_export.wait_for(state="visible", timeout=25000)
            btn_export.click()
            time.sleep(3)

            # Kiểm tra Modal xuất dữ liệu
            modal = page.locator(".ant-modal, .modal, div[role='dialog']").first
            if modal.is_visible():
                print("📌 Đang xử lý tick chọn đại lý trong Modal...")
                cb = modal.locator("input[type='checkbox']").first
                if cb.count() > 0 and not cb.is_checked():
                    cb.check()
                    time.sleep(1)

            # Chờ file tạo xong và xuất hiện nút 'Tải file xuống'
            print("⏳ Đang chờ hệ thống tạo file và bấm nút 'Tải file xuống'...")
            btn_download = page.locator("button:has-text('Tải file xuống'), button:has-text('Tải xuống')").last
            btn_download.wait_for(state="visible", timeout=60000)
            btn_download.scroll_into_view_if_needed()

            # Bắt luồng tải file về máy
            with page.expect_download(timeout=60000) as download_info:
                btn_download.click(force=True)

            download = download_info.value
            
            # Lưu đè vào file database.xlsx trong thư mục dự án
            filepath_final = os.path.join(BASE_DIR, "database.xlsx")
            download.save_as(filepath_final)
            print(f"💾 ĐÃ TẢI VÀ LƯU THÀNH CÔNG: {filepath_final}")

        except Exception as e:
            print(f"⚠️ Lỗi khi xuất và tải file: {e}")

        browser.close()

    # 4. Nạp trực tiếp vào Google Sheets
    if filepath_final and os.path.exists(filepath_final):
        print("🔄 Đang đẩy số liệu vào Google Sheets (MasterData)...")
        new_cnt, upd_cnt, tot_cnt = nap_du_lieu_vao_master(filepath_final)
        if new_cnt >= 0:
            ghi_nhat_ky(new_cnt, upd_cnt, tot_cnt, trang_thai="Thành công")
            print("=" * 60)
            print("🎉 HOÀN THÀNH TOÀN BỘ TIẾN TRÌNH:")
            print(f" - Lệnh mới: +{new_cnt}")
            print(f" - Lệnh cập nhật: {upd_cnt}")
            print(f" - Tổng dữ liệu Google Sheets: {tot_cnt}")
            print("=" * 60)
        else:
            ghi_nhat_ky(0, 0, 0, trang_thai="Lỗi nạp Google Sheets")
    else:
        print("❌ Không tìm thấy file database.xlsx tải về.")

if __name__ == "__main__":
    chay_bot()
