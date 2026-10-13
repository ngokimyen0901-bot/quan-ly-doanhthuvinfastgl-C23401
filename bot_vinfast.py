import json
import datetime
import os

def ghi_nhat_ky_dong_bo(new_cnt, upd_cnt, tot_cnt, trang_thai="Thành công"):
    thoi_gian_hien_tai = datetime.datetime.now().strftime("%H:%M:%S ngày %d/%m/%Y")
    data_sync = {
        "thoi_gian": thoi_gian_hien_tai,
        "trang_thai": trang_thai,
        "so_lenh_moi": new_cnt,
        "so_lenh_cap_nhat": upd_cnt,
        "tong_so_lenh": tot_cnt
    }
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_sync.json")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data_sync, f, ensure_ascii=False, indent=4)
    print(f"Đã lưu nhật ký đồng bộ: {thoi_gian_hien_tai}")
