"""
bot_cyber.py - Module điều khiển tự động phần mềm Cyber Enterprise 9.0 - R3 trên Windows
"""

import os
import time
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

CYBER_EXPORT_FILE = os.path.join(DOWNLOAD_DIR, "cyber_export.xlsx")
NGAY_BAT_DAU = "29/08/2026"

def tu_dong_xuat_bao_cao_cyber() -> dict:
    """
    Kết nối vào cửa sổ Cyber Enterprise 9.0 đang mở, tự động chọn điều kiện và xuất Excel.
    """
    try:
        from pywinauto import Application, Desktop
        import pywinauto.keyboard as keyboard
    except ImportError:
        return {
            "success": False,
            "message": "Chưa cài đặt thư viện pywinauto trên máy tính."
        }

    now_str = datetime.now().strftime("%d/%m/%Y")

    try:
        app = None
        try:
            app = Application(backend="win32").connect(title_re=".*CYBER ENTERPRISE.*", timeout=5)
        except Exception:
            try:
                app = Application(backend="uia").connect(title_re=".*CYBER ENTERPRISE.*", timeout=5)
            except Exception:
                pass

        if not app:
            return {
                "success": False,
                "message": "Không tìm thấy phần mềm Cyber đang mở! Vui lòng mở và đăng nhập Cyber Enterprise trước khi bấm nút."
            }

        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(1)

        # Mở Menu Báo cáo quản trị dịch vụ -> Bảng tổng hợp lệnh theo 2 chỉ tiêu
        try:
            main_window.menu_select("Báo cáo -> Báo cáo quản trị dịch vụ -> 3. Báo cáo theo lệnh -> Bảng tổng hợp lệnh theo 2 chỉ tiêu")
        except Exception:
            keyboard.send_keys("%b")  # Alt + B
            time.sleep(0.5)
            keyboard.send_keys("{DOWN 8}{RIGHT}{DOWN 2}{RIGHT}{DOWN 2}{ENTER}")

        time.sleep(2)

        # Tìm popup Điều kiện lọc báo cáo
        dlg = None
        try:
            dlg = app.window(title_re=".*Điều kiện lọc báo cáo.*")
            dlg.wait('visible', timeout=10)
            dlg.set_focus()
        except Exception:
            desktop = Desktop(backend="win32")
            dlg = desktop.window(title_re=".*Điều kiện lọc báo cáo.*")
            dlg.set_focus()

        # Chọn Tổng hợp theo lệnh và điền dải ngày
        keyboard.send_keys("{DOWN}")
        time.sleep(0.5)
        keyboard.send_keys("{TAB}")
        keyboard.send_keys(f"{NGAY_BAT_DAU}")
        time.sleep(0.5)
        keyboard.send_keys("{TAB}")
        keyboard.send_keys(f"{now_str}")
        time.sleep(0.5)

        # Bấm Chấp nhận
        try:
            dlg.child_window(title="Chấp nhận", class_name="Button").click()
        except Exception:
            keyboard.send_keys("%c")
            keyboard.send_keys("{ENTER}")

        time.sleep(4)

        # Xuất Excel bằng Ctrl + E
        keyboard.send_keys("^e")
        time.sleep(2)

        # Lưu file
        try:
            save_dlg = app.window(title_re=".*Save.*|.*Lưu.*|.*Xuất.*")
            if save_dlg.exists(timeout=5):
                save_dlg.set_focus()
                keyboard.send_keys(CYBER_EXPORT_FILE.replace("\\", "/"))
                keyboard.send_keys("{ENTER}")
                time.sleep(2)
                keyboard.send_keys("%y")
        except Exception:
            pass

        time.sleep(2)

        return {
            "success": True,
            "file_path": CYBER_EXPORT_FILE,
            "message": "Đã tự động lấy bảng tổng hợp lệnh từ Cyber thành công!"
        }

    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "message": f"Không thể tự động xuất từ Cyber: {e}. Bạn vẫn có thể kéo thả file thủ công như cũ."
        }

if __name__ == "__main__":
    res = tu_dong_xuat_bao_cao_cyber()
    print("Kết quả:", res)
