import os
import sys
import time
import glob
import re
import zipfile
import pickle
import base64
from cryptography.fernet import Fernet
import shutil
from datetime import datetime, timedelta, date
import pandas as pd
from openpyxl import load_workbook
import requests
import pyotp
from playwright.sync_api import sync_playwright
from concurrent.futures import ThreadPoolExecutor
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Tự động tìm đường dẫn đến Desktop của User
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
EXCEL_FILE = os.path.join(DESKTOP_PATH, "CVUP_Cmd_uyen.xlsm")

# Biến toàn cục để theo dõi trạng thái thực thi của từng dòng Media
MEDIA_STATUS_REPORT = {}

# Biến toàn cục lưu danh sách các hàng cần gửi mail và tạo log sau khi tiến trình tải hoàn tất
PENDING_MAIL_ROWS = []

import threading
mail_lock = threading.Lock()

def get_current_date_str():
    return datetime.now().strftime("%Y%m%d")

def print_ascii_cat():
    """
    In hình chú mèo lười đang ngủ ZZZzz tối giản, màu cam nổi bật trên Terminal kèm khung 100 dấu bằng.
    """
    CAM = "\033[38;5;208m"
    RESET = "\033[0m"

    cat_art = r"""
        |\      _,,,---,,_
ZZZzz /,`.-'`'    -.  ;-;;,_
     |,4-  ) )-,_. ,\ (  `'-'
    '---''(_/--'  `-'\_)
    """
    
    print(CAM + "=" * 100 + RESET)
    print(CAM + cat_art + RESET)
    print(CAM + "=" * 100 + RESET)

def write_error_log(today_dir, message):
    """
    Ghi nhận lỗi chi tiết vào file error_log.txt nằm trong thư mục ngày của Media
    """
    try:
        os.makedirs(today_dir, exist_ok=True)
        log_path = os.path.join(today_dir, "error_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as log_err:
        print(f"      ! Khong the ghi file log loi: {log_err}")

def create_execution_log(today_dir, account_info, cmd_row, current_media, dl_status, mail_status, status_msg, current_acc_id=None):
    """
    Tạo file log gộp dữ liệu chuẩn format: log_{số thứ tự}_{CAD}_{media}.csv
    Nếu 1 dòng tải nhiều media hoặc nhiều acc, dữ liệu sẽ gộp chung vào 1 file log, chia thành các dòng riêng biệt.
    """
    try:
        os.makedirs(today_dir, exist_ok=True)
        
        cad_prefix = "CAD"
        for col_key in account_info.index:
            if str(col_key).strip().upper() == "CAD":
                cad_prefix = str(account_info[col_key]).strip()
                break

        # Tên file log cố định theo quy tắc: không chứa account_id riêng biệt ở tên file nữa
        log_file_name = f"log_{cmd_row['row_stt']}_{cad_prefix}_{current_media}.csv"
        log_path = os.path.join(today_dir, log_file_name)

        af_copy = account_info.copy()
        if current_acc_id:
            af_copy['account_id'] = current_acc_id

        df_af = af_copy.to_frame().T.add_prefix("AF_")
        df_cmd = cmd_row.to_frame().T.add_prefix("Cmd_")
        
        df_log_combined = pd.concat([df_af.reset_index(drop=True), df_cmd.reset_index(drop=True)], axis=1)
        df_log_combined.insert(0, "Log_Execution_Time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        df_log_combined["Download"] = dl_status
        df_log_combined["Mail"] = mail_status
        df_log_combined["Log_Detail_Status"] = status_msg
        
        # Nếu file log của dòng này/media này đã tồn tại trước đó (do account hoặc media khác tạo trước)
        if os.path.exists(log_path):
            try:
                df_old = pd.read_csv(log_path)
                # Loại bỏ dòng cũ nếu trùng cả account_id và media để tránh ghi đè lặp dữ liệu
                if "AF_account_id" in df_old.columns and current_acc_id:
                    df_old = df_old[df_old["AF_account_id"].astype(str) != str(current_acc_id)]
                
                df_final_log = pd.concat([df_old, df_log_combined], ignore_index=True)
                df_final_log.to_csv(log_path, index=False, encoding="utf-8-sig")
            except Exception:
                df_log_combined.to_csv(log_path, index=False, encoding="utf-8-sig")
        else:
            df_log_combined.to_csv(log_path, index=False, encoding="utf-8-sig")
            
        print(f"      => [LOG] Da cap nhat file log dong bo: {log_path}")
    except Exception as log_err:
        print(f"      ! Khong the tao file log gop du lieu: {log_err}")

def calculate_date_range(time_range_str):
    """
    Hàm tính toán chính xác danh sách ngày (target_dates) dựa trên cột 'time' trong sheet AF_Cmd
    """
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    date_list = []
    
    tr_clean = str(time_range_str).strip()
    tr_lower = tr_clean.lower()
    
    if tr_lower == "today":
        date_list.append(today)
    elif tr_lower == "yesterday":
        date_list.append(yesterday)
    elif tr_lower == "this month":
        start_date = today.replace(day=1)
        current = start_date
        while current <= yesterday:
            date_list.append(current)
            current += timedelta(days=1)
    elif tr_lower == "this month (include today)":
        start_date = today.replace(day=1)
        current = start_date
        while current <= today:
            date_list.append(current)
            current += timedelta(days=1)
    elif tr_lower == "last month":
        first_of_this_month = today.replace(day=1)
        last_of_last_month = first_of_this_month - timedelta(days=1)
        first_of_last_month = last_of_last_month.replace(day=1)
        current = first_of_last_month
        while current <= last_of_last_month:
            date_list.append(current)
            current += timedelta(days=1)
    elif tr_lower.startswith("custom;"):
        parts = tr_clean.split(";")
        if len(parts) >= 3:
            try:
                start_dt = datetime.strptime(parts[1].strip(), "%Y%m%d").date()
                end_dt = datetime.strptime(parts[2].strip(), "%Y%m%d").date()
                current = start_dt
                while current <= end_dt:
                    date_list.append(current)
                    current += timedelta(days=1)
            except Exception as e:
                print(f"      [LỖI ĐỊNH DẠNG] Không thể phân tích chuỗi ngày Custom '{tr_clean}': {e}")
    elif "last" in tr_lower and "day" in tr_lower:
        match = re.search(r'last\s+(\d+)\s+day', tr_lower)
        days_count = int(match.group(1)) if match else 1
        for i in range(days_count, 0, -1):
            date_list.append(today - timedelta(days=i))
    elif "last" in tr_lower and "month" in tr_lower:
        match = re.search(r'last\s+(\d+)\s+month', tr_lower)
        months_count = int(match.group(1)) if match else 1
        if today.day == 1:
            first_of_this_month = today.replace(day=1)
            last_of_target_month = first_of_this_month - timedelta(days=1)
            start_year = today.year
            start_month = today.month - months_count
            while start_month <= 0:
                start_month += 12
                start_year -= 1
            start_of_target_month = date(start_year, start_month, 1)
            current = start_of_target_month
            while current <= last_of_target_month:
                date_list.append(current)
                current += timedelta(days=1)
        else:
            start_year = today.year
            start_month = today.month - (months_count - 1)
            while start_month <= 0:
                start_month += 12
                start_year -= 1
            start_of_target_month = date(start_year, start_month, 1)
            current = start_of_target_month
            while current <= yesterday:
                date_list.append(current)
                current += timedelta(days=1)
    else:
        date_list.append(yesterday)
        
    return sorted(list(set(date_list)))

def read_excel_safely(file_path):
    wb = load_workbook(file_path, data_only=True)
    
    ws_cmd = wb["AF_Cmd"]
    data_cmd = []
    for row in ws_cmd.iter_rows(values_only=True):
        data_cmd.append(list(row))
    while len(data_cmd) < 12:
        data_cmd.append([None] * 20)
    header_cmd = [str(x).strip() if x is not None else "" for x in data_cmd[10]]
    rows_cmd = data_cmd[11:]
    df_cmd = pd.DataFrame(rows_cmd, columns=header_cmd)
    df_cmd.columns = [c.lower() for c in df_cmd.columns]
    df_cmd['row_stt'] = range(12, len(df_cmd) + 12)
    
    ws_af = wb["AF"]
    data_af = []
    for row in ws_af.iter_rows(values_only=True):
        data_af.append(list(row))
    while len(data_af) < 2:
        data_af.append([None] * 20)
    header_af = [str(x).strip() if x is not None else "" for x in data_af[0]]
    rows_af = data_af[1:]
    df_af = pd.DataFrame(rows_af, columns=header_af)
    df_af['row_stt'] = range(2, len(df_af) + 2)
    
    wb.close()
    return df_cmd, df_af


# Thư mục lưu cookie cố định theo yêu cầu của bạn
COOKIE_DIR = r"C:\Users\S24870\Desktop\Ifmt"

def decode_password(encoded_pw):
    """
    Giải mã mật khẩu bằng Fernet Token dựa trên Khóa bảo mật của hệ thống
    """
    try:
        # Khóa bảo mật lấy từ màn hình decode_pass.py của bạn
        SECRET_KEY = b"DopIo6rNYYXrw1CVGggHGR2bTbZC_czE9w4QRdkAU5M="
        
        f = Fernet(SECRET_KEY)
        # Tiến hành giải mã chuỗi gAAAAA...
        decoded_bytes = f.decrypt(encoded_pw.strip().encode('utf-8'))
        return decoded_bytes.decode('utf-8')
    except Exception as e:
        print(f"      => [Cảnh báo] Lỗi giải mã Fernet (Có thể pass là text thô): {e}")
        return encoded_pw


# ========================================================
# HÀM TRỢ GIÚP LỌC NGÀY / THÁNG BỊ THIẾU
# ========================================================
def get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload):
    """Lọc và trả về danh sách các ngày thực sự cần tải (chưa có file hoặc bị ép tải lại)"""
    dates_to_download = []
    for single_date in target_dates:
        t_str = single_date.strftime("%Y%m%d")
        file_std = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_str}.csv")
        if force_redownload or not os.path.exists(file_std):
            dates_to_download.append(single_date)
    return dates_to_download

def get_missing_monthly_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload):
    """Lọc và trả về danh sách các tháng (PPDSP) thực sự cần tải"""
    months_to_download = sorted(list(set([d.strftime("%Y-%m") for d in target_dates])))
    months_need_action = []
    for year_month in months_to_download:
        ym_clean = year_month.replace('-', '')
        expected_file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{ym_clean}.csv")
        if force_redownload or not os.path.exists(expected_file_path):
            months_need_action.append(year_month)
    return months_need_action


# ========================================================
# MODULE CHẠY RIÊNG CHO LOOKER (BẢN VÁ BYPASS ACS TUYỆT ĐỐI)
# ========================================================
def run_looker(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"    -> [Looker Row {row_stt} | ID {current_acc_id}] Luồng chính yêu cầu quét / tải dữ liệu...")
    
    # 1. Lấy thông tin tài khoản cơ bản
    login_id = str(account_info.get('login_id', '')).strip()
    casso_id = str(account_info.get('casso_id', '')).strip()
    casso_pw = str(account_info.get('casso_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    if not target_dates:
        print(f"      => Không có ngày cần tải cho Looker.")
        return
        
    start_date = min(target_dates)
    end_date = max(target_dates)
    t_start_str = start_date.strftime("%Y%m%d")
    t_end_str = end_date.strftime("%Y%m%d")
    
    expected_file_name = f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_start_str}_{t_end_str}.csv"
    master_file_path = os.path.join(processing_dir, expected_file_name)
    
    clean_acc_id = str(current_acc_id).strip()
    data_url = f"https://datastudio.google.com/u/0/reporting/{clean_acc_id}"
    
    # Tạo thư mục lưu Profile tạm thời
    user_data_dir = os.path.join(COOKIE_DIR, f"chrome_prof_{row_stt}_{clean_acc_id[:6]}")
    os.makedirs(user_data_dir, exist_ok=True)
    
    print(f"      => Khởi động luồng trình duyệt Chrome Thật cải tiến...")
    
    with sync_playwright() as p:
        # Loại bỏ sạch các flag bị Google quét gắt gao (--no-sandbox, --disable-infobars)
        # Thêm các tham số ngăn chặn lưu mật khẩu/popup phá đám luồng đi
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            channel="chrome",
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--password-store=basic", # Tránh gọi credentials manager của Windows
                "--disable-features=PasswordGeneration,PasswordLeakDetection",
                "--disable-popup-blocking"
            ],
            ignore_default_args=["--enable-automation"], # Loại bỏ cờ báo bot mặc định của Playwright
            viewport={"width": 1366, "height": 768}
        )
        
        page = context.pages[0] if context.pages else context.new_page()
        
        try:
            # Bước làm sạch ban đầu
            print("   + [Web Looker] Làm sạch lịch sử session ngầm trên Chrome...")
            page.goto("https://accounts.google.com/Logout", wait_until="load")
            time.sleep(2.0)
            
            print(f"   + [Web Looker] Điều hướng trực tiếp đến link báo cáo: {data_url}")
            page.goto(data_url, wait_until="load")
            time.sleep(4.0)
            
            # Khử lỗi ACS nếu lỡ dính ngay từ đầu trang
            if "acs" in page.url.lower() or page.locator("text=ログアウト, text=Logout").first.count() > 0:
                print("   + [Web Looker] Phát hiện kẹt trang lỗi ACS đầu phiên. Tiến hành nhấn Logout để giải phóng...")
                logout_btn = page.locator("text=ログアウト, text=Logout").first
                if logout_btn.count() > 0:
                    logout_btn.click(force=True)
                    time.sleep(4.0)
                page.goto(data_url, wait_until="load")
                time.sleep(4.0)

            # Định vị ô nhập Email ban đầu của Google
            email_box = page.locator("input[type='email'], input[name='identifier']").first
            
            if email_box.count() > 0:
                print("   + [Web Looker] Đang tiến hành nhập login_id...")
                email_box.wait_for(state="visible", timeout=15000)
                email_box.fill(login_id)
                
                next_btn = page.locator("#identifierNext, button:has-text('次へ'), button:has-text('Next')").first
                next_btn.click(force=True)
                
                print("   + [Web Looker] Đang theo dõi tiến trình chuyển hướng bảo mật...")
                
                is_casso_page = False
                is_password_page = False
                
                for _ in range(30):
                    time.sleep(0.5)
                    current_url = page.url.lower()
                    
                    if "datastudio.google.com" in current_url or "lookerstudio.google.com" in current_url:
                        if "signin" not in current_url:
                            break
                    
                    if "casso" in current_url or page.locator("input#username, input[name='pf.username']").count() > 0:
                        is_casso_page = True
                        break
                        
                    pwd_input = page.locator("input[type='password']:not([aria-hidden='true'])").first
                    if "password" in current_url or (pwd_input.count() > 0 and pwd_input.is_visible()):
                        is_password_page = True
                        break

                # --- XỬ LÝ ĐĂNG NHẬP KHI ĐẾN MÀN HÌNH XÁC THỰC ---
                if is_casso_page or is_password_page:
                    if is_casso_page:
                        print("   + [Web Looker] Đã đến màn hình CASSO. Điền thông tin đăng nhập...")
                        page.locator("input#username, input[name='pf.username']").wait_for(state="visible", timeout=15000)
                        page.locator("input#username, input[name='pf.username']").fill(casso_id)
                        
                        real_pw = decode_password(casso_pw)
                        page.locator("input#password, input[name='pf.pass']").fill(real_pw)
                        
                        signon_btn = page.locator("button:has-text('サインオン'), input[type='submit'][value='サインオン']").first
                        if signon_btn.count() > 0:
                            signon_btn.click(force=True)
                        else:
                            page.locator("input#password, input[name='pf.pass']").press("Enter")
                    
                    elif is_password_page:
                        print("   + [Web Looker] Đã đến màn hình Mật khẩu Google Workspace. Tiến hành nhập...")
                        real_pw = decode_password(casso_pw)
                        page.locator("input[type='password']").fill(real_pw)
                        page.locator("input[type='password']").press("Enter")
                    
                    # --- ĐỢI VÀ XỬ LÝ ĐIỀU HƯỚNG / BẪY BYPASS ACS ---
                    print("   + [Web Looker] Đang đợi trạng thái chứng thực [認証済み]...")
                    max_wait_seconds = 40
                    authenticated = False
                    
                    for i in range(max_wait_seconds):
                        time.sleep(1.0)
                        current_url = page.url.lower()
                        
                        # CƠ CHẾ VÁ VÀ SỬA LỖI ACS TẠI CHỖ: Nếu phát hiện bị đẩy về trang acs, bấm Logout khẩn cấp rồi quay lại link báo cáo
                        if "acs" in current_url:
                            print("   + [Web Looker] Phát hiện Google đẩy về trang lỗi ACS sau khi nhập mật khẩu. Tiến hành phá bẫy...")
                            logout_link = page.locator("text=ログアウト, text=Logout").first
                            if logout_link.count() > 0:
                                logout_link.click(force=True)
                                time.sleep(4.0)
                            print("   + [Web Looker] Đã giải phóng phiên. Thử quay lại link báo cáo để làm mới token...")
                            page.goto(data_url, wait_until="load")
                            time.sleep(5.0)
                            # Kiểm tra nếu sau khi load lại nó tự động nhận diện phiên sạch và chuyển thẳng vào data
                            if "signin" not in page.url.lower() and ("datastudio" in page.url.lower() or "lookerstudio" in page.url.lower()):
                                authenticated = True
                                break
                        
                        try:
                            content = page.content()
                        except Exception:
                            continue
                            
                        if "認証済み" in content:
                            print("   + [Web Looker] Đã xác nhận trạng thái [認証済み]! Chờ chuyển hướng...")
                            time.sleep(6.0)
                            authenticated = True
                            break
                        
                        continue_btn = page.locator("button:has-text('続行'), button:has-text('Continue')").first
                        if continue_btn.count() > 0:
                            continue_btn.click(force=True)
                            time.sleep(5.0)
                            authenticated = True
                            break
                            
                        if ("datastudio.google.com" in current_url or "lookerstudio.google.com" in current_url) and "signin" not in current_url:
                            authenticated = True
                            break
                    
                    if not authenticated:
                        raise TimeoutError("Quá thời gian chờ nhưng không thấy trạng thái [認証済み] hoặc không thể điều hướng vào trang báo cáo.")
            
            # --- CHỜ LAYOUT BÁO CÁO LOAD ỔN ĐỊNH ---
            print("   + [Web Looker] Đã vào trang dữ liệu Looker Studio thành công. Đang đợi tải bộ lọc lịch...")
            try:
                page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass
            time.sleep(8.0)

            # ====================================================
            # LUỒNG THAY ĐỔI NGÀY BÁO CÁO (ĐA TỐI ƯU CHỐNG ĐỨNG)
            # ====================================================
            print("   + [Web Looker] Đang định vị bộ lọc lịch ngày...")
            date_selectors = [
                "gcb-date-range-picker",
                ".gcb-date-range-picker-container",
                "div.daterangepicker-input",
                ".embedded-component-date-range-picker",
                "div[data-ng-click*='date']",
                "[ng-click*='dateRange']",
                "//span[contains(text(), '/')]/ancestor::div[@role='button' or contains(@class, 'picker')][1]",
                "//div[contains(text(), '2026/') or contains(., '2026/')]" 
            ]
            
            date_dropdown = None
            for retry in range(3):
                for sel in date_selectors:
                    try:
                        locator = page.locator(sel).first
                        if locator.count() > 0:
                            locator.wait_for(state="visible", timeout=3000)
                            date_dropdown = locator
                            print(f"      => Đã tìm thấy bộ lọc ngày bằng selector: {sel}")
                            break
                    except Exception:
                        continue
                if date_dropdown:
                    break
                time.sleep(2.0)
            
            if not date_dropdown:
                fallback_locator = page.locator("//div[contains(text(), '/') and contains(@role, 'button')]").first
                if fallback_locator.count() > 0:
                    date_dropdown = fallback_locator
                    print("      => Đã tìm thấy bộ lọc ngày bằng fallback text selector.")
                else:
                    raise RuntimeError("Không thể định vị được ô chọn ngày trên giao diện Looker Studio.")

            date_dropdown.scroll_into_view_if_needed()
            time.sleep(1.5)
            
            date_dropdown.click(force=True)
            print("   + [Web Looker] Đã click mở thành công bảng lịch thả xuống.")
            time.sleep(4.0)
            
            # --- CHỌN NGÀY BẮT ĐẦU (開始日) ---
            print(f"      => Cấu hình Ngày Bắt Đầu: {start_date.strftime('%Y/%m/%d')}")
            left_ym_btn = page.locator("//div[contains(@class, 'left')]//button[contains(@class, 'month-picker') or contains(text(), '年')]").first
            left_ym_btn.wait_for(state="visible", timeout=5000)
            
            expected_ym_text = f"{start_date.year}年{start_date.month}月"
            if expected_ym_text not in left_ym_btn.inner_text():
                left_ym_btn.click(force=True)
                time.sleep(1.0)
                page.locator(f"//div[contains(@class, 'left')]//div[contains(@class, 'year-item') and text()='{start_date.year}']").first.click(force=True)
                time.sleep(0.5)
                page.locator(f"//div[contains(@class, 'left')]//div[contains(@class, 'month-item') and text()='{start_date.month}月']").first.click(force=True)
                time.sleep(1.0)
            
            page.locator(f"//div[contains(@class, 'left')]//div[contains(@class, 'day') and text()='{start_date.day}']").first.click(force=True)
            time.sleep(1.0)
            
            # --- CHỌN NGÀY KẾT THÚC (終了日) ---
            print(f"      => Cấu hình Ngày Kết Thúc: {end_date.strftime('%Y/%m/%d')}")
            right_ym_btn = page.locator("//div[contains(@class, 'right')]//button[contains(@class, 'month-picker') or contains(text(), '年')]").first
            
            expected_end_ym_text = f"{end_date.year}年{end_date.month}月"
            if expected_end_ym_text not in right_ym_btn.inner_text():
                right_ym_btn.click(force=True)
                time.sleep(1.0)
                page.locator(f"//div[contains(@class, 'right')]//div[contains(@class, 'year-item') and text()='{end_date.year}']").first.click(force=True)
                time.sleep(0.5)
                page.locator(f"//div[contains(@class, 'right')]//div[contains(@class, 'month-item') and text()='{end_date.month}月']").first.click(force=True)
                time.sleep(1.0)
                
            page.locator(f"//div[contains(@class, 'right')]//div[contains(@class, 'day') and text()='{end_date.day}']").first.click(force=True)
            time.sleep(1.0)
            
            apply_btn = page.locator("button:has-text('適用'), button:has-text('Apply')").first
            apply_btn.click(force=True)
            print("      => Đã nhấn nút [適用] hệ thống lịch thành công.")
            time.sleep(6.0)

            # ====================================================
            # LUỒNG HOVER BẢNG DATA VÀ TẢI FILE BÁO CÁO CSV
            # ====================================================
            print("   + [Web Looker] Đang định vị bảng dữ liệu để kích hoạt nút 3 chấm...")
            data_table = page.locator("div[role='grid'], .table-container, .ng-scope.table, div.visual-container").first
            data_table.wait_for(state="visible", timeout=15000)
            
            data_table.hover()
            time.sleep(1.5)
            
            three_dots_btn = page.locator("//div[contains(@class, 'action-button') or @role='button']//i[contains(text(), 'more_vert')] | //button[contains(@aria-label, 'その他のオプション') or contains(@class, 'context-menu')]").first
            three_dots_btn.wait_for(state="visible", timeout=5000)
            three_dots_btn.click(force=True)
            time.sleep(1.5)
            
            export_graph_option = page.locator("text=グラフをエクスポート..., text=Export").first
            export_graph_option.wait_for(state="visible", timeout=5000)
            export_graph_option.click(force=True)
            time.sleep(2.0)
            
            export_data_sub = page.locator("text=データのエクスポート, text=データのエкスポ").first
            export_data_sub.click(force=True)
            time.sleep(2.0)
            
            format_dropdown = page.locator("//div[contains(text(), '形式') or contains(., '形式')]/following::select[1] | //select[contains(@class, 'format-select')]").first
            if format_dropdown.count() > 0:
                format_dropdown.select_option(label="CSV（Excel）")
            else:
                format_div = page.locator("text=CSV（Excel）").first
                if format_div.count() == 0:
                    page.locator("//div[text()='CSV' or text()='Google スプレッドシート']").first.click(force=True)
                    time.sleep(0.5)
                    page.locator("text=CSV（Excel）").first.click(force=True)
            time.sleep(1.0)
            
            green_export_btn = page.locator("button.btn-primary, button:has-text('エクスポート')").first
            
            with page.expect_download(timeout=60000) as download_info:
                green_export_btn.click(force=True)
                
            download = download_info.value
            download.save_as(master_file_path)
            print(f"      => [TẢI THÀNH CÔNG] File kết quả Looker lưu tại: {master_file_path}")
            time.sleep(2.0)

        except Exception as e:
            print(f"      => [LỖI] Tiến trình chạy module Looker thất bại: {e}")
            raise e
        finally:
            context.close()


# ========================================================
# MODULE CHẠY RIÊNG CHO MICROSOFT (TUÂN THỦ LUỒNG CHÍNH)
# ========================================================
def run_microsoft(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"    -> [Microsoft Row {row_stt} | ID {current_acc_id}] Luong chinh yeu cau quet / tai moi du lieu...")
    
    # Lấy thông tin tài khoản cơ bản
    login_id = str(account_info.get('login_id', '')).strip()
    casso_id = str(account_info.get('casso_id', '')).strip()
    casso_pw = str(account_info.get('casso_pw', '')).strip()
    
    cookie_file_name = f"microsoft_{login_id}.txt"
    cookie_file_path = os.path.join(COOKIE_DIR, cookie_file_name)
    
    # ===========================================================================
    # CƠ CHẾ BÓC TÁCH DỮ LIỆU THÔNG MINH
    # ===========================================================================
    raw_dimensions = ""
    raw_metrics = ""
    
    if cmd_info is not None:
        if isinstance(cmd_info, dict):
            for k, v in cmd_info.items():
                if str(k).strip().lower() == "dimensions": raw_dimensions = str(v)
                elif str(k).strip().lower() == "metrics": raw_metrics = str(v)
        elif hasattr(cmd_info, "index"):
            for col in cmd_info.index:
                if str(col).strip().lower() == "dimensions": raw_dimensions = str(cmd_info[col])
                elif str(col).strip().lower() == "metrics": raw_metrics = str(cmd_info[col])

    if (not raw_dimensions or raw_dimensions.lower() == 'nan') and account_info is not None:
        if isinstance(account_info, dict):
            for k, v in account_info.items():
                if str(k).strip().lower() == "dimensions": raw_dimensions = str(v)
                elif str(k).strip().lower() == "metrics": raw_metrics = str(v)
        elif hasattr(account_info, "index"):
            for col in account_info.index:
                if str(col).strip().lower() == "dimensions": raw_dimensions = str(account_info[col])
                elif str(col).strip().lower() == "metrics": raw_metrics = str(account_info[col])

    dimensions = [d.strip() for d in raw_dimensions.split(',') if d.strip() and d.lower() != 'nan']
    metrics = [m.strip() for m in raw_metrics.split(',') if m.strip() and m.lower() != 'nan']
    
    print(f"   => [DEBUG EXCEL - HOÀN CHỈNH] Dimensions thực tế bóc tách: {dimensions}")
    print(f"   => [DEBUG EXCEL - HOÀN CHỈNH] Metrics thực tế bóc tách: {metrics}")

    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    if not target_dates:
        print(f"      => Không có ngày cần tải cho Microsoft.")
        return
        
    start_date = min(target_dates)
    end_date = max(target_dates)
    t_start_str = start_date.strftime("%Y%m%d")
    t_end_str = end_date.strftime("%Y%m%d")
    
    input_start_date = start_date.strftime("%Y/%m/%d")
    input_end_date = end_date.strftime("%Y/%m/%d")
    
    master_file_name = f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_start_str}_{t_end_str}_master.xlsx"
    master_file_path = os.path.join(processing_dir, master_file_name)

    print(f"      => Tien hanh xu ly du lieu. Mo trinh duyet...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768}
        )
        
        cookie_loaded = False
        if os.path.exists(cookie_file_path):
            try:
                with open(cookie_file_path, "rb") as f:
                    cookies = pickle.load(f)
                context.add_cookies(cookies)
                print(f"   + [Web MS Ads] Da nap Cookie tu file: {cookie_file_name}")
                cookie_loaded = True
            except Exception as e:
                print(f"   + [Web MS Ads] Khong the nap cookie cu: {e}")

        page = context.new_page()
        
        try:
            data_url = f"https://ui.ads.microsoft.com/campaign/vnext/reporting/create?{current_acc_id}"
            print(f"   + [Web MS Ads] Dieu huong den link data: {data_url}")
            page.goto(data_url, wait_until="load")
            time.sleep(30.0)
            
            # --- KIỂM TRA ĐĂNG NHẬP ---
            current_url = page.url.lower()
            has_login_input = page.locator("input[name='loginfmt'], input[type='email']").count() > 0
            has_signin_btn = page.locator("a:has-text('サインイン'), button:has-text('サインイン')").count() > 0
            is_sign_in_page = ("signin" in current_url or "login" in current_url or "sessionexpired" in current_url or has_login_input or (has_signin_btn and "microsoft" in current_url))
            
            if is_sign_in_page or not cookie_loaded:
                print("   + [Web MS Ads] Cookie cũ hết hạn. Đang đăng nhập lại...")
                page.goto("https://ui.ads.microsoft.com/", wait_until="networkidle")
                time.sleep(2.0)
                
                cookie_accept_btn = page.locator("button:has-text('承知'), #msccAccept, .cc-cookie-accept").first
                if cookie_accept_btn.count() > 0:
                    cookie_accept_btn.click(force=True)
                
                page.locator("a.bing-sign-in-link, a:has-text('サインイン')").first.click(force=True)
                time.sleep(3.0)
                
                ms_btn = page.locator("button[data-action-name='ms-sign-in'], .sign-in-button-ms").first
                if ms_btn.count() > 0:
                    ms_btn.click(force=True)
                
                page.locator("input[name='loginfmt'], input[type='email']").fill(login_id)
                page.locator("input[type='submit'], #idSIButton9").click(force=True)
                time.sleep(2.0)
                
                page.locator("input#username, input[name='pf.username']").fill(casso_id)
                real_pw = decode_password(casso_pw)
                page.locator("input#password, input[name='pf.pass']").fill(real_pw)
                page.locator("input#password, input[name='pf.pass']").press("Enter")
                
                max_wait_seconds = 60
                authenticated = False
                for i in range(max_wait_seconds):
                    time.sleep(1.0)
                    try:
                        content = page.content()
                        chk_url = page.url.lower()
                    except Exception:
                        continue
                    if "認証中" not in content and "認証済み" in content:
                        authenticated = True
                        break
                    if "stay signed in" in content.lower() or "idsibutton9" in content.lower() or "ui.ads.microsoft.com" in chk_url:
                        authenticated = True
                        break
                
                if not authenticated:
                    raise TimeoutError("Không thể hoàn tất xác thực.")
                
                time.sleep(2.0)
                yes_btn = page.locator("input[type='submit'][value='Yes'], input[type='submit'][value='はい'], input#idSIButton9").first
                if yes_btn.count() > 0:
                    yes_btn.click(force=True)
                    time.sleep(4.0)
                
                os.makedirs(COOKIE_DIR, exist_ok=True)
                with open(cookie_file_path, "wb") as f:
                    pickle.dump(context.cookies(), f)
                
                page.goto(data_url, wait_until="load")
                time.sleep(5.0)

            # ====================================================
            # [FIX MỚI] ĐẢM BẢO CHẮC CHẮN ĐANG Ở TRANG REPORTING/CREATE
            # ====================================================
            current_page_url = page.url.lower()
            if "reporting/create" not in current_page_url:
                print("   + [Web MS Ads] Bị redirect ra trang chủ. Đang điều hướng lại vào trang Report...")
                # Thử load lại link data lần 2 (lúc này bối cảnh tài khoản đã được nạp xong)
                page.goto(data_url, wait_until="load")
                time.sleep(5.0)
                
                # Nếu hệ thống vẫn ngoan cố, dùng UI để click vào menu "レポート" ở thanh bên trái
                if "reporting/create" not in page.url.lower():
                    print("   + [Web MS Ads] Vẫn kẹt ở trang chủ, đang ép mở bằng click UI...")
                    report_menu = page.locator("div[title='レポート'], a[title='レポート'], nav span:has-text('レポート')").first
                    if report_menu.count() > 0:
                        report_menu.click(force=True)
                        time.sleep(3.0)
                    
                    # Force load lại url data lần cuối cùng
                    page.goto(data_url, wait_until="load")
                    time.sleep(5.0)
            # ====================================================

            if not dimensions and not metrics:
                raise ValueError("Danh sách cột trống.")

            # ====================================================
            # LUỒNG THÊM DIMENSIONS (内訳の追加)
            # ====================================================
            for dim in dimensions:
                print(f"   + [Web MS Ads] Dang tim kiem Dimension: {dim}")
                page.locator("//div[text()='内訳']/following-sibling::button | //button[@aria-label='内訳の追加']").first.click(force=True)
                time.sleep(1.5)
                
                search_box = page.locator("//div[contains(., '内訳の追加') or text()='内訳の追加']/following::input[@placeholder='列名' or @type='search'][1]").first
                search_box.focus()
                search_box.press("Control+A")
                search_box.press("Backspace")
                search_box.fill(dim)
                time.sleep(2.5)
                
                target_text_item = page.locator(f"//div[contains(., '内訳の追加') or text()='内訳の追加']/following::*[text()='{dim}']").first
                if target_text_item.count() > 0:
                    target_text_item.click(force=True)
                else:
                    search_box.focus()
                    search_box.press("ArrowDown")
                    time.sleep(0.3)
                    search_box.press("Enter")
                time.sleep(1.5)

            # ====================================================
            # LUỒNG THÊM METRICS (メトリックの追加)
            # ====================================================
            for met in metrics:
                print(f"   + [Web MS Ads] Dang tim kiem Metric: {met}")
                page.locator("//div[text()='メトリクス']/following-sibling::button | //button[@aria-label='メトリックの追加']").first.click(force=True)
                time.sleep(1.5)
                
                search_box = page.locator("//div[contains(., '内訳の追加') or contains(., '追加')]/following::input[@placeholder='列名' or @type='search'][1]").first
                search_box.focus()
                search_box.press("Control+A")
                search_box.press("Backspace")
                search_box.fill(met)
                time.sleep(2.5)
                
                target_text_item = page.locator(f"//div[contains(., '追加')]/following::*[text()='{met}']").first
                if target_text_item.count() > 0:
                    target_text_item.click(force=True)
                else:
                    search_box.focus()
                    search_box.press("ArrowDown")
                    time.sleep(0.3)
                    search_box.press("Enter")
                time.sleep(1.5)

            # ====================================================
            # LUỒNG THAY ĐỔI NGÀY BÁO CÁO (DATE PICKER) - ĐÃ SỬA LỖI SELECTOR
            # ====================================================
            print(f"   + [Web MS Ads] Mở bảng lịch chọn ngày bóc tách dữ liệu...")
            
            # Sử dụng Selector chính xác theo text tiếng Nhật hiển thị trên nút hoặc nút có icon Calendar kề cạnh
            date_picker_toggle = page.locator("//button[contains(@aria-label, '日付範囲')] | //div[contains(@class, 'compared-date-range-picker')]//button").first
            date_picker_toggle.wait_for(state="visible", timeout=30000)
            date_picker_toggle.click(force=True)
            time.sleep(2.5)
            
            # Định vị hai ô input điền ngày dựa vào giá trị chứa dấu gạch chéo '/' dạng YYYY/MM/DD
            date_inputs = page.locator("div.compared-date-range-picker-popup input[type='text'], div.compared-date-range-picker input[type='text'], input[value*='/']")
            
            if date_inputs.count() >= 2:
                # Điền ngày bắt đầu
                print(f"      => Điền Ngày Bắt Đầu: {input_start_date}")
                date_inputs.nth(0).focus()
                date_inputs.nth(0).press("Control+A")
                date_inputs.nth(0).press("Backspace")
                date_inputs.nth(0).fill(input_start_date)
                date_inputs.nth(0).press("Tab")
                time.sleep(0.5)
                
                # Điền ngày kết thúc
                print(f"      => Điền Ngày Kết Thúc: {input_end_date}")
                date_inputs.nth(1).focus()
                date_inputs.nth(1).press("Control+A")
                date_inputs.nth(1).press("Backspace")
                date_inputs.nth(1).fill(input_end_date)
                date_inputs.nth(1).press("Enter")
                time.sleep(1.0)
            else:
                # Phương án dự phòng bằng XPath trực diện vào 2 ô input text kề nhau
                print("      => [Dự phòng] Điền ngày bằng XPath chi tiết...")
                start_input = page.locator("//input[contains(@value, '/')][1]")
                end_input = page.locator("//input[contains(@value, '/')][2]")
                
                start_input.fill(input_start_date)
                time.sleep(0.5)
                end_input.fill(input_end_date)
                time.sleep(0.5)
                
            # Click nút "適用する" (Áp dụng - Chốt ngày)
            apply_btn = page.locator("button:has-text('適用する'), button:has-text('Apply')").first
            if apply_btn.count() > 0:
                apply_btn.click(force=True)
                print("      => Đã nhấn nút [適用する] thành công.")
            else:
                page.keyboard.press("Enter")
                print("      => Tự động áp dụng ngày bằng phím Enter.")
            
            time.sleep(2.5)

            # ====================================================
            # LUỒNG TẢI FILE BÁO CÁO .xlsx 
            # ====================================================
            print("   + [Web MS Ads] Click mo menu tai file (ダウンロード)...")
            download_btn = page.locator("#button-download, button:has-text('ダウンロード'), button:has-text('Download')").first
            download_btn.wait_for(state="visible", timeout=30000)
            download_btn.click(force=True)
            time.sleep(2.0)
            
            excel_option = page.locator("xpath=//span[text()='.xlsx'] | //li[contains(., '.xlsx')] | //button[contains(., '.xlsx')]").first
            excel_option.wait_for(state="visible", timeout=30000)
            
            with page.expect_download(timeout=90000) as download_info:
                excel_option.click(force=True)
                
            download = download_info.value
            download.save_as(master_file_path)
            print(f"      => [TAI XONG] File tho da luu tai: {master_file_path}")

            # ===========================================================================
            # MỚI: XỬ LÝ HẬU KỲ FILE EXCEL (XÓA DÒNG 1-6 & LỌC DÒNG RÁC)
            # ===========================================================================
            print(f"   + [Xử lý Excel] Đang làm sạch dữ liệu tệp {master_file_name}...")
            if os.path.exists(master_file_path):
                # Đọc từ dòng 7 (header=6 trong python vì index bắt đầu từ 0)
                df = pd.read_excel(master_file_path, header=6)
                
                # Loại bỏ khoảng trắng thừa trong tên cột để tránh lỗi so khớp cột '日付'
                df.columns = [str(col).strip() for col in df.columns]
                
                if '日付' in df.columns:
                    # Chuyển đổi toàn bộ cột sang dạng chuỗi để thực hiện lọc text chính xác
                    df['日付'] = df['日付'].astype(str).str.strip()
                    
                    # Điều kiện lọc: loại bỏ dòng chứa "合計" hoặc chứa "Microsoft Corporation"
                    # Giữ lại các dòng KHÔNG thỏa mãn điều kiện rác
                    condition_clean = (
                        (df['日付'] != '合計') & 
                        (~df['日付'].str.contains('Microsoft Corporation', case=False, na=False)) &
                        (df['日付'] != 'nan')
                    )
                    df_cleaned = df[condition_clean]
                    
                    # Ghi đè (Save) lại dữ liệu sạch vào chính file Master ban đầu
                    df_cleaned.to_excel(master_file_path, index=False)
                    print(f"      => [THANH CÔNG] Đã xóa dòng 1-6 và làm sạch dữ liệu thành công!")
                else:
                    print("      => [CẢNH BÁO] Không tìm thấy cột '日付' tại dòng 7 để tiến hành lọc.")

        except Exception as e:
            print(f"      => [LOI] Qua trinh chay Microsoft bi hu hong: {e}")
            raise e
        finally:
            browser.close()


# ========================================================
# MODULE CHẠY RIÊNG CHO SKYFLAG
# ========================================================
def run_skyflag(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [SKYFLAG Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        print(f"      => Thieu hoac can tai lai {len(dates_to_download)} ngay. Mo trinh duyet...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                page.goto("https://dashboard.skyflag.jp/advertiserLogin", wait_until="networkidle")
                time.sleep(1.5)
                
                page.get_by_placeholder("E-mail").first.fill(login_id)
                time.sleep(0.5)
                page.get_by_placeholder("Password").first.fill(login_pw)
                time.sleep(1.0)
                
                login_btn = page.locator("button:has-text('ログイン'), .btn-primary").first
                login_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(2.5)
                
                page.goto("https://dashboard.skyflag.jp/advertiserOfferReport", wait_until="networkidle")
                time.sleep(2.0)
                
                is_first_day = True
                for single_date in dates_to_download:
                    target_date_str = single_date.strftime("%Y%m%d")
                    file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
                    
                    print(f"   + [Web SKYFLAG] Thiet lap thoi gian cho ngay: {single_date.strftime('%Y/%m/%d')}")
                    formatted_input_date = single_date.strftime("%Y/%m/%d")
                    
                    inputs_locator = page.locator("div.v-text-field input, input[placeholder*='/'], input[value*='/']")
                    inputs_locator.first.wait_for(state="visible", timeout=5000)
                    
                    target_input_1 = inputs_locator.nth(0)
                    target_input_1.click()
                    time.sleep(0.3)
                    page.keyboard.press("Control+A")
                    time.sleep(0.1)
                    page.keyboard.type(formatted_input_date, delay=30)
                    time.sleep(0.3)
                    
                    target_input_2 = inputs_locator.nth(1)
                    target_input_2.click()
                    time.sleep(0.3)
                    page.keyboard.press("Control+A")
                    time.sleep(0.1)
                    page.keyboard.type(formatted_input_date, delay=30)
                    time.sleep(0.3)
                    
                    target_input_2.press("Enter")
                    time.sleep(0.5)
                    
                    page.locator("body").click(position={"x": 5, "y": 5})
                    time.sleep(0.5)
                    
                    search_btn_selector = "button.btn-primary, button:has-text('検索')"
                    page.wait_for_selector(search_btn_selector, state="visible", timeout=5000)
                    page.locator(search_btn_selector).filter(has_text="検索").first.click(force=True)
                    
                    page.wait_for_load_state("networkidle")
                    time.sleep(3.0)
                    
                    if "No data" in page.content() or page.locator("text=No data").is_visible():
                        print(f"      [THÔNG BÁO] Ngay {target_date_str} hiển thị 'No data'.")
                        if is_first_day:
                            print("      => [SKIP TOÀN BỘ] Ngay dau tien ko co data. Dung luong va bo qua cac ngay sau!")
                            return
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write("0")
                        is_first_day = False
                        continue
                    
                    download_success = False
                    try:
                        export_btn = page.locator("button[data-testid='exportButton']").first
                        if export_btn.count() == 0 or not export_btn.is_visible():
                            export_btn = page.locator("button.ant-btn-icon-only").filter(has_text="").first

                        export_btn.scroll_into_view_if_needed()
                        time.sleep(0.5)
                        
                        box = export_btn.bounding_box()
                        if box:
                            btn_center_x = box["x"] + box["width"] / 2
                            btn_center_y = box["y"] + box["height"] / 2
                            
                            page.mouse.click(btn_center_x, btn_center_y)
                            time.sleep(0.8) 
                            
                            csv_y_position = btn_center_y + 40
                            page.mouse.move(btn_center_x, csv_y_position)
                            time.sleep(0.2)
                            
                            with page.expect_download(timeout=15000) as download_info:
                                page.mouse.click(btn_center_x, csv_y_position)
                                
                            download = download_info.value
                            download.save_as(file_path)
                            
                            try:
                                df_day = pd.read_csv(file_path, encoding="utf-8-sig")
                            except Exception:
                                df_day = pd.read_csv(file_path, encoding="shift_jis")
                                
                            df_day.insert(0, "Date", target_date_str)
                            df_day.insert(1, "account_id", current_acc_id)
                            df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                            
                            print(f"      => [THANH CONG] Da tai xong data SKYFLAG ngay {target_date_str}")
                            download_success = True
                        else:
                            print("      ! Khong lay duoc toa do vi tri cua nut Export.")
                            
                    except Exception as e:
                        print(f"      ! Loi phat sinh khi thao tac tai file ngay {target_date_str}: {e}")

                    if is_first_day:
                        is_first_day = False
                        if not download_success:
                            print(f"   [DỪNG KHẨN CẤP] Ngay dau tien cua SKYFLAG bi loi xuat file. Dung luong va bo qua cac ngay sau!")
                            return
                    time.sleep(1.0)
            except Exception as system_err:
                print(f"   [DỪNG KHẨN CẤP] Gap loi nghiem trong trong luong cua SKYFLAG: {system_err}")
                return
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho SKYFLAG Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY RIÊNG CHO ZucksDSP (RÚT GỌN TUÂN THỦ LUỒNG CHÍNH)
# ========================================================
def run_zucksdsp(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"    -> [ZucksDSP Row {row_stt} | ID {current_acc_id}] Luong chinh yeu cau quet / tai moi du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    api_key = ""
    for col_key in account_info.index:
        if str(col_key).strip().lower() in ["api_key", "api key"]:
            api_key = str(account_info[col_key]).strip()
            break
    if not api_key or api_key == "None":
        api_key = str(account_info.get('api_key', account_info.get('params', ''))).strip()

    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    if not target_dates:
        print(f"      => Không có ngày cần tải cho ZucksDSP.")
        return
        
    start_date = min(target_dates)
    end_date = max(target_dates)
    from_str = start_date.strftime("%Y-%m-%d")
    to_str = end_date.strftime("%Y-%m-%d")
    t_start_str = start_date.strftime("%Y%m%d")
    t_end_str = end_date.strftime("%Y%m%d")
    
    master_file_name = f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_start_str}_{t_end_str}_master.xlsx"
    master_file_path = os.path.join(processing_dir, master_file_name)

    print(f"      => Tien hanh tai du lieu tu {from_str} den {to_str}. Mo trinh duyet...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        page = context.new_page()
        try:
            print("   + [Web ZucksDSP] Dang truy cap trang dang nhap...")
            page.goto("https://zucks-dsp.jp.auth0.com/u/login", wait_until="networkidle")
            time.sleep(1.5)
            
            print("   + [Web ZucksDSP] Dien Email login...")
            page.locator("input[type='email'], input[name='username'], #username").fill(login_id)
            time.sleep(0.5)
            page.locator("button:has-text('続ける'), button[type='submit']").first.click(force=True)
            
            page.wait_for_load_state("networkidle")
            time.sleep(1.5)
            
            print("   + [Web ZucksDSP] Dien mat khau...")
            page.locator("input[type='password'], input[name='password'], #password").fill(login_pw)
            time.sleep(0.5)
            page.locator("button:has-text('続ける'), button[type='submit']").first.click(force=True)
            
            page.wait_for_load_state("networkidle")
            time.sleep(4.0)
            
            data_url = f"https://zucks.dsp.cartajp.com/reporting/{current_acc_id}/general/report/{api_key}&timezone=Asia%2FTokyo&from={from_str}&to={to_str}"
            print(f"   + [Web ZucksDSP] Dieu huong den link data: {data_url}")
            page.goto(data_url, wait_until="commit")
            
            print("   + [Web ZucksDSP] Dang doi bang du lieu tai xong (Quet chu '合計値')...")
            page.wait_for_selector("//*[contains(text(), '合計値')]", timeout=60000)
            time.sleep(3.0)
            
            print("   + [Web ZucksDSP] Click mo menu tai file (Icon Download)...")
            download_trigger = page.locator(".mat-mdc-button-base, button").filter(has_text=re.compile(r"file_download|ダウンロード", re.IGNORECASE))
            if download_trigger.count() == 0:
                download_trigger = page.locator("button:has(span.mat-icon), .mat-icon:has-text('download')").first
            
            download_trigger.click(force=True)
            time.sleep(1.0)
            
            print("   + [Web ZucksDSP] Click chon 'Excelファイルをダウンロード'...")
            excel_btn = page.locator(".mat-mdc-menu-item-text, .mat-menu-item").get_by_text("Excelファイルをダウンロード", exact=True)
            
            with page.expect_download(timeout=30000) as download_info:
                excel_btn.click(force=True)
                
            download = download_info.value
            download.save_as(master_file_path)
            print(f"      => [THANH CONG] Da tai xong file Master Excel cho ZucksDSP.")
            
            if str(account_info.get('summary_type', 'daily')).strip().lower() == "summary":
                print("   + [Web ZucksDSP] Convert file Excel Master sang file CSV chu ky...")
                df_master = pd.read_excel(master_file_path)
                df_master.insert(0, "account_id", current_acc_id)
                
                final_summary_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_start_str}_{t_end_str}.csv")
                df_master.to_csv(final_summary_path, index=False, encoding="utf-8-sig")
            else:
                print("   + [Web ZucksDSP] Dang phân tách dữ liệu Excel thành từng ngày CSV...")
                df_master = pd.read_excel(master_file_path)
                
                date_col = None
                for col in df_master.columns:
                    if str(col).strip() in ['日', 'Date', '日付']:
                        date_col = col
                        break
                if date_col is None:
                    raise ValueError("Không tìm thấy cột chứa dữ liệu Ngày ('日') trong file xuất của ZucksDSP!")

                for single_date in target_dates:
                    d_str_dash = single_date.strftime("%Y-%m-%d")
                    d_str_clean = single_date.strftime("%Y%m%d")
                    file_daily_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{d_str_clean}.csv")
                    df_daily = df_master[df_master[date_col].astype(str).str.contains(d_str_dash)]
                    df_daily.insert(0, "Date", d_str_clean)
                    df_daily.insert(1, "account_id", current_acc_id)
                    df_daily.to_csv(file_daily_path, index=False, encoding="utf-8-sig")
                print(f"      => [XỬ LÝ OK] Đã phân tách thành công các file daily CSV.")
        except Exception as system_err:
            print(f"   [DỪNG KHẨN CẤP] Gap loi nghiem trong trong luong cua ZucksDSP: {system_err}")
            raise system_err
        finally:
            browser.close()


# ========================================================
# MODULE CHẠY RIÊNG CHO JAMPP (TỐI GIẢN - TINH GỌN THEO LUỒNG CHÍNH)
# ========================================================
def run_jampp(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [JAMPP Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    summary_type = str(account_info.get('summary_type', 'daily')).strip().lower()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    print(f"      => Kich hoat trinh duyet theo lenh luong chinh...")
    with sync_playwright() as p:
        jampp_profile_dir = os.path.join(COOKIE_DIR, "jampp_profile")
        context = p.chromium.launch_persistent_context(
            user_data_dir=jampp_profile_dir,
            headless=False,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage"
            ]
        )
        page = context.new_page()
        
        try:
            page.goto("https://jampp.microadtech.com/app/#/users/login", wait_until="networkidle")
            time.sleep(2.0)
            
            if page.locator("input[placeholder='Email']").count() > 0:
                page.get_by_placeholder("Email").fill(login_id)
                time.sleep(0.5)
                page.get_by_placeholder("Password").fill(login_pw)
                time.sleep(1.0)
                
                login_btn = page.locator("button:has-text('Login'), .btn").first
                login_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(4.0)
            
            campaign_url = f"https://jampp.microadtech.com/app/#/campaigns/{current_acc_id}"
            print(f"   + [Web JAMPP] Dang vao trang danh sach tong: {campaign_url}")
            page.goto(campaign_url, wait_until="networkidle")
            time.sleep(4.0) 
            
            print("   + [Web JAMPP] Dang tim va chon channel 'Jampp' trong bang...")
            try:
                page.wait_for_selector(f"a[href*='channel/jampp'], table, .table", timeout=15000)
            except Exception:
                pass

            channel_link = page.locator(f"a[href='#/campaigns/{current_acc_id}/channel/jampp']").first
            if channel_link.count() == 0 or not channel_link.is_visible():
                channel_link = page.locator("a").filter(has_text=re.compile(r"^jampp$", re.IGNORECASE)).first
            if channel_link.count() == 0 or not channel_link.is_visible():
                channel_link = page.locator("table td a, .table a").get_by_text("Jampp", exact=True).first

            if channel_link.count() > 0 and channel_link.is_visible():
                channel_link.scroll_into_view_if_needed()
                time.sleep(0.5)
                channel_link.click(force=True)
            else:
                page.evaluate(f"""() => {{
                    let el = document.querySelector("a[href*='channel/jampp']");
                    if(el) {{ el.click(); }}
                    else {{
                        let links = Array.from(document.querySelectorAll('a'));
                        let target = links.find(a => a.textContent.trim().toLowerCase() === 'jampp');
                        if(target) target.click();
                    }}
                }}""")
            
            page.wait_for_load_state("networkidle")
            time.sleep(4.0) 
            
            # 🌟 XỬ LÝ CHIA LUỒNG ĐIỀU KIỆN TẢI TẠI ĐÂY
            if summary_type == "summary":
                s_date, e_date = target_dates[0], target_dates[-1]
                start_date_str = s_date.strftime("%Y%m%d")
                end_date_str = e_date.strftime("%Y%m%d")
                file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{start_date_str}_{end_date_str}.csv")
                execution_loops = [("summary", s_date, e_date, file_path, f"{start_date_str}-{end_date_str}")]
            else:
                # Nếu là luồng Daily (mặc định ban đầu của bạn)
                execution_loops = []
                for d in target_dates:
                    d_str = d.strftime("%Y%m%d")
                    f_p = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{d_str}.csv")
                    execution_loops.append((d_str, d, d, f_p, d_str))

            for label_str, s_date, e_date, file_path, display_date in execution_loops:
                formatted_start = s_date.strftime("%m/%d/%Y")
                formatted_end = e_date.strftime("%m/%d/%Y")
                print(f"   + [Web JAMPP] Thiet lap thoi gian tu: {formatted_start} den {formatted_end}")
                
                date_picker_trigger = page.locator(".dropdown-toggle, [ui-daterangepicker], #reportrange").filter(has_text=re.compile(r"2026")).first
                if date_picker_trigger.count() == 0:
                    date_picker_trigger = page.locator("//div[contains(@class,'pull-right')]//span[contains(text(),'20')]").first
                    
                date_picker_trigger.scroll_into_view_if_needed()
                date_picker_trigger.click(force=True)
                time.sleep(1.0)
                
                custom_range_btn = page.locator("li").filter(has_text="Custom Range").filter(visible=True).first
                custom_range_btn.click(force=True)
                time.sleep(1.0)
                
                start_input = page.locator("input[name='daterangepicker_start']").filter(visible=True).first
                start_input.click(force=True)
                start_input.fill("") 
                start_input.fill(formatted_start)
                start_input.press("Enter")
                time.sleep(0.5)
                
                end_input = page.locator("input[name='daterangepicker_end']").filter(visible=True).first
                end_input.click(force=True)
                end_input.fill("") 
                end_input.fill(formatted_end)
                end_input.press("Enter")
                time.sleep(0.5)
                
                apply_btn = page.locator(".daterangepicker .applyBtn, button.btn-success, button").filter(has_text="Apply").filter(visible=True).first
                apply_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(4.0)

                creative_title = page.locator("//*[contains(text(), 'Jampp - Creative Report')]").first
                creative_title.scroll_into_view_if_needed()
                time.sleep(1.0)
                
                select_columns_btn = page.locator("button[ng-click='toggleDropdown()']").filter(has_text="Select Columns").nth(4)
                if select_columns_btn.count() == 0:
                    select_columns_btn = page.locator("button[ng-click='toggleDropdown()']").filter(has_text="Select Columns").last
                    
                select_columns_btn.click(force=True)
                time.sleep(1.0)
                
                campaign_checkbox = page.locator("a[role='menuitem']").filter(has_text="Campaign Name").locator("input.checkboxInput").filter(visible=True).first
                if campaign_checkbox.count() > 0 and not campaign_checkbox.is_checked():
                    campaign_checkbox.click(force=True)
                installs_checkbox = page.locator("a[role='menuitem']").filter(has_text="Installs").locator("input.checkboxInput").filter(visible=True).first
                if installs_checkbox.count() > 0 and not installs_checkbox.is_checked():
                    installs_checkbox.click(force=True)
                    
                select_columns_btn.click(force=True)
                time.sleep(0.5)
                
                export_btn = page.locator("button[data-toggle='dropdown']").filter(has_text="Export").nth(4)
                if export_btn.count() == 0:
                    export_btn = page.locator("button[data-toggle='dropdown']").filter(has_text="Export").last
                export_btn.scroll_into_view_if_needed()
                export_btn.click(force=True)
                time.sleep(1.0)
                
                csv_menu_item = page.locator(".dropdown-menu a, .dropdown-menu li").get_by_text("CSV", exact=True).filter(visible=True).first
                if csv_menu_item.count() == 0:
                    csv_menu_item = page.get_by_text("CSV", exact=True).filter(visible=True).first
                
                try:
                    with page.expect_download(timeout=25000) as download_info:
                        csv_menu_item.click(force=True)
                    download = download_info.value
                    download.save_as(file_path)
                    
                    try:
                        df_day = pd.read_csv(file_path, encoding="utf-8-sig")
                    except Exception:
                        df_day = pd.read_csv(file_path, encoding="shift_jis")
                        
                    if "Campaign Name" in df_day.columns:
                        df_day = df_day[df_day["Campaign Name"].astype(str).str.strip().str.lower() != "total"]
                    
                    df_day.insert(0, "Date", display_date)
                    df_day.insert(1, "account_id", current_acc_id)
                    df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                except Exception as download_err:
                    print(f"      ! Loi tai file CSV: {download_err}")
                    raise download_err
                    
                page.evaluate("window.scrollTo(0, 0);")
                time.sleep(1.5)
        except Exception as system_err:
            print(f"   [DỪNG KHẨN CẤP] Gap loi trong luong cua JAMPP: {system_err}")
            raise system_err
        finally:
            context.close()


# ========================================================
# MODULE CHẠY CHO PPDSP (RADIO BUTTON UTF-8)
# ========================================================
def run_ppdsp(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [{current_media} Row {row_stt} | ID {current_acc_id}] Bat dau tai bao cao thang...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    secret_key = ""
    for col_key in account_info.index:
        if "API KEY" in str(col_key).upper() or "API_KEY" in str(col_key).upper():
            secret_key = str(account_info[col_key]).strip()
            break
    if not secret_key or secret_key == "None" or secret_key == "":
        raise ValueError(f"Khong tim thay OTP Secret Key cho PPDSP Row {row_stt}.")

    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    months_need_action = get_missing_monthly_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if months_need_action:
        print(f"      => Thieu du lieu {len(months_need_action)} thang. Mo trinh duyet...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                login_target_url = f"https://dashboard.prismpartner.smt.docomo.ne.jp/agencies/3/advertisers/{current_acc_id}/report"
                page.goto(login_target_url, wait_until="networkidle")
                time.sleep(3.0) 
                
                form_context = page
                if len(page.frames) > 1:
                    for frame in page.frames:
                        if frame.get_by_placeholder("name@host.com").count() > 0:
                            form_context = frame
                            break
                
                email_input = form_context.get_by_placeholder("name@host.com").filter(visible=True).first
                email_input.wait_for(state="visible", timeout=10000)
                email_input.fill(login_id)
                email_input.dispatch_event("change")
                
                pw_input = form_context.get_by_placeholder("Password").filter(visible=True).first
                pw_input.fill(login_pw)
                pw_input.dispatch_event("blur")
                time.sleep(1)
                
                signin_btn = form_context.locator("button:has-text('Sign in'), input[type='submit'][value*='Sign']").filter(visible=True).first
                signin_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(4)
                
                time_remaining = 30 - (int(time.time()) % 30)
                if time_remaining < 6:
                    time.sleep(time_remaining + 1)
                
                totp = pyotp.TOTP(secret_key)
                current_otp_code = totp.now()

                otp_context = page
                if len(page.frames) > 1:
                    for frame in page.frames:
                        if frame.locator("input").count() > 0:
                            otp_context = frame
                            break

                otp_input = otp_context.locator("input[autocomplete*='one-time-code'], input[placeholder*='code'], input[type='text']").filter(visible=True).first
                otp_input.wait_for(state="visible", timeout=15000)
                otp_input.press_sequentially(current_otp_code, delay=120)
                time.sleep(1)
                
                otp_btn = otp_context.locator("button:has-text('Sign in'), button[type='submit'], button").filter(visible=True).first
                otp_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                
                try:
                    page.wait_for_url(lambda url: "auth" not in url, timeout=20000)
                except Exception:
                    raise RuntimeError("Xac thuc OTP hoac thong tin dang nhap that bai.")

                if page.url != login_target_url:
                    page.goto(login_target_url, wait_until="networkidle")
                    time.sleep(4)

                for year_month in months_need_action:
                    promo_dropdown = page.locator("div.v-select:has-text('プロモーションタイプ'), div:has-text('プロモーションタイプ')").last
                    promo_dropdown.click(force=True)
                    time.sleep(0.8)
                    page.locator("div.v-menu__content .v-list-item:has-text('デジタル広告'), [role='option']:has-text('デジタル広告')").first.click(force=True)
                    time.sleep(1.0)
                    
                    date_dropdown = page.locator("div.v-select:has-text('対象期間'), div:has-text('対象期間')").last
                    date_dropdown.click(force=True)
                    time.sleep(1.5)
                    
                    split_ym = year_month.split('-')
                    target_month_label = f"{int(split_ym[1])}月" 
                    target_month_btn = page.locator("button").filter(has_text=target_month_label).filter(visible=True).first
                    target_month_btn.click(force=True)
                    time.sleep(2.0) 

                    utf8_radio = page.locator("label:has-text('UTF-8'), div:has-text('UTF-8'), input[value*='UTF-8']").filter(visible=True).first
                    utf8_radio.click(force=True)
                    time.sleep(1.0)

                    csv_download_btn = page.locator("button:has-text('CSVダウンロード')").first
                    zip_temp_path = os.path.join(processing_dir, f"temp_{year_month}.zip")
                    with page.expect_download(timeout=20000) as download_info:
                        csv_download_btn.click(force=True)
                    
                    download = download_info.value
                    download.save_as(zip_temp_path)

                    extracted_dir = os.path.join(processing_dir, f"extracted_{year_month}")
                    os.makedirs(extracted_dir, exist_ok=True)
                    with zipfile.ZipFile(zip_temp_path, 'r') as zip_ref:
                        zip_ref.extractall(extracted_dir)
                    
                    source_csv_path = os.path.join(extracted_dir, f"実績レポート_DSP_{year_month}.csv")
                    if not os.path.exists(source_csv_path):
                        all_files = glob.glob(os.path.join(extracted_dir, "*DSP*"))
                        if all_files: source_csv_path = all_files[0]

                    dest_processed_csv = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{year_month.replace('-', '')}.csv")
                    if os.path.exists(source_csv_path):
                        try:
                            df_month = pd.read_csv(source_csv_path, encoding="utf-8-sig")
                        except Exception:
                            df_month = pd.read_csv(source_csv_path, encoding="shift_jis")

                        date_col_name = ""
                        for c in df_month.columns:
                            if any(k in str(c) for k in ["日", "date", "Date", "期間"]):
                                date_col_name = c
                                break
                        if date_col_name:
                            valid_str_dates = [d.strftime("%Y-%m-%d") for d in target_dates] + [d.strftime("%Y/%m/%d") for d in target_dates]
                            df_month = df_month[df_month[date_col_name].astype(str).apply(lambda x: any(v in x for valid_part in valid_str_dates for v in [valid_part]))]

                        df_month.insert(0, "account_id", current_acc_id)
                        df_month.to_csv(dest_processed_csv, index=False, encoding="utf-8-sig")
                    else:
                        with open(dest_processed_csv, "w", encoding="utf-8") as f: f.write("0")

                    try:
                        os.remove(zip_temp_path)
                        shutil.rmtree(extracted_dir)
                    except Exception: pass
            except Exception as system_err:
                print(f"   [DỪNG KHẨN CẤP] Loi luong PPDSP: {system_err}")
                raise system_err
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Du lieu thang da ton tai san cho PPDSP Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY CHO CATSYS
# ========================================================
def run_catsys(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [Catsys Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        print(f"      => Mo trinh duyet cao du lieu cho {len(dates_to_download)} ngay...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                page.goto("https://wb-p01.catsys.jp/front/login/", wait_until="networkidle")
                time.sleep(1.5)
                
                page.locator("input[type='text'], input[name*='id'], input[id*='id']").first.fill(login_id)
                time.sleep(0.5)
                page.locator("input[type='password'], input[name*='pass'], input[id*='pass']").first.fill(login_pw)
                time.sleep(1)
                
                login_btn = page.locator("//button[contains(@class,'login') or contains(text(),'ログイン')] | //input[@type='submit' or @value='ログイン'] | //form//button").first
                if login_btn.is_visible():
                    login_btn.click(force=True)
                else:
                    page.locator("input[type='password']").first.press("Enter")
                    
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                
                page.goto("https://wb-p01.catsys.jp/admin/profilecontentmedia/list", wait_until="networkidle")
                time.sleep(2)
                
                is_first_day = True
                for single_date in dates_to_download:
                    target_date_str = single_date.strftime("%Y%m%d")
                    file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
                    
                    day_val = str(single_date.day)
                    print(f"   + [Web Catsys] Thiet lap lich cho ngay: {single_date.strftime('%Y/%m/%d')}")
                    
                    container_input = page.locator("div.form-group:has-text('期間選択') input, input.daterangepicker-field, input[id*='date']").first
                    if container_input.count() == 0 or not container_input.is_visible():
                        container_input = page.locator("input[value*='/']").first
                        
                    container_input.click(force=True)
                    time.sleep(1.5)
                    
                    target_calendar_block = page.locator("div.drp-calendar.right, div.calendar.right").first
                    if target_calendar_block.count() == 0 or not target_calendar_block.is_visible():
                        target_calendar_block = page.locator("div.drp-calendar, div.calendar").nth(1)

                    day_cell = target_calendar_block.locator(f"td:not(.off):has-text('{day_val}')").filter(has_text=re.compile(rf"^{day_val}$")).first
                    
                    if day_cell.is_visible():
                        day_cell.click(force=True)
                        time.sleep(0.5)
                        day_cell.click(force=True)
                        time.sleep(0.5)
                    else:
                        print("      ! Khong tim thay ngay, dung JS ghi de text.")
                        formatted_date_slash = single_date.strftime("%Y/%m/%d")
                        page.evaluate(f"arguments[0].value = '{formatted_date_slash} - {formatted_date_slash}';", container_input.element_handle())
                        container_input.dispatch_event("change")
                        time.sleep(0.5)
                    
                    search_btn = page.locator("//button[contains(text(),'検索') or contains(text(),'適用')] | //input[@value='検索'] | //*[text()='検索']").first
                    if search_btn.is_visible():
                        search_btn.click(force=True)
                        page.wait_for_load_state("networkidle")
                        time.sleep(2.5)
                    
                    if "データはありません" in page.content() or page.locator("text=データはありません。").is_visible():
                        print(f"      [SKIP] Ngay {target_date_str} thong bao khong co du lieu. Tao file gia.")
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write("0")
                        is_first_day = False
                        continue
                    
                    download_success = False
                    try:
                        csv_gen_locator = page.locator("//button[contains(text(),'CSV生成')] | //input[@value='CSV生成'] | //*[contains(text(), 'CSV生成')]").first
                        with page.expect_download(timeout=10000) as download_info:
                            csv_gen_locator.click(force=True)
                            
                        download = download_info.value
                        download.save_as(file_path)
                        
                        try:
                            df_day = pd.read_csv(file_path, encoding="utf-8-sig")
                        except Exception:
                            df_day = pd.read_csv(file_path, encoding="shift_jis")
                            
                        df_day.insert(0, "Date", target_date_str)
                        df_day.insert(1, "account_id", current_acc_id)
                        df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                        print(f"      => [THANH CONG] Da tai xong data ngay {target_date_str}")
                        download_success = True
                    except Exception as e:
                        print(f"      ! Khong co du lieu hoac loi phan sinh tai ngay {target_date_str}: {e}")
                    
                    if is_first_day:
                        is_first_day = False
                        if not download_success:
                            print(f"   [DỪNG KHẨN CẤP] Ngay dau tien cua Catsys bi loi. Dung luong!")
                            return
                    time.sleep(1.0)
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho Catsys Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY CHO ACCESSTRADE (API)
# ========================================================
def run_accesstrade(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [Accesstrade Row {row_stt} | ID {current_acc_id}] Bat dau goi API lay du lieu...")
    api_key = ""
    for col_key in account_info.index:
        if "API KEY" in str(col_key).upper() or "API_KEY" in str(col_key).upper():
            api_key = str(account_info[col_key]).strip()
            break
    if not api_key or api_key == "None":
        raise ValueError("Khong co API KEY cho Accesstrade.")

    report_type = str(account_info.get('reportType', 'SiteReport')).strip()

    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        is_first_day = True
        for single_date in dates_to_download:
            target_date_str = single_date.strftime("%Y%m%d")
            file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
            
            api_url = f"https://m-api.accesstrade.net/matv3_api/external_api/v2/program/{current_acc_id}/report/{report_type}"
            params = {"targetFrom": single_date.strftime("%Y-%m-%d"), "targetTo": single_date.strftime("%Y-%m-%d")}
            headers = {"Authenticate-Key": api_key, "Accept": "application/json"}
            
            try:
                response = requests.get(api_url, params=params, headers=headers, timeout=15)
                if response.status_code == 200:
                    json_result = response.json()
                    data_array = json_result if isinstance(json_result, list) else (json_result.get('data') or json_result.get('items') or None)
                    if not data_array:
                        with open(file_path, "w", encoding="utf-8") as f: f.write("0")
                    else:
                        df_day = pd.DataFrame(data_array)
                        df_day.insert(0, "Date", target_date_str)
                        df_day.insert(1, "account_id", current_acc_id)
                        df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                else:
                    raise RuntimeError(f"API tra ve status code {response.status_code}")
            except Exception as e:
                if is_first_day: raise e
            is_first_day = False
            time.sleep(0.5)
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho Accesstrade Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY CHO JANET
# ========================================================
def run_janet(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [JANet Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = []
    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        print(f"      => Thieu hoac can tai lai {len(dates_to_download)} ngay. Mo trinh duyet...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                page.goto("https://j-a-net.jp/client/promotion/report/report_media?SID=868f395b2099751724bf8b19393d3588a4cdd7b970823ad62025509029575853", wait_until="networkidle")
                time.sleep(1.5)
                
                page.locator("input[type='text']").first.fill(login_id)
                time.sleep(0.5)
                page.locator("input[type='password']").first.fill(login_pw)
                time.sleep(1)
                
                login_btn = page.locator("//input[@value='ログイン' or @type='submit' or @type='image'][contains(@src, 'login') or contains(@class, 'login')] | //*[text()='ログイン']").first
                if login_btn.is_visible():
                    login_btn.click(force=True)
                else:
                    page.locator("input[type='password'] >> xpath=../following-sibling::input").first.click(force=True)
                    
                page.wait_for_load_state("networkidle")
                time.sleep(3.0)
                
                print("   + [Web JANet] Click chon nut 'プロモーションTOPへ'...")
                promo_top_btn = page.locator("img[alt='プロモーションTOPへ'], img[src*='promo_totop.gif']").first
                promo_top_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(2.0)
                
                print("   + [Web JANet] Click chon nut 'レポート'...")
                report_btn = page.locator("a").filter(has_text=re.compile(r"^レポート$")).first
                if report_btn.count() == 0:
                    report_btn = page.locator("//a[text()='レポート' or contains(@href, 'report_month')]").first
                report_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(2.0)
                
                print("   + [Web JANet] Click chon nut 'サイト別レポート'...")
                media_report_btn = page.locator("a").filter(has_text=re.compile(r"^サイト別レポート$")).first
                if media_report_btn.count() == 0:
                    media_report_btn = page.locator("//a[text()='サイト別レポート' or contains(@href, 'report_media')]").first
                media_report_btn.click(force=True)
                page.wait_for_load_state("networkidle")
                time.sleep(2.0)
                
                print(f"   + [Web JANet] Dropdown lua chon Account ID phu hop: {current_acc_id}")
                select_element = page.locator("select[name='promotion_id']").first
                select_element.wait_for(state="visible", timeout=10000)
                
                select_element.select_option(value=current_acc_id)
                page.wait_for_load_state("networkidle")
                time.sleep(2.5) 
                
                is_first_day = True
                for single_date in dates_to_download:
                    target_date_str = single_date.strftime("%Y%m%d")
                    formatted_date_str = single_date.strftime("%Y-%m-%d")
                    file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
                    
                    js_date_script = f"""
                        let textInputs = Array.from(document.querySelectorAll("input[type='text']")).filter(i => i.placeholder.includes('20') || i.value.includes('20') || i.name.includes('date'));
                        const inputs = document.querySelectorAll("input[type='text'][value*='20'], input[name*='date'], input[id*='date']");
                        if(textInputs.length >= 2) {{
                            textInputs[0].value = '{formatted_date_str}';
                            textInputs[1].value = '{formatted_date_str}';
                            textInputs[0].dispatchEvent(new Event('change', {{ bubbles: true }}));
                            textInputs[1].dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }} else if(inputs.length >= 2) {{
                            inputs[0].value = '{formatted_date_str}';
                            inputs[1].value = '{formatted_date_str}';
                            inputs[0].dispatchEvent(new Event('change', {{ bubbles: true }}));
                            inputs[1].dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """
                    page.evaluate(js_date_script)
                    time.sleep(1.5)
                    
                    search_btn = page.locator("input[type='image'][src*='search'], input[value='検索'], button:has-text('検索'), .btn:has-text('検索'), [class*='search']").first
                    if search_btn.is_visible():
                        search_btn.click(force=True)
                    else:
                        page.locator("//input[@value='検索'] | //button[contains(text(),'検索')] | //*[text()='検索']").first.click(force=True)
                        
                    page.wait_for_load_state("networkidle")
                    time.sleep(2.5)
                    
                    if "該当するデータが見つかりません" in page.content() or page.locator("text=該当する data が見つかりません").is_visible():
                        print(f"      [SKIP] Ngay {target_date_str} hien thi thong bao trong. Tao file gia bong.")
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write("0")
                        is_first_day = False
                        continue
                    
                    download_success = False
                    try:
                        csv_locator = page.locator("//a[contains(@href, 'csv') or contains(@onclick, 'csv')] | //input[contains(@value, 'CSV') or contains(@src, 'csv')] | //*[contains(text(), 'CSVダウンロード')] | //img[contains(@src, 'csv') or contains(@alt, 'CSV')]/..").first
                        
                        with page.expect_download(timeout=10000) as download_info:
                            csv_locator.click(force=True)
                            
                        download = download_info.value
                        download.save_as(file_path)
                        
                        try:
                            df_day = pd.read_csv(file_path, encoding="shift_jis")
                        except Exception:
                            df_day = pd.read_csv(file_path, encoding="utf-8-sig")
                            
                        df_day.insert(0, "Date", target_date_str)
                        df_day.insert(1, "account_id", current_acc_id)
                        df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                        print(f"      => [THANH CONG] Da tai xong data ngay {target_date_str}")
                        download_success = True
                    except Exception as e:
                        print(f"       ! Khong co data hoac loi tai ngay {target_date_str}: {e}")
                    
                    if is_first_day:
                        is_first_day = False
                        if not download_success:
                            print(f"   [DỪNG KHẨN CẤP] Ngay dau tien cua JANet (Row {row_stt}) bi loi. Dung tien trinh luong!")
                            return
                            
                    time.sleep(1.0)
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho JANet Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY CHO GMOAFFITOWN
# ========================================================
def run_gmoaffitown(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [GMO Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        print(f"      => Thieu hoac can tai lai {len(dates_to_download)} ngay. Mo trinh duyet...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                page.goto("https://affitown.jp/", wait_until="networkidle")
                page.locator("input[type='text'], input[name*='id']").nth(1).fill(login_id)
                page.locator("input[type='password']").nth(1).fill(login_pw)
                time.sleep(1)
                page.locator("input[type='submit'][value*='広告主様'].btn-primary").click(delay=150)
                page.wait_for_load_state("networkidle")
                
                page.goto("https://affitown.jp/clientadmin/report/ad/list/", wait_until="networkidle")
                
                is_first_day = True
                for single_date in dates_to_download:
                    target_date_str = single_date.strftime("%Y%m%d")
                    file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
                    
                    search_box = page.locator("div.panel-wrapper")
                    if not search_box.is_visible() or "in" not in search_box.get_attribute("class", default=""):
                        page.locator("text=検索条件").click()
                        time.sleep(0.5)
                    
                    formatted_date_str = single_date.strftime("%Y年%m月%d日") + " - " + single_date.strftime("%Y年%m月%d日")
                    js_script = f"const input = document.querySelector('#searchDate'); if(input) {{ input.value = '{formatted_date_str}'; input.dispatchEvent(new Event('change', {{ bubbles: true }})); }}"
                    page.evaluate(js_script)
                    time.sleep(2.5) 
                    
                    page.locator("button:has-text('検索する')").first.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(2.5) 
                    
                    if "該当する data が見つかりません" in page.content() or "データがありません" in page.content():
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write("0")
                        is_first_day = False
                        continue
                    
                    download_success = False
                    try:
                        with page.expect_download(timeout=5000) as download_info:
                            page.locator("text=CSVダウンロード").first.click()
                        download = download_info.value
                        download.save_as(file_path)
                        
                        try:
                            df_day = pd.read_csv(file_path, encoding="utf-8-sig")
                        except Exception:
                            df_day = pd.read_csv(file_path, encoding="shift_jis")
                        df_day.insert(0, "Date", target_date_str)
                        df_day.insert(1, "account_id", current_acc_id)
                        df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                        download_success = True
                    except Exception:
                        pass
                    
                    if is_first_day:
                        is_first_day = False
                        if not download_success:
                            print(f"   [DỪNG KHẨN CẤP] Ngay dau tien cua GMO bi loi. Dung luong!")
                            return
                    time.sleep(1.0)
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho GMO Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# MODULE CHẠY CHO ADTRACK
# ========================================================
def run_adtrack(account_info, cmd_info, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id):
    print(f"   -> [ADTRACK Row {row_stt} | ID {current_acc_id}] Bat dau chay luong du lieu...")
    login_id = str(account_info.get('login_id', '')).strip()
    login_pw = str(account_info.get('login_pw', '')).strip()
    
    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    redownload_val = str(cmd_info.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    dates_to_download = get_missing_daily_files(target_dates, processing_dir, cad_prefix, row_stt, current_media, current_acc_id, force_redownload)

    if dates_to_download:
        print(f"      => Thieu hoac can tai lai {len(dates_to_download)} ngay. Mo trinh duyet...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            page = context.new_page()
            try:
                page.goto("https://admin.ad-track.jp/report/index.php", wait_until="networkidle")
                time.sleep(1.5)
                page.locator("input[type='text']").nth(1).fill(login_id)
                page.locator("input[type='password']").nth(1).fill(login_pw)
                time.sleep(1)
                page.locator("input[type='submit'][value='ログイン']").nth(1).click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                
                page.goto("https://admin.ad-track.jp/report/index.php?a=report/site/SiteReport", wait_until="networkidle")
                time.sleep(2)
                
                is_first_day = True
                for single_date in dates_to_download:
                    target_date_str = single_date.strftime("%Y%m%d")
                    formatted_date_str = single_date.strftime("%Y-%m-%d")
                    file_path = os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{target_date_str}.csv")
                    
                    js_date_script = f"""
                        const inputs = document.querySelectorAll("input[type='text'][value*='20']");
                        if(inputs.length >= 2) {{
                            inputs[0].value = '{formatted_date_str}';
                            inputs[1].value = '{formatted_date_str}';
                            inputs[0].dispatchEvent(new Event('change', {{ bubbles: true }}));
                            inputs[1].dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """
                    page.evaluate(js_date_script)
                    time.sleep(1.5)
                    
                    inputs = page.locator("input[type='submit'][value='検索']")
                    if inputs.count() > 0:
                        inputs.first.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(2.5)
                    
                    if "該当する data が見vòng tẩm" in page.content():
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write("0")
                        is_first_day = False
                        continue
                    
                    download_success = False
                    try:
                        with page.expect_download(timeout=8000) as download_info:
                            page.locator("input[alt*='CSV'], [title*='CSV']").first.click()
                        download = download_info.value
                        download.save_as(file_path)
                        
                        try:
                            df_day = pd.read_csv(file_path, encoding="shift_jis", skiprows=2)
                        except Exception:
                            df_day = pd.read_csv(file_path, encoding="utf-8-sig", skiprows=2)
                        
                        if "メディア名" in df_day.columns:
                            df_day = df_day[df_day["メディア名"].astype(str).str.strip() != "[total]"]

                        df_day.insert(0, "Date", target_date_str)
                        df_day.insert(1, "account_id", current_acc_id)
                        df_day.to_csv(file_path, index=False, encoding="utf-8-sig")
                        download_success = True
                    except Exception:
                        pass
                    
                    if is_first_day:
                        is_first_day = False
                        if not download_success:
                            print(f"   [DỪNG KHẨN CẤP] Ngay dau tien cua ADTRACK bi loi. Dung luong!")
                            return
                    time.sleep(1.0)
            finally:
                browser.close()
    else:
        print(f"      => [OFFLINE OK] Thu muc processing da co du file cho ADTRACK Row {row_stt} | ID {current_acc_id}.")


# ========================================================
# HÀM TRỢ GIÚP GỘP FILE VÀ KIỂM TRA ĐIỀU KIỆN CHẶN FILE THIẾU
# ========================================================
def gộp_và_lưu_dữ_liệu(all_expected_files, target_dates, cad_prefix, media_name, account_id, today_dir, final_data_dir, row_stt, summary_type="daily", processing_dir=None):
    start_date_file = target_dates[0].strftime("%Y%m%d")
    end_date_file = target_dates[-1].strftime("%Y%m%d")
    summary_file_name = f"{cad_prefix}_{row_stt}_{media_name}_{account_id}_{start_date_file}_{end_date_file}.csv"
    
    os.makedirs(final_data_dir, exist_ok=True)
    os.makedirs(today_dir, exist_ok=True)
    
    path_final_root = os.path.join(final_data_dir, summary_file_name)
    path_final_today = os.path.join(today_dir, summary_file_name)

    if summary_type == "summary":
        list_summary_dfs = []
        for f_path in all_expected_files:
            if os.path.exists(f_path):
                list_summary_dfs.append(pd.read_csv(f_path))
            else:
                master_file_name = f"{cad_prefix}_{row_stt}_{media_name}_{account_id}_{start_date_file}_{end_date_file}_master.xlsx"
                master_file_path = os.path.join(processing_dir, master_file_name)
                if os.path.exists(master_file_path):
                    try:
                        df_master = pd.read_excel(master_file_path)
                        if "account_id" not in df_master.columns:
                            df_master.insert(0, "account_id", account_id)
                        df_master.to_csv(f_path, index=False, encoding="utf-8-sig")
                        list_summary_dfs.append(pd.read_csv(f_path))
                    except Exception as e:
                        print(f"      ! Loi convert du lieu master Excel cho ID {account_id}: {e}")
                        
        if list_summary_dfs:
            combined_summary_df = pd.concat(list_summary_dfs, ignore_index=True)
            combined_summary_df.to_csv(path_final_root, index=False, encoding="utf-8-sig")
            combined_summary_df.to_csv(path_final_today, index=False, encoding="utf-8-sig")
            print(f"      + [OK] Loai [Summary] Da dong bo va ghi file ket qua: {path_final_root}")
            return True
        return False

    list_df = []
    for f_path in all_expected_files:
        if os.path.exists(f_path):
            try:
                with open(f_path, "r", encoding="utf-8", errors="ignore") as f_chk:
                    content_preview = f_chk.read(5).strip()
                if media_name != "SKYFLAG" and (content_preview == "0" or content_preview == ""):
                    continue
                list_df.append(pd.read_csv(f_path))
            except Exception: 
                pass
    
    if list_df:
        combined_df = pd.concat(list_df, ignore_index=True)
        combined_df.to_csv(path_final_root, index=False, encoding="utf-8-sig")
        combined_df.to_csv(path_final_today, index=False, encoding="utf-8-sig")
        print(f"      + [OK] Da luu file gop ({summary_type}) tai folder ket qua: {path_final_today}")
    else:
        with open(path_final_root, "w", encoding="utf-8-sig") as f_empty_root: f_empty_root.write("0")
        with open(path_final_today, "w", encoding="utf-8-sig") as f_empty_today: f_empty_today.write("0")
        print(f"      ! [THÔNG BÁO] Tai khoan ID {account_id} khong co du lieu.")
    return True


# ========================================================
# BỌC XỬ LÝ CHO TỪNG DÒNG (FIX: TÁCH BIỆT REPORT KEY VÀ SỐ LƯỢNG)
# ========================================================
def execute_single_account_row(account_info, cmd_row, target_dates, media_col):
    current_media = str(account_info[media_col]).strip()
    row_stt = account_info['row_stt']
    
    summary_type = str(account_info.get('summary_type', 'daily')).strip().lower()
    if summary_type not in ["daily", "monthly", "summary"]:
        summary_type = "daily"

    raw_acc_id_str = str(account_info.get('account_id', '')).strip()
    if not raw_acc_id_str or raw_acc_id_str.lower() in ["none", "nan"]:
        print(f"❌ [SKIP ACC TRỐNG] STT Dong {row_stt} - Media {current_media} khong co account_id hieu luc!")
        return

    account_ids_list = [id_item.strip() for id_item in raw_acc_id_str.split(',') if id_item.strip()]
    
    clean_account_ids = []
    for single_id in account_ids_list:
        if single_id.endswith(".0"):
            single_id = single_id[:-2]
        else:
            try:
                if '.' in single_id:
                    single_id = str(int(float(single_id)))
            except Exception:
                pass
        clean_account_ids.append(single_id.strip())

    cad_prefix = "CAD"
    for col_key in account_info.index:
        if str(col_key).strip().upper() == "CAD":
            cad_prefix = str(account_info[col_key]).strip()
            break

    data_folder_name = str(cmd_row.get('data_folder', 'AF_DATA')).strip()
    final_data_dir = os.path.join(DESKTOP_PATH, data_folder_name)
    media_dir = os.path.join(final_data_dir, current_media)        
    today_str = get_current_date_str()
    today_dir = os.path.join(media_dir, today_str)                 
    processing_dir = os.path.join(today_dir, "processing")         
    os.makedirs(processing_dir, exist_ok=True)
    
    redownload_val = str(cmd_row.get('redownload', '')).strip()
    force_redownload = (redownload_val.upper() == "ON" or redownload_val.lower() == current_media.lower())

    today_nodash = datetime.now().strftime("%Y%m%d")
    yesterday_nodash = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")

    for current_acc_id in clean_account_ids:
        sub_acc_expected_files = []
        
        if summary_type == "summary":
            t_start_str = target_dates[0].strftime("%Y%m%d")
            t_end_str = target_dates[-1].strftime("%Y%m%d")
            sub_acc_expected_files.append(os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_start_str}_{t_end_str}.csv"))
            
        elif summary_type == "monthly" or current_media == "PPDSP":
            months_list = sorted(list(set([d.strftime("%Y-%m") for d in target_dates])))
            for year_month in months_list:
                ym_clean = year_month.replace('-', '')
                sub_acc_expected_files.append(os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{ym_clean}.csv"))
                
        else:
            for single_date in target_dates:
                t_str = single_date.strftime("%Y%m%d")
                sub_acc_expected_files.append(os.path.join(processing_dir, f"{cad_prefix}_{row_stt}_{current_media}_{current_acc_id}_{t_str}.csv"))

        acc_files_exist_count = sum(1 for f_p in sub_acc_expected_files if os.path.exists(f_p))
        need_download_sub_acc = (acc_files_exist_count < len(sub_acc_expected_files)) or force_redownload
        
        if not need_download_sub_acc and summary_type != "summary":
            for f in sub_acc_expected_files:
                if today_nodash in f or yesterday_nodash in f:
                    need_download_sub_acc = True; break

        try:
            if need_download_sub_acc:
                print(f"      => [TIẾN TRÌNH] Tài khoản ID {current_acc_id} trống file hoặc cần làm mới. Kích hoạt trình duyệt...")
                media_lower = current_media.strip().lower()
                func_name = f"run_{media_lower}"
                
                target_function = None
                for global_key, global_val in list(globals().items()):
                    if global_key.lower() == func_name:
                        target_function = global_val; break
                
                if target_function and callable(target_function):
                    target_function(account_info, cmd_row, current_media, target_dates, processing_dir, today_dir, final_data_dir, row_stt, current_acc_id)
                else:
                    raise NotImplementedError(f"Hệ thống không tìm thấy hàm con dạng '{func_name}'!")
            else:
                print(f"      => [OFFLINE OK] ID {current_acc_id} của Media {current_media} đã tồn tại đủ file thành phần.")

            is_merged = gộp_và_lưu_dữ_liệu(sub_acc_expected_files, target_dates, cad_prefix, current_media, current_acc_id, today_dir, final_data_dir, row_stt, summary_type, processing_dir)
            
            start_date_str = target_dates[0].strftime("%Y%m%d")
            end_date_str = target_dates[-1].strftime("%Y%m%d")
            
            subject_val = "Bao cao dinh ky"
            for col_key in account_info.index:
                if "件名" in str(col_key) or "kenmei" in str(col_key).lower():
                    subject_val = str(account_info[col_key]).strip(); break
            
            mail_send_val = ""
            for col_key in account_info.index:
                if "mail_send" in str(col_key).lower():
                    mail_send_val = str(account_info[col_key]).strip().upper(); break

            # 🎯 FIX LỖI TÊN HIỂN THỊ TRÊN BÁO CÁO CUỐI: Chỉ lưu Key dạng Media thuần chủng phục vụ in ấn
            if is_merged:
                with mail_lock:  # <- THÊM DÒNG NÀY TẠI ĐÂY
                    # Tất cả các dòng bên dưới này được thụt lề vào thêm 1 lượt (4 dấu cách)
                    MEDIA_STATUS_REPORT[f"{current_media}_{current_acc_id}"] = {"dl_status": "DL_OK", "mail_send": mail_send_val, "subject": subject_val, "base_media": current_media}
                    
                    PENDING_MAIL_ROWS.append({
                        "media": current_media,
                        "cad": cad_prefix,
                        "row_stt": row_stt,
                        "account_id": current_acc_id,
                        "start_date": start_date_str,
                        "end_date": end_date_str,
                        "subject": subject_val,
                        "today_dir": today_dir,
                        "account_info": account_info,
                        "cmd_row": cmd_row,
                        "dl_status": "OK"
                    })
            else:
                raise RuntimeError(f"Huy tien trinh do thieu file thanh phan kết quả của ID {current_acc_id}.")
                
        except Exception as single_id_err:
            err_clean_msg = str(single_id_err).replace('\n', ' ')
            err_msg = f"STT Dong {row_stt} - Media {current_media} - ID {current_acc_id} gap su co: {err_clean_msg}"
            print(f"❌ [SKIP ACC LỖI] {err_msg}")
            write_error_log(today_dir, err_msg)
            
            subject_val = "Bao cao dinh ky"
            for col_key in account_info.index:
                if "件名" in str(col_key) or "kenmei" in str(col_key).lower():
                    subject_val = str(account_info[col_key]).strip(); break
            mail_send_val = ""
            for col_key in account_info.index:
                if "mail_send" in str(col_key).lower():
                    mail_send_val = str(account_info[col_key]).strip().upper(); break
                    
            MEDIA_STATUS_REPORT[f"{current_media}_{current_acc_id}"] = {"dl_status": "DL_ERROR", "mail_send": mail_send_val, "subject": subject_val, "base_media": current_media}
            mail_status_err = "ERROR" if mail_send_val == "ON" else "SKIP"
            create_execution_log(today_dir, account_info, cmd_row, current_media, "ERROR", mail_status_err, f"Loi ID {current_acc_id}: {err_clean_msg}", current_acc_id)


# ========================================================
# HÀM GỬI MAIL TRỰC TIẾP QUA GMAIL DOANH NGHIỆP CỦA CÔNG TY
# ========================================================
def auto_send_report_emails_company_gmail(cmd_row):
    mail_execution_results = {}
    resend_val = ""
    for col_key in cmd_row.index:
        if "resend" in str(col_key).lower():
            resend_val = str(cmd_row[col_key]).strip().upper(); break
    
    for unique_report_key, info in MEDIA_STATUS_REPORT.items():
        if info["mail_send"] == "ON":
            if info["dl_status"] == "DL_OK":
                mail_execution_results[unique_report_key] = {"status": "MAIL_OK", "count": 0}
            else:
                mail_execution_results[unique_report_key] = {"status": "MAIL_ERROR", "count": 0}
        else:
            mail_execution_results[unique_report_key] = {"status": "MAIL_SKIP", "count": 0}

    if PENDING_MAIL_ROWS:
        print("\n" + "="*50)
        print("[GMAIL] Đang gửi mail đến report-rodeta-vietnam@ca-adv.co.jp...")
        print("="*50)
        
        SENDER_EMAIL = "tran_thuuyen@ca-adv.co.jp" 
        SENDER_PASSWORD = "mstt fddm aodb tntf" 

        groups = {}
        for item in PENDING_MAIL_ROWS:
            subject = item["subject"]
            groups.setdefault(subject, []).append(item)

        for subject, items_list in groups.items():
            attachments_to_send = []

            for item in items_list:
                m_name = item['media']
                current_acc_id = item['account_id']
                unique_report_key = f"{m_name}_{current_acc_id}"
                
                current_mail_send_val = ""
                for col_key in item["account_info"].index:
                    if "mail_send" in str(col_key).lower():
                        current_mail_send_val = str(item["account_info"][col_key]).strip().upper(); break

                if current_mail_send_val != "ON":
                    print(f"   -> [SKIP EMAIL] Media [{m_name}] ID [{current_acc_id}] co mail_send='{current_mail_send_val}', khong gui email.")
                    if unique_report_key in mail_execution_results:
                        mail_execution_results[unique_report_key]["status"] = "MAIL_SKIP"
                    create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "SKIP", "Bo qua vi mail_send khong bat ON.", current_acc_id)
                    continue

                is_already_sent = False
                should_force_resend = (resend_val == "ON" or m_name.lower() in [r.strip().lower() for m in resend_val.split(',') for r in m.split(' ') if r.strip()])
                
                if not should_force_resend:
                    existing_logs = glob.glob(os.path.join(item["today_dir"], f"log_*_{item['cad']}_{m_name}.csv"))
                    for log_p in existing_logs:
                        try:
                            df_chk_log = pd.read_csv(log_p)
                            # Kiểm tra xem file log có đủ cả 3 cột cần thiết hay không
                            if "Mail" in df_chk_log.columns and "AF_account_id" in df_chk_log.columns and "件名" in df_chk_log.columns:
                                # Lọc các dòng khớp ĐỒNG THỜI account_id VÀ subject (件名)
                                sub_df = df_chk_log[
                                    (df_chk_log["AF_account_id"].astype(str) == str(current_acc_id)) & 
                                    (df_chk_log["件名"].astype(str).str.strip() == str(subject).strip())
                                ]
                                if not sub_df.empty:
                                    # Lấy trạng thái Mail của dòng dữ liệu khớp gần nhất (values[-1])
                                    m_status_logged = str(sub_df["Mail"].values[-1]).strip().upper()
                                    if m_status_logged in ["OK", "SENT"]:
                                        is_already_sent = True
                                        break
                        except Exception: 
                            pass
                
                if is_already_sent:
                    print(f"   -> [CHẶN TRÙNG] Media [{m_name}] ID [{current_acc_id}] đã gửi thành công trước đó hôm nay.")
                    if unique_report_key in mail_execution_results:
                        mail_execution_results[unique_report_key]["status"] = "MAIL_SENT"
                    create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "SENT", "Da chan gui trung.", current_acc_id)
                    continue

                target_file_name = f"{item['cad']}_{item['row_stt']}_{item['media']}_{current_acc_id}_{item['start_date']}_{item['end_date']}.csv"
                full_file_path = os.path.join(item["today_dir"], target_file_name)

                if os.path.exists(full_file_path):
                    try:
                        with open(full_file_path, "r", encoding="utf-8", errors="ignore") as f_chk:
                            content = f_chk.read(5).strip()
                        if content == "0" or content == "":
                            print(f"   ! File gop {target_file_name} rong, bo qua file nay.")
                            mail_execution_results[unique_report_key]["status"] = "MAIL_ERROR"
                            create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "ERROR", f"File gop {target_file_name} rong.", current_acc_id)
                            continue 
                    except Exception:
                        mail_execution_results[unique_report_key]["status"] = "MAIL_ERROR"
                        create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "ERROR", "Khong the mo doc file gop.", current_acc_id)
                        continue

                    if full_file_path not in attachments_to_send:
                        attachments_to_send.append(full_file_path)
                else:
                    print(f"   ! Khong tim thay file {target_file_name}, bo qua file nay.")
                    mail_execution_results[unique_report_key]["status"] = "MAIL_ERROR"
                    create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "ERROR", f"Khong tim thay file {target_file_name}", current_acc_id)
                    continue

            # Đoạn này đính kèm và gửi email cho tất cả các file thu thập được trong hàng
            if attachments_to_send:
                print(f"   -> Đang tiến hành gửi mail cho nhóm các file hợp lệ với {len(attachments_to_send)} file gộp...")
                try:
                    msg = MIMEMultipart()
                    msg['From'] = SENDER_EMAIL
                    msg['To'] = "report-rodeta-vietnam@ca-adv.co.jp"
                    msg['Subject'] = subject
                    
                    body_content = f"""
                    <html>
                      <body>
                        <p>UynUyn đã gửi mail tự động ngày <b>{datetime.now().strftime('%d/%m/%Y')}</b>.</p>
                        <p>Danh sách file đính kèm:</p>
                        <ul>
                    """
                    for f_p in attachments_to_send:
                        body_content += f"<li>{os.path.basename(f_p)}</li>"
                    body_content += "</ul><br><p>Mãi iu!</p></body></html>"
                    
                    msg.attach(MIMEText(body_content, 'html'))
                    
                    for f_p in attachments_to_send:
                        with open(f_p, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            
                            f_name = os.path.basename(f_p)
                            part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', f_name))
                            msg.attach(part)
                    
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    server.login(SENDER_EMAIL, SENDER_PASSWORD)
                    server.sendmail(SENDER_EMAIL, msg['To'], msg.as_string())
                    server.quit()
                    
                    print(f"      => [OK] Đã gửi thành công email tiêu đề '{subject}'!")
                    
                    for item in items_list:
                        m_name = item['media']
                        current_acc_id = item['account_id']
                        unique_report_key = f"{m_name}_{current_acc_id}"
                        target_f = f"{item['cad']}_{item['row_stt']}_{item['media']}_{current_acc_id}_{item['start_date']}_{item['end_date']}.csv"
                        if any(target_f in f for f in attachments_to_send):
                            if unique_report_key in mail_execution_results:
                                mail_execution_results[unique_report_key]["count"] = len(attachments_to_send)
                                mail_execution_results[unique_report_key]["status"] = "MAIL_OK"
                            create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "OK", "Gui email thanh cong ruc ro.", current_acc_id)

                except Exception as e:
                    print(f"      ! [LỖI GỬI MAIL] Lỗi kết nối server: {e}")
                    for item in items_list:
                        m_name = item['media']
                        current_acc_id = item['account_id']
                        unique_report_key = f"{m_name}_{current_acc_id}"
                        target_f = f"{item['cad']}_{item['row_stt']}_{item['media']}_{current_acc_id}_{item['start_date']}_{item['end_date']}.csv"
                        if any(target_f in f for f in attachments_to_send):
                            if unique_report_key in mail_execution_results:
                                mail_execution_results[unique_report_key]["status"] = "MAIL_ERROR"
                            create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "ERROR", f"Loi he thong gui email: {str(e)}", current_acc_id)

    for unique_report_key, info in MEDIA_STATUS_REPORT.items():
        if info["mail_send"] != "ON" and info["dl_status"] == "DL_OK":
            for item in PENDING_MAIL_ROWS:
                if f"{item['media']}_{item['account_id']}" == unique_report_key:
                    create_execution_log(item["today_dir"], item["account_info"], item["cmd_row"], item["media"], "OK", "SKIP", "Khong on gui mail. Ket thuc ghi log chu ky.", item['account_id'])

    print("\n" + "="*100)
    print("=== BÁO CÁO TỔNG HỢP TÌNH TRẠNG THỰC THI TOÀN HỆ THỐNG ===")
    print("="*100)
    
    RED = "\033[91m"
    RESET = "\033[0m"
    
    for unique_report_key, info in MEDIA_STATUS_REPORT.items():
        dl_status = info["dl_status"]
        subject = info["subject"]
        base_media = info["base_media"]
        
        mail_res = mail_execution_results.get(unique_report_key, {"status": "MAIL_SKIP", "count": 0})
        mail_status = mail_res["status"]
        file_count = mail_res["count"]
        
        if dl_status == "DL_OK" and mail_status == "MAIL_OK":
            print(f"・OK - {base_media}（DL_OK）（MAIL_OK）- 件名:{subject} - file: {file_count}")
        elif dl_status == "DL_OK" and mail_status == "MAIL_SENT":
            print(f"・OK - {base_media}（DL_OK）（MAIL_SENT）- 件名:{subject}")
        elif dl_status == "DL_OK" and mail_status == "MAIL_SKIP":
            print(f"・OK - {base_media}（DL_OK）（MAIL_SKIP）")
        elif dl_status == "DL_OK" and mail_status == "MAIL_ERROR":
            print(f"{RED}・ERROR - {base_media}（DL_OK）（MAIL_ERROR）件名:{subject}{RESET}")
        elif dl_status == "DL_ERROR" and mail_status == "MAIL_ERROR":
            print(f"{RED}・ERROR - {base_media}（DL_ERROR）（MAIL_ERROR）{RESET}")
        elif dl_status == "DL_ERROR" and mail_status == "MAIL_SKIP":
            print(f"{RED}・ERROR - {base_media}（DL_ERROR）（MAIL_SKIP）{RESET}")
            
    print("="*100 + "\n")


# ========================================================
# TIẾN TRÌNH ĐIỀU HƯỚNG CHÍNH VÀ QUẢN LÝ ĐA NHIỆM MAX_WORKER
# ========================================================
def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"-> Khong tim thay file Excel tai: {EXCEL_FILE}")
        return None

    try:
        df_cmd, df_af = read_excel_safely(EXCEL_FILE)
    except Exception as e:
        print(f"-> Loi khi doc file Excel: {e}")
        return None

    if df_cmd.empty:
        print("-> Sheet AF_Cmd khong co du lieu.")
        return None

    for index_cmd, cmd_row in df_cmd.iterrows():
        cmd_status = str(cmd_row.get('status', '')).strip()
        if cmd_status != "ON":
            continue

        print_ascii_cat()

        redownload_val = str(cmd_row.get('redownload', '')) if not pd.isna(cmd_row.get('redownload', '')) else ""
        media_val = str(cmd_row.get('media', '')) if not pd.isna(cmd_row.get('media', '')) else ""
        time_val = str(cmd_row.get('time', '')) if not pd.isna(cmd_row.get('time', '')) else ""

        print(f"redownload: {redownload_val.strip()}")
        print(f"media: {media_val.strip()}")
        print(f"time: {time_val.strip()}")

        media_raw = media_val.strip()
        if not media_raw or media_raw == "None" or media_raw == "nan":
            continue

        time_range_value = time_val.strip() if time_val.strip() != "" else "This month"
        media_list = [m.strip() for m in media_raw.split(',') if m.strip()]
        target_dates = calculate_date_range(time_range_value)

        media_col = next((col for col in df_af.columns if str(col).lower().strip() == 'media'), '')
        status_col = next((col for col in df_af.columns if str(col).lower().strip() == 'status'), '')

        if not media_col:
            return None

        tasks_to_execute = []
        for idx, row in df_af.iterrows():
            af_status = str(row[status_col]).strip().upper() if status_col else ""
            af_media_val = str(row[media_col]).strip()

            if af_status != "ON" or not af_media_val or af_media_val in ["None", "nan"]:
                continue

            sub_medias = [sm.strip() for sm in af_media_val.split(',') if sm.strip()]
            for current_sm in sub_medias:
                if current_sm in media_list:
                    task_copy = row.copy()
                    task_copy[media_col] = current_sm
                    tasks_to_execute.append(task_copy)

        if not tasks_to_execute:
            continue

        try:
            max_workers_val = cmd_row.get('max_worker', 1)
            max_workers = int(float(max_workers_val)) if not pd.isna(max_workers_val) and max_workers_val != "" else 1
        except Exception:
            max_workers = 1
        if max_workers < 1: max_workers = 1

        print(f"\n=== media: {len(tasks_to_execute)}, max_worker: {max_workers} ===")
        print(f"=== Time: {target_dates[0]} den {target_dates[-1]} ===")
        
        PENDING_MAIL_ROWS.clear()
        MEDIA_STATUS_REPORT.clear()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [
                executor.submit(execute_single_account_row, task_row, cmd_row, target_dates, media_col)
                for task_row in tasks_to_execute
            ]
            for future in futures:
                future.result() # Đợi gom đủ tất cả các thread chạy xong

        #  ĐÚNG: Đặt ngoài block 'with' để chắc chắn tải đủ all account mới up mail 1 lần
        auto_send_report_emails_company_gmail(cmd_row)
            
    return True

# ========================================================
# LUỒNG CHẠY CHÍNH
# ========================================================
if __name__ == "__main__":
    main()
    
    print("[Enter]...")
    try:
        input()
    except Exception: pass
    os.system("taskkill /f /im cmd.exe")
    sys.exit(0)