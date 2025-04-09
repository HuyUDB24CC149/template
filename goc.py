import requests
import time
import random
import re
import threading
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

# ---------------------------------------
# Hàm cập nhật proxy cho profile
# ---------------------------------------
def update_proxy(profile_id, raw_proxy):
    """
    Cập nhật proxy cho profile trước khi mở.
    Gửi POST request với raw_proxy (prefix là "http://").
    Nếu trả về "Profile not found" thì log profile_id ra file profileloi.txt.
    """
    update_url = f"http://127.0.0.1:19995/api/v3/profiles/update/{profile_id}"
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    data = {"raw_proxy": f"http://{raw_proxy}"}
    try:
        r = requests.post(update_url, headers=headers, json=data)
        r.raise_for_status()
        response_json = r.json()
        if response_json.get("success") and response_json.get("message") == "Update profile success":
            print(f"Proxy updated successfully for profile {profile_id}.")
            return True
        elif (not response_json.get("success")) and response_json.get("message") == "Profile not found":
            print(f"Update failed. Profile not found: {profile_id}")
            with open("profileloi.txt", "a") as f:
                f.write(str(profile_id) + "\n")
            return False
        else:
            print(f"Unexpected response when updating proxy for profile {profile_id}: {response_json}")
            return False
    except Exception as e:
        print(f"Exception updating proxy for profile {profile_id}: {e}")
        return False

# ---------------------------------------
# Hàm đổi IP
# ---------------------------------------
def change_ip(change_ip_url):
    """
    Gọi GET request tới API đổi IP.
    Nếu API trả về lỗi "Vui lòng chờ sau X giây" thì đợi (X+2) giây rồi retry.
    """
    while True:
        try:
            r = requests.get(change_ip_url)
            try:
                response_json = r.json()
            except Exception:
                print("Lỗi khi parse JSON từ API đổi IP, retry sau 10 giây...")
                time.sleep(10)
                continue

            if response_json.get("status") == "success":
                print("Đổi IP thành công.")
                return True
            elif response_json.get("status") == "error":
                error_msg = response_json.get("error", "")
                m = re.search(r"Vui lòng chờ sau (\d+) giây", error_msg)
                if m:
                    wait_seconds = int(m.group(1)) + 2
                    print(f"Lỗi đổi IP: {error_msg}. Đợi {wait_seconds} giây rồi retry lại...")
                    time.sleep(wait_seconds)
                    continue
                else:
                    print(f"Phản hồi lỗi không xác định từ API đổi IP: {response_json}")
                    return False
            else:
                print(f"Phản hồi không mong đợi từ API đổi IP: {response_json}")
                return False
        except Exception as e:
            print(f"Exception while changing IP: {e}. Đợi 10 giây rồi retry lại...")
            time.sleep(10)
            continue

# ---------------------------------------
# Đọc file proxy.txt
# ---------------------------------------
try:
    with open("proxy.txt", "r") as f:
        proxy_lines = [line.strip() for line in f.readlines()]
    if len(proxy_lines) != 4:
        raise ValueError("Expected 4 proxy lines in proxy.txt.")

    proxies = []
    for line in proxy_lines:
        proxy_parts = line.split("|")
        if len(proxy_parts) != 2:
            raise ValueError(f"Invalid format in proxy.txt.\nExpected: proxy|changeIP_url. Line: {line}")
        raw_proxy = proxy_parts[0]
        change_ip_url = proxy_parts[1]
        proxies.append({"raw_proxy": raw_proxy, "change_ip_url": change_ip_url})
        print(f"Loaded proxy: {raw_proxy}, Change IP URL: {change_ip_url}")

except Exception as e:
    print(f"Error reading proxy.txt: {e}")
    proxies = []

# ---------------------------------------
# Đọc file Excel profiles.xlsx
#  - Cột A: Profile ID
#  - Cột F: Kết quả (có thể chứa "Thành Công")
# Chỉ lấy profile nếu cột F là "Thành Công"
# ---------------------------------------
workbook = load_workbook('profiles.xlsx')
worksheet = workbook.active

red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

profiles = []
for i, row in enumerate(worksheet.iter_rows(min_row=2, max_col=6, values_only=False), start=2):
    profile_id = row[0].value  # Cột A
    result_value = worksheet.cell(row=i, column=6).value  # Cột F

    # Chỉ xử lý nếu cột F là "Thành Công"
    if result_value and str(result_value).strip() == "Thành Công":
        profiles.append({
            "id": profile_id,
            "row": i
        })
    else:
        print(f"Skipping profile {profile_id} (row {i}) vì cột F không phải 'Thành Công'.")

# Biến toàn cục để xử lý đa luồng
profiles_lock = threading.Lock()
profile_index = 0  # index profile hiện tại

def process_profile(thread_id, proxy_data, window_pos):
    global profile_index

    raw_proxy = proxy_data["raw_proxy"]
    change_ip_url = proxy_data["change_ip_url"]

    while True:
        with profiles_lock:
            if profile_index >= len(profiles):
                print(f"Thread {thread_id}: No more profiles to process.")
                break
            current_profile_index = profile_index
            profile_index += 1

        profile = profiles[current_profile_index]
        profile_id = profile["id"]
        row_number = profile["row"]

        print(f"Thread {thread_id}: Processing profile {profile_id} (Row {row_number})")

        # 1) Cập nhật proxy
        if raw_proxy and change_ip_url:
            if not update_proxy(profile_id, raw_proxy):
                print(f"Thread {thread_id}: Skipping profile {profile_id} due to proxy update failure.")
                with profiles_lock:
                    workbook.save('profiles.xlsx')
                continue

            # 2) Đổi IP
            if not change_ip(change_ip_url):
                print(f"Thread {thread_id}: Sự cố đổi IP, không loại bỏ profile {profile_id}, sẽ retry.")
        else:
            print(f"Thread {thread_id}: Proxy info not available; skipping proxy update.")

        # 3) Mở profile qua API
        start_url = f"http://127.0.0.1:19995/api/v3/profiles/start/{profile_id}?addination_args=--lang%3Dvi&win_pos={window_pos}&win_size=1800%2C1080&win_scale=0.35"
        print(f"Thread {thread_id}: Opening profile via URL: {start_url}")
        try:
            start_resp = requests.get(start_url)
            start_resp.raise_for_status()
        except Exception as e:
            print(f"Thread {thread_id}: Error opening profile {profile_id}: {e}")
            continue

        start_data = start_resp.json()
        if not start_data.get("success"):
            print(f"Thread {thread_id}: Failed to open profile {profile_id}: {start_data}")
            continue

        driver_path = start_data.get("data", {}).get("driver_path")
        remote_debugging_address = start_data.get("data", {}).get("remote_debugging_address")
        browser_location = start_data.get("data", {}).get("browser_location")

        if not driver_path or not remote_debugging_address:
            print(f"Thread {thread_id}: Missing driver_path or remote_debugging_address, skipping profile.")
            continue

        # 4) Khởi tạo Selenium
        options = Options()
        # options.binary_location = browser_location # nếu cần
        options.add_experimental_option("debuggerAddress", remote_debugging_address)
        service = Service(executable_path=driver_path)
        try:
            driver = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            print(f"Thread {thread_id}: Error initializing webdriver for profile {profile_id}: {e}")
            continue

        # ---------------------------
        # 5) Truy cập trang tiktok
        # ---------------------------
        wait = WebDriverWait(driver, 30)
        tiktok_url = "https://www.tiktok.com"
        print(f"Thread {thread_id}: Navigating to {tiktok_url}")
        try:
            driver.get(tiktok_url)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            print(f"Thread {thread_id}: TikTok page loaded.")
        except Exception as e:
            print(f"Thread {thread_id}: Error loading TikTok page for profile {profile_id}: {e}")
            close_url = f"http://127.0.0.1:19995/api/v3/profiles/close/{profile_id}"
            try:
                requests.get(close_url)
                print(f"Thread {thread_id}: Profile {profile_id} closed successfully after error.")
            except Exception as err_close:
                print(f"Thread {thread_id}: Error closing profile {profile_id}: {err_close}")
            continue

        # (Tại đây, bạn có thể thêm code upload video nếu muốn)
        
        # 6) Do đã lọc sẵn cột F = "Thành Công", có thể không cần ghi Excel nữa

        # 7) Đóng profile qua API
        close_url = f"http://127.0.0.1:19995/api/v3/profiles/close/{profile_id}"
        print(f"Thread {thread_id}: Closing profile with URL: {close_url}")
        try:
            requests.get(close_url)
            print(f"Thread {thread_id}: Profile {profile_id} closed successfully.")
        except Exception as e:
            print(f"Thread {thread_id}: Error closing profile {profile_id}: {e}")
        time.sleep(1)

# ---------------------------------------
# Main
# ---------------------------------------
if __name__ == "__main__":
    # Kiểm tra đủ 4 proxy chưa
    if len(proxies) < 4:
        print("Không đủ 4 proxy để chạy 4 luồng, vui lòng kiểm tra lại proxy.txt")
        exit(1)

    # Tạo 4 luồng, mỗi luồng xài 1 proxy
    thread1 = threading.Thread(target=process_profile, args=(1, proxies[0], "0,0"))
    thread2 = threading.Thread(target=process_profile, args=(2, proxies[1], "1800,0"))
    thread3 = threading.Thread(target=process_profile, args=(3, proxies[2], "3600,0"))
    thread4 = threading.Thread(target=process_profile, args=(4, proxies[3], "0,1080"))

    thread1.start()
    thread2.start()
    thread3.start()
    thread4.start()

    thread1.join()
    thread2.join()
    thread3.join()
    thread4.join()

    print("\nCompleted processing all profiles.")
