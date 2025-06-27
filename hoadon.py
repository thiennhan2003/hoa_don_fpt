from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ===== Cấu hình Email =====
EMAIL_SEND = "nhanvai2003@gmail.com"
EMAIL_PASSWORD = "oeiqjzazidtaezmi"
EMAIL_RECEIVE = "nhanvai2003@gmail.com"

# ===== Thư mục lưu hóa đơn =====
folder_download = os.path.join(os.getcwd(), "hoadon_pdf")
os.makedirs(folder_download, exist_ok=True)

# ===== Cấu hình Chrome =====
options = Options()
options.add_argument("--start-maximized")
prefs = {
    "download.default_directory": folder_download,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
}
options.add_experimental_option("prefs", prefs)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# ===== Đọc mã tra cứu =====
df_input = pd.read_excel("matracuu.xlsx")
df_input.columns = df_input.columns.str.strip()
ma_tra_cuu_list = df_input["Mã tra cứu"].dropna().tolist()

# ===== Xử lý từng mã =====
for ma in ma_tra_cuu_list:
    try:
        print(f"\n===> Đang xử lý mã: {ma}")
        driver.get("https://www.meinvoice.vn/tra-cuu")
        time.sleep(1)

        # Điền mã tra cứu
        input_box = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[placeholder='Nhập mã tra cứu hóa đơn']")))
        input_box.clear()
        input_box.send_keys(ma)

        # Bấm nút tra cứu
        btn_tra_cuu = wait.until(EC.element_to_be_clickable((By.ID, "btnSearchInvoice")))
        btn_tra_cuu.click()
        print("Đã bấm Tra cứu, chờ kết quả...")

        time.sleep(3)

        # Bấm nút Tải hóa đơn
        btn_tai_hoa_don = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Tải hóa đơn')]")))
        driver.execute_script("arguments[0].click();", btn_tai_hoa_don)
        print("Đã bấm nút Tải hóa đơn, chờ menu xổ xuống...")

        time.sleep(0.5)  # Chờ menu hiện

        # Dùng ActionChains click tọa độ lệch xuống menu PDF
        actions = ActionChains(driver)
        actions.move_to_element_with_offset(btn_tai_hoa_don, 10, 40).click().perform()
        print("Đã click Tải hóa đơn dạng PDF, chờ tải file...")

        # Theo dõi file tải về
        before_files = set(os.listdir(folder_download))
        timeout = 30
        start_time = time.time()
        file_pdf = None

        while time.time() - start_time < timeout:
            current_files = set(os.listdir(folder_download))
            new_files = current_files - before_files
            for f in new_files:
                if f.lower().endswith(".pdf"):
                    file_pdf = os.path.join(folder_download, f)
                    break
            if file_pdf:
                break
            time.sleep(1)

        if file_pdf and os.path.isfile(file_pdf):
            print(f"Đã tải file: {file_pdf}")

            msg = MIMEMultipart()
            msg['From'] = EMAIL_SEND
            msg['To'] = EMAIL_RECEIVE
            msg['Subject'] = f"Hóa đơn từ mã tra cứu {ma}"

            part = MIMEBase('application', 'octet-stream')
            with open(file_pdf, "rb") as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_pdf)}"')
            msg.attach(part)

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(EMAIL_SEND, EMAIL_PASSWORD)
                server.send_message(msg)

            print("Đã gửi mail thành công.")
        else:
            print("Tải thất bại hoặc không tìm thấy file PDF.")
    except Exception as e:
        print(f"Lỗi với mã {ma}: {e}")

# ===== Kết thúc =====
driver.quit()
print("Hoàn tất tất cả tác vụ.")
