from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os

# Khởi tạo Chrome
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

wait = WebDriverWait(driver, 10)
driver.get("https://tracuuhoadon.fpt.com")

# Input mẫu
data_input = [
    {"mst": "0304244470", "ma_tra_cuu": "r08e17y79g"},
    {"mst": "0304244471", "ma_tra_cuu": "r46jvxmvxg"},
    {"mst": "0304308445", "ma_tra_cuu": "rzmwy1yo4g"}
]

# Tạo thư mục lưu ảnh
os.makedirs("hoadon_screenshots", exist_ok=True)

def get_text_by_label(label_text):
    try:
        xpath = f"//label[contains(text(), '{label_text}')]/following-sibling::*"
        return driver.find_element(By.XPATH, xpath).text.strip()
    except:
        return ""

all_data = []

for i, item in enumerate(data_input, 1):
    try:
        driver.get("https://tracuuhoadon.fpt.com")
        time.sleep(2)

        # Điền mã
        driver.find_element(By.ID, "mst").clear()
        driver.find_element(By.ID, "mst").send_keys(item['mst'])

        driver.find_element(By.ID, "code").clear()
        driver.find_element(By.ID, "code").send_keys(item['ma_tra_cuu'])

        driver.find_element(By.ID, "btnTraCuu").click()

        # Đợi hiển thị hóa đơn
        wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Số hóa đơn')]")))

        # Trích xuất thông tin
        data_row = {
            "Số hóa đơn": get_text_by_label("Số hóa đơn"),
            "Đơn vị bán hàng": get_text_by_label("Đơn vị bán hàng"),
            "Mã số thuế bán": get_text_by_label("Mã số thuế"),
            "Địa chỉ bán": get_text_by_label("Địa chỉ"),
            "Số tài khoản bán": get_text_by_label("Số tài khoản"),
            "Họ tên người mua hàng": get_text_by_label("Người mua"),
            "Địa chỉ mua": get_text_by_label("Địa chỉ",),  # địa chỉ người mua
            "Mã số thuế mua": get_text_by_label("Mã số thuế"),  # người mua
            "Mã tra cứu": item["ma_tra_cuu"],
            "Mã số thuế input": item["mst"],
        }

        # Lưu screenshot
        screenshot_path = f"hoadon_screenshots/hoadon_{i}_{item['mst']}.png"
        driver.save_screenshot(screenshot_path)
        data_row["Ảnh hóa đơn"] = screenshot_path

        all_data.append(data_row)

    except Exception as e:
        print(f"Lỗi với {item}: {e}")
        all_data.append({
            "Số hóa đơn": "",
            "Đơn vị bán hàng": "",
            "Mã số thuế bán": "",
            "Địa chỉ bán": "",
            "Số tài khoản bán": "",
            "Họ tên người mua hàng": "",
            "Địa chỉ mua": "",
            "Mã số thuế mua": "",
            "Mã tra cứu": item["ma_tra_cuu"],
            "Mã số thuế input": item["mst"],
            "Ảnh hóa đơn": "Không chụp được"
        })

driver.quit()

# Xuất ra Excel
df = pd.DataFrame(all_data)
df.to_excel("ket_qua_hoa_don.xlsx", index=False)
print("✅ Đã lưu file: ket_qua_hoa_don.xlsx và ảnh hóa đơn trong thư mục.")
