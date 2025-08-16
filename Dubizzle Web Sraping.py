import time
import re
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

SAVE_PATH = r"C:\Users\NEXT\Desktop\dubizzle_data__.xlsx"
DETAILS_WAIT = 15
OPEN_SLEEP = 1.2
LIST_WAIT = 20

def build_driver():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    options.add_argument("--lang=ar")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/127.0.0.1 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def safe_text(driver, by, selector):
    try:
        el = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((by, selector))
        )
        return el.text.strip()
    except:
        try:
            el = driver.find_element(by, selector)
            return el.text.strip()
        except:
            return ""

def get_by_label(driver, label_text):
    xpath = f"//span[normalize-space()='{label_text}']/following-sibling::span[1]"
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except:
        return "غير محدد"

def scrape_ad_details(driver, url):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])
    WebDriverWait(driver, DETAILS_WAIT).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    time.sleep(OPEN_SLEEP)

    data = {"الرابط": url}
    data["العنوان"] = safe_text(driver, By.TAG_NAME, "h1")
    data["السعر"] = safe_text(driver, By.CSS_SELECTOR, 'span[aria-label="Price"]')
    data["الموقع"] = safe_text(driver, By.CSS_SELECTOR, '[aria-label="Location"]')

    labels = ["النوع", "ملكية", "المساحة (م٢)", "غرف نوم", "الحمامات", "مفروش"]
    for lbl in labels:
        data[lbl] = get_by_label(driver, lbl)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return data

def main():
    driver = build_driver()

    if os.path.exists(SAVE_PATH):
        existing_df = pd.read_excel(SAVE_PATH)
        all_rows = existing_df.to_dict(orient="records")
        existing_links = set(existing_df["الرابط"].tolist())
        print(f"📂 تم تحميل {len(all_rows)} إعلان موجود مسبقًا.")
    else:
        all_rows = []
        existing_links = set()

    try:
        for page_num in range(122, 200):  # ✅ تعديل: يبدأ من 122 وينتهي عند 199
            url = f"https://www.dubizzle.com.eg/properties/apartments-duplex-for-sale/?page={page_num}"
            print(f"\n📄 صفحة {page_num}")
            driver.get(url)
            WebDriverWait(driver, LIST_WAIT).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'a[href*="/ad/"]'))
            )
            time.sleep(2)

            anchors = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/ad/"]')
            links = []
            seen = set()
            for a in anchors:
                href = a.get_attribute("href")
                if href and "/ad/" in href and href not in seen:
                    seen.add(href)
                    links.append(href)

            print(f"📌 {len(links)} إعلان في الصفحة.")

            for link in links:
                if link in existing_links:
                    print(f"⏩ تم تخطي إعلان مكرر: {link}")
                    continue
                try:
                    row = scrape_ad_details(driver, link)
                    all_rows.append(row)
                    existing_links.add(link)
                    print(f"✅ {row.get('العنوان','')[:40]}")
                except Exception as e:
                    print(f"❌ خطأ في الإعلان: {e}")

    finally:
        driver.quit()

    pd.DataFrame(all_rows).to_excel(SAVE_PATH, index=False)
    print(f"\n💾 تم الحفظ في {SAVE_PATH} - الإجمالي: {len(all_rows)} إعلان")

if __name__ == "__main__":
    main()
