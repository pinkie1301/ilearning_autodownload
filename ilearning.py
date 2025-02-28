from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from PyPDF2 import PdfMerger
import os
import time
import requests
import logging
import configparser

class CourseBot:
    def download(i, Account, Password, Classname, Chapter):
        # 設定下載目錄(預設為當前資料夾)
        dir = os.getcwd()
        os.makedirs(dir, exist_ok=True)

        # 創建一個以 Chapter 變數命名的資料夾，將所有下載的檔案存入此資料夾中
        chapter_dir = os.path.join(dir, Chapter.replace(" ", ""))
        os.makedirs(chapter_dir, exist_ok=True)

        # 啟動 Selenium WebDriver
        driver = webdriver.Chrome()  # 確保已安裝對應版本的 ChromeDriver

        # 開啟目標網頁
        url = "https://i-learning.cycu.edu.tw/index.php"
        driver.get(url)

        try:
            # 登入步驟
            username_input = driver.find_element(By.NAME, "username")
            password_input = driver.find_element(By.NAME, "password")

            username_input.send_keys(Account)
            password_input.send_keys(Password)
            password_input.send_keys(Keys.RETURN)

            time.sleep(5)  # 等待登入完成

            # 切換到指定的 frame
            driver.switch_to.frame(driver.find_element(By.NAME, "s_main"))

            # 找到目標課程並點擊進入
            course_link = driver.find_element(By.LINK_TEXT, Classname)
            course_link.click()

            time.sleep(3)  # 等待頁面加載

            # 再次切換到 frame
            driver.switch_to.default_content()
            driver.switch_to.frame(driver.find_element(By.NAME, "s_main"))

            # 利用 Chapter 變數搜尋指定的章節（使用 f-string）
            chapter_link = driver.find_element(By.XPATH, f"//div[@class='title' and contains(text(),'{Chapter}')]")
            chapter_link.click()

            time.sleep(3)  # 等待頁面加載

            # 搜尋 PDF 連結
            pdf_links = driver.find_elements(By.TAG_NAME, "a")
            downloaded_files = []

            for index, link in enumerate(pdf_links):
                href = link.get_attribute("href")
                pdf_name = link.text.strip()

                if href and href.endswith(".pdf"):
                    # 檔案名稱設定
                    filename = f"{index+1}_{pdf_name}.pdf"
                    file_path = os.path.join(chapter_dir, filename)

                    # 下載 PDF 文件
                    try:
                        response = requests.get(href, timeout=10)
                        response.raise_for_status()
                        with open(file_path, "wb") as f:
                            f.write(response.content)

                        downloaded_files.append(file_path)
                        print(f"已下載: {filename}")
                    except requests.exceptions.RequestException as download_error:
                        print(f"下載失敗: {href}, 錯誤: {download_error}")

            # 如果有下載到 PDF 文件，進行合併操作
            if downloaded_files:
                merge_count = 1  # 用來生成不同的合併檔名
                while True:
                    print("\n下載的 PDF 文件清單:")
                    for idx, file in enumerate(downloaded_files, start=1):
                        print(f"{idx}: {os.path.basename(file)}")

                    user_input = input("\n請輸入要合併的文件範圍 (eg. 1-3：合併1到3份文件，2-4,6-8：合併2到4及6到8兩段文件): ").strip()
                    merge_indices = []

                    if user_input:
                        ranges = user_input.split(',')
                        for r in ranges:
                            r = r.strip()
                            if '-' in r:
                                parts = r.split('-')
                                if len(parts) == 2:
                                    try:
                                        start = int(parts[0])
                                        end = int(parts[1])
                                        if start > end:
                                            start, end = end, start  # 自動調整順序
                                        for i in range(start, end+1):
                                            if 1 <= i <= len(downloaded_files):
                                                merge_indices.append(i-1)  # 轉換成 0-based index
                                            else:
                                                print(f"文件編號 {i} 超出範圍")
                                    except ValueError:
                                        print(f"無效的範圍: {r}")
                            else:
                                try:
                                    num = int(r)
                                    if 1 <= num <= len(downloaded_files):
                                        merge_indices.append(num-1)
                                    else:
                                        print(f"文件編號 {num} 超出範圍")
                                except ValueError:
                                    print(f"無效的輸入: {r}")

                    merge_indices = sorted(set(merge_indices))
                    if merge_indices:
                        merger = PdfMerger()
                        for idx in merge_indices:
                            merger.append(downloaded_files[idx])
                        merged_file_path = os.path.join(chapter_dir, f"combined_{merge_count}.pdf")
                        merger.write(merged_file_path)
                        merger.close()
                        print(f"\nPDF 文件已合併: {merged_file_path}")
                        merge_count += 1
                    else:
                        print("\n沒有選擇到任何文件，合併取消。")

                    cont = input("\n是否繼續合併下一份文件？(Y/N): ").strip().upper()
                    if cont != 'Y':
                        break

            print("所有文件下載並合併完成！")

        except Exception as e:
            print(f"發生錯誤: {e}")

        finally:
            # 關閉瀏覽器
            time.sleep(3)
            driver.quit()
        
if __name__ == "__main__":
    configFilename = "accounts.ini"
    if not os.path.isfile(configFilename):
        with open(configFilename, "a",encoding="utf-8") as f:
            f.writelines(
                ["[Default]\n", "Account= your account\n", "Password= your password\n", "Classname= your classname\n", "Chapter= chapter title"]
            )
            print("請在 accounts.ini 中輸入帳密")
            exit()
    # get account info fomr ini config file
    config = configparser.ConfigParser()
    config.read(configFilename,encoding="utf-8")
    Account = config["Default"]["Account"]
    Password = config["Default"]["Password"]
    Classname = config["Default"]["Classname"]
    Chapter = config["Default"]["Chapter"]

    courseBot = CourseBot()
    courseBot.download(Account, Password, Classname, Chapter)