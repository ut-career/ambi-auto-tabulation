import os
import time
from datetime import datetime, timedelta
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials

def main():
    driver = None
    all_data = []

    try:
        # Webドライバーの設定
        driver = setup_driver()

        # AMBIにログイン
        if not login_to_ambi(driver):
            return

        # データ収集
        today = datetime.today()
        start_date = today - timedelta(days=3)  # 過去7日分
        # 指定するjobNameのリスト
        target_job_names = ["山中沙矢", "橘萌生", "奥野翔子"]

        for single_date in (start_date + timedelta(n) for n in range(3)):
            formatted_date = single_date.strftime("%Y-%m-%d")
            print(f"\n{formatted_date}のデータ収集を開始:")

            for data_type in ENDPOINTS.keys():
                data = fetch_data_by_job_names(driver, formatted_date, data_type, target_job_names)
                
                for entry in data:
                    all_data.append({
                        "date": formatted_date,
                        "data_type": data_type,
                        "jobName": entry["jobName"],
                        "data": entry["data"]
                    })

                time.sleep(1)

        # データをGoogleスプレッドシートに書き込む
        write_to_google_sheets(all_data)

    except Exception as e:
        print(f"スクリプト実行中に致命的なエラーが発生しました: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()