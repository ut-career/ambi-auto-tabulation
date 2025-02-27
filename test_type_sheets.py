import os
import time
import string
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

# 環境変数の読み込み
load_dotenv()

# AMBIログイン情報
AMBI_LOGIN_URL = os.getenv("AMBI_LOGIN_URL")
AMBI_LOGIN_ID = os.getenv("AMBI_LOGIN_ID")
AMBI_PASSWORD = os.getenv("AMBI_PASSWORD")

# 基本URLとデータ種別ごとのエンドポイント
BASE_URL = "https://en-ambi.com/company/effect_ma"
ENDPOINTS = {
    "platinum": "/acc_scout/platinum/",
    "regular": "/acc_scout/",
    "interested": "/acc_interests/",
}
COMMON_PARAMS = "PK=CA19C6"

# Google Sheets設定
SHEET_NAME = "テスト"  # スプレッドシート名
SERVICE_ACCOUNT_FILE = "service_account.json"  # サービスアカウントのJSONファイル

def setup_driver():
    """
    Chromeウェブドライバーを設定する
    """
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def login_to_ambi(driver):
    """
    AMBIサイトにログイン
    """
    driver.get(AMBI_LOGIN_URL)
    wait = WebDriverWait(driver, 10)

    username_field = wait.until(EC.presence_of_element_located((By.NAME, "accLoginID")))
    password_field = wait.until(EC.presence_of_element_located((By.NAME, "accLoginPW")))
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))

    username_field.send_keys(AMBI_LOGIN_ID)
    password_field.send_keys(AMBI_PASSWORD)
    login_button.click()

    # ログイン後の画面が読み込まれるまで待機
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    print("AMBIサイトにログインしました。")
    return True

def fetch_data_by_contact_names(driver, date, data_type, contact_names):
    """
    指定されたjobNameに対応するデータを取得する
    """
    query_params = f"?_pp_=date_from%3D{date}%7Cdate_to%3D{date}&{COMMON_PARAMS}"
    url = f"{BASE_URL}{ENDPOINTS[data_type]}{query_params}"
    
    driver.get(url)
    time.sleep(2)  # ページロード待機
    
    results = []

    for contact_name in contact_names:
        try:
            # jobNameに一致する行を特定
            row = driver.find_element(By.XPATH, f"//div[@class='jobName' and text()='{contact_name}']/ancestor::tr")
            
            # その行に含まれるdataクラスを持つtd要素を取得
            data_tds = row.find_elements(By.CLASS_NAME, "data")
            
            # 各td要素のテキストを取得
            scout_mail_stats = [td.text for td in data_tds]
            if(len(scout_mail_stats) == 8):
                scout_mail_stats_dict = dict(
                    contact_name = scout_mail_stats[0],
                    interested_count = scout_mail_stats[1],
                    passed_judgement_count = scout_mail_stats[2],
                    passed_judgement_rate = scout_mail_stats[3],
                    entry_count = scout_mail_stats[4],
                    entry_rate = scout_mail_stats[5],
                    interview_req_count = scout_mail_stats[6],
                    interview_req_rate = scout_mail_stats[7],
                )
            else:
                scout_mail_stats_dict = dict(
                    contact_name = scout_mail_stats[0],
                    send_count = scout_mail_stats[1],
                    opens_count = scout_mail_stats[2],
                    open_rate = scout_mail_stats[3],
                    refusals_count = scout_mail_stats[4],
                    entry_count = scout_mail_stats[5],
                    post_opening_entry_rate = scout_mail_stats[6],
                    entry_rate = scout_mail_stats[7],
                    interview_req_count = scout_mail_stats[8],
                    interview_req_rate = scout_mail_stats[9],
                )

            # 結果を保存
            results.append({
                "contact_name": contact_name,
                "scout_mail_stats_dict": scout_mail_stats_dict
            })
            print(f"{contact_name}のデータ:{scout_mail_stats_dict}")
        
        except Exception as e:
            print(f"{contact_name}のデータ取得中にエラーが発生しました: {e}")
    
    return results

def get_current_month():
    today = datetime.today()
    current_month = f"{today.year}.{today.month:02d}"
    toStr = str(current_month)
    return toStr

# def test_write_to_google_sheets():
#     data_list = [
#         {
#             "jobName": "山中沙矢",
#             "data": ['山中沙矢', '3', '0', '0.0%', '0', '0', '---', '0.0%', '0', '---']
#         },
#         {
#             "jobName": "橘萌生",
#             "data": ['橘萌生', '2', '0', '0.0%', '0', '0', '---', '0.0%', '0', '---']
#         },
#         {
#             "jobName": "奥野翔子",
#             "data": ['奥野翔子', '2', '0', '0.0%', '0', '0', '---', '0.0%', '0', '---']
#         }
#     ]

#     """
#     データをGoogleスプレッドシートに書き込む
#     """
#     # Google Sheets API認証
#     scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
#     credentials = Credentials.from_service_account_file(
#         SERVICE_ACCOUNT_FILE,
#         scopes=scope
#     )
#     gc = gspread.authorize(credentials)
   

#     current_month_value = get_current_month()
#     # スプレッドシートを取得
#     sheet = gc.open("テスト").worksheet("2025.01")
#     sheet_data = sheet.get_all_values()

#     # 書き込み
#     row = 2  # データの開始行（1行目はヘッダー）
#     for entry in data_list:
#         sheet.update_cell(row, 1, entry["date"])           # 日付
#         sheet.update_cell(row, 2, entry["data_type"])      # データ種別
#         sheet.update_cell(row, 3, entry["jobName"])        # jobName
#         sheet.update_cell(row, 4, ", ".join(entry["data"]))  # データ（カンマ区切り）
#         row += 1

#     print("Googleスプレッドシートにデータを書き込みました。")

def data_entry_position(contact_name, scout_type):
    """
    データを書き込む位置を取得
    """
    if scout_type == "platinum":
        if contact_name == "山中沙矢":
            return 20
        elif contact_name == "橘萌生":
            return 37
        elif contact_name == "奥野翔子":
            return 71
    elif scout_type == "regular":
        if contact_name == "山中沙矢":
            return 25
        elif contact_name == "橘萌生":
            return 42
        elif contact_name == "奥野翔子":
            return 76
    elif scout_type == "interested":
        if contact_name == "山中沙矢":
            return 30
        elif contact_name == "橘萌生":
            return 47
        elif contact_name == "奥野翔子":
            return 81


def get_column_from_date(date_str):
    """
    YYYY-MM-DD形式の日付を受け取り、対応する列番号（整数）を返す。
    例: '2025-01-01' → 7 (G列), '2025-01-02' → 8 (H列)
    """

    base_column_index = ord("G") - ord("A") + 1  # 'G'列はExcel上で7列目
    day = int(date_str.split("-")[2])  # 日にちを取得
    column_number = base_column_index + day - 1  # G列からスタート
    print(f"column_number の値: {column_number}, 型: {type(column_number)}")
    return column_number  

def write_to_google_sheets(all_scout_data):
    # all_scout_data = [
    #     ("2025-01-01", "platinum", "山中沙矢", {'contact_name': '山中沙矢', 'send_count': '3', 'opens_count': '0', 'open_rate': '0.0%', 'refusals_count': '0', 'entry_count': '0', 'post_opening_entry_rate': '---', 'entry_rate': '0.0%', 'interview_req_count': '0', 'interview_req_rate': '---'}),
    #     ("2025-01-02", "regular", "橘萌生", {'contact_name': '橘萌生', 'send_count': '2', 'opens_count': '1', 'open_rate': '50.0%', 'refusals_count': '0', 'entry_count': '1', 'post_opening_entry_rate': '100.0%', 'entry_rate': '50.0%', 'interview_req_count': '1', 'interview_req_rate': '50.0%'}),
    #     ("2025-01-03", "interested", "奥野翔子", {'contact_name': '奥野翔子', 'interested_count': '2', 'passed_judgement_count': '0', 'passed_judgement_rate': '0.0%', 'refusals_count': '0', 'entry_count': '0','entry_rate': '0.0%', 'interview_req_count': '0', 'interview_req_rate': '---'}),
    # ]

    """
    データをGoogleスプレッドシートに書き込む
    """
    # Google Sheets API認証
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scope
    )
    gc = gspread.authorize(credentials)

    


    # current_month_value = get_current_month()
    # スプレッドシートを取得
    sheet = gc.open("テスト").worksheet("シート2")
    sheet_data = sheet.get_all_values()
    
    # 書き込み
    for entry in all_scout_data:
        print(entry)
        date = entry[0]
        scout_type = entry[1]
        contact_name = entry[2]
        scout_mail_stats_dict = entry[3]
        send_count = scout_mail_stats_dict["send_count"] if scout_type in ["platinum", "regular"] else 0 # 送信数
        opens_count = scout_mail_stats_dict["opens_count"] if scout_type in ["platinum", "regular"] else 0 # 開封数
        entry_count = scout_mail_stats_dict["entry_count"]
        interested_count = scout_mail_stats_dict["interested_count"] if scout_type == "interested" else 0
        interested_entry_count = scout_mail_stats_dict["entry_count"] if scout_type == "interested" else 0

        # 該当セルにデータを書き込む
        if scout_type in ["platinum", "regular"]:
          # これで型を確認
            sheet.update_cell(data_entry_position(contact_name,scout_type), get_column_from_date(date), send_count)  # 送信数
            sheet.update_cell(data_entry_position(contact_name,scout_type) + 1, get_column_from_date(date), opens_count)   # 開封数
            sheet.update_cell(data_entry_position(contact_name,scout_type) + 3, get_column_from_date(date), entry_count)  # エントリー数
        elif scout_type == "interested":
            sheet.update_cell(data_entry_position(contact_name,scout_type), get_column_from_date(date), interested_count)  # 興味あり数
            sheet.update_cell(data_entry_position(contact_name,scout_type) + 1, get_column_from_date(date), interested_entry_count)  # エントリー数

print("データの更新が完了しました！")

def main():
    driver = None
    all_scout_data = []

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
        contact_names = ["山中沙矢", "橘萌生", "奥野翔子"]

        for single_date in (start_date + timedelta(n) for n in range(3)):
            formatted_date = single_date.strftime("%Y-%m-%d")
            print(f"\n{formatted_date}のデータ収集を開始:")

            for data_type in ENDPOINTS.keys():
                data = fetch_data_by_contact_names(driver, formatted_date, data_type, contact_names)
                
                for entry in data:
                    all_scout_data.append({
                        "date": formatted_date,
                        "data_type": data_type,
                        "contact_name": entry["contact_name"],
                        "scout_mail_stats_dict": entry["scout_mail_stats_dict"]
                    })

                time.sleep(1)
        # データをGoogleスプレッドシートに書き込む
        print(all_scout_data)
        write_to_google_sheets(all_scout_data)

    except Exception as e:
        print(f"スクリプト実行中に致命的なエラーが発生しました: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
