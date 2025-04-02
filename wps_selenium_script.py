import os
import time
import glob
import shutil
import sys
import subprocess
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------------------------
# 경로 설정
# ---------------------------------
DOWNLOAD_DIR = "/Users/hyeonuk/Downloads"
# 로컬 판매 파일도 Downloads 폴더에 "sales.xlsx"로 둠 (원하면 다른 경로 가능)
LOCAL_SALES_PATH = os.path.join(DOWNLOAD_DIR, "sales.xlsx")

# ---------------------------------
# 크롬드라이버 생성
# ---------------------------------
def get_chrome_driver():
    """Selenium용 ChromeDriver 생성"""
    chrome_options = Options()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # 필요시 headless 모드
    # chrome_options.add_argument("--headless")

    return webdriver.Chrome(options=chrome_options)

# ---------------------------------
# Explicit Wait 유틸
# ---------------------------------
def wait_for_user_center(driver, max_wait=60):
    """
    최대 max_wait초 동안 'User Center' 화면 뜨는지 감지.
    예: <div class="header-title">User Center</div> 이 보이면 로그인 완료로 판단.
    """
    wait = WebDriverWait(driver, max_wait)
    wait.until(
        EC.text_to_be_present_in_element((By.CSS_SELECTOR, "div.header-title"), "User Center")
    )

def wait_for_element(driver, by, locator, max_wait=10):
    """특정 요소가 DOM에 나타날 때까지 최대 max_wait초 대기"""
    wait = WebDriverWait(driver, max_wait)
    return wait.until(EC.presence_of_element_located((by, locator)))

# ---------------------------------
# 파일 다운로드 대기
# ---------------------------------
def wait_for_file_download(keyword="stock", timeout=30):
    """
    DOWNLOAD_DIR에서 .xlsx 중 filename에 keyword가 들어간 파일이
    최대 timeout초 내 생길 때까지 폴링. 찾으면 'stock.xlsx'로 rename 후 경로 반환.
    """
    start_time = time.time()
    found_path = None
    while True:
        xlsx_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        for f in xlsx_files:
            if keyword.lower() in os.path.basename(f).lower():
                found_path = f
                break
        if found_path:
            break
        if time.time() - start_time > timeout:
            raise Exception(f"'{keyword}' 파일을 {timeout}초 내에 찾지 못했습니다.")
        time.sleep(1)
    final_path = os.path.join(DOWNLOAD_DIR, "stock.xlsx")
    if found_path != final_path:
        shutil.move(found_path, final_path)
    return final_path

# ---------------------------------
# 재고+판매 엑셀 병합
# ---------------------------------
def merge_local_sales_with_downloaded_stock(stock_file: str, sales_file: str) -> int:
    """
    stock_file + sales_file 병합, 재고 차감 후 'stock_updated.xlsx'로 출력.
    """
    if not os.path.exists(stock_file):
        raise FileNotFoundError("다운로드된 재고 파일이 없음.")
    if not os.path.exists(sales_file):
        raise FileNotFoundError("로컬 판매 파일(sales.xlsx)이 없음.")

    stock_df = pd.read_excel(stock_file)
    sales_df = pd.read_excel(sales_file)

    # 판매 수량 합산
    sales_agg = sales_df.groupby("product_id")["quantity_sold"].sum().reset_index()

    # 병합 & 차감
    merged_df = pd.merge(stock_df, sales_agg, on="product_id", how="left")
    merged_df["quantity_sold"] = merged_df["quantity_sold"].fillna(0)
    merged_df["updated_stock"] = merged_df["stock_qty"] - merged_df["quantity_sold"]
    merged_df.loc[merged_df["updated_stock"] < 0, "updated_stock"] = 0

    stock_df["stock_qty"] = merged_df["updated_stock"]

    updated_path = os.path.join(DOWNLOAD_DIR, "stock_updated.xlsx")
    stock_df.to_excel(updated_path, index=False)

    return len(sales_df)

# ---------------------------------
# AppleScript: 한 글자씩 천천히 입력
# ---------------------------------
def type_slowly_mac(path_str):
    """
    AppleScript로 macOS에서 `path_str`을 한 글자씩 천천히 입력하고,
    Enter를 2번 눌러 Finder '열기' 버튼까지 자동 실행하는 예시.
    """
    script_lines = [
        'tell application "System Events"'
    ]

    # 1) 한 글자씩 천천히 입력
    for ch in path_str:
        if ch == '"':
            script_lines.append('    keystroke "\\"')
        elif ch == '\\':
            script_lines.append('    keystroke "\\\\"')
        else:
            script_lines.append(f'    keystroke "{ch}"')
        script_lines.append('    delay 0.1')  # 각 글자 입력 후 약간 쉼

    # 2) 첫 번째 Enter (경로 입력 완료)
    script_lines.append('    key code 36')  # 36 = Return
    script_lines.append('    delay 1')      # 조금 대기

    # 3) 두 번째 Enter (Finder가 경로 선택 후 '열기' 버튼 누르는 효과)
    script_lines.append('    key code 36')
    script_lines.append('end tell')

    full_script = "\n".join(script_lines)
    subprocess.run(["osascript", "-e", full_script])

# ---------------------------------
# 메인 로직: 로그인 -> 문서 페이지 -> 다운로드 -> 병합
# ---------------------------------
def download_merge_upload_with_finder():
    driver = get_chrome_driver()
    stock_file_path = None
    try:
        # (1) 로그인
        driver.get("https://account.wps.com/")
        # User Center 뜨면 로그인 완료
        wait_for_user_center(driver, max_wait=120)
        print("User Center 감지됨 → 문서 페이지로 이동")

        # (2) 스프레드시트 페이지
        driver.get("https://sg.docs.wps.com/p/89316816846831")
        time.sleep(10)

        # (3) File Operations 버튼
        file_ops_btn = wait_for_element(
            driver, By.CSS_SELECTOR, "button.kd-button.kd-button-icon", max_wait=30
        )
        file_ops_btn.click()
        print("File Operations 버튼 클릭")

        # (4) Download 항목
        download_elem = wait_for_element(
            driver, By.XPATH, '//div[@data-key="Download"]', max_wait=30
        )
        download_elem.click()
        print("Download 메뉴 클릭")

        # (5) 다운 완료 대기 -> stock.xlsx
        stock_file_path = wait_for_file_download("stock", timeout=30)
        print(f"다운로드된 재고 파일: {stock_file_path}")

        # 3) 병합
        updated_path = os.path.join(DOWNLOAD_DIR, "stock_updated.xlsx")
        merge_local_sales_with_downloaded_stock(stock_file_path, LOCAL_SALES_PATH)
        print("병합 완료 →", updated_path)

        # 4) 업로드 페이지
        driver.get("https://docs.wps.com/")  # 업로드 버튼 있는 URL

        # 5) "Upload" 버튼 클릭
        upload_btn = wait_for_element(driver, By.CSS_SELECTOR, "label.upload-btn-warp", max_wait=10)
        upload_btn.click()
        print("Upload 버튼 클릭")

        # 6) "File" 버튼 클릭 -> Finder 창 열림
        file_btn = wait_for_element(driver, By.CSS_SELECTOR, "label.upload-file", max_wait=10)
        file_btn.click()
        print("File 버튼 클릭 -> OS 대화상자 열림")

        # ---------------------
        # (A) macOS -> AppleScript
        # (B) Windows -> pywinauto
        # ---------------------
        if sys.platform == "darwin":
            # macOS: 한 글자씩 천천히 입력
            print(f"AppleScript로 (천천히) 경로 입력: {updated_path}")
            type_slowly_mac(updated_path)
            time.sleep(5)
            print("업로드 완료(추정)!")
        # elif sys.platform.startswith("win"):
        #     # Windows: pywinauto 사용 (예시)
        #     # pip install pywinauto
        #     from pywinauto import Desktop
        #     from pywinauto.controls.win32_controls import EditWrapper, ButtonWrapper

        #     print("Windows OS -> pywinauto 사용, 파일 열기 창 제어")
        #     # "열기" 대화상자의 제목(한글 OS: "열기", 영문 OS: "Open") 확인
        #     app = Desktop(backend="win32")
        #     dlg = app["열기"]  # or "Open"

        #     edit_box = EditWrapper(dlg["Edit"])
        #     edit_box.set_text(updated_path)

        #     open_btn = ButtonWrapper(dlg["열기"])  # 영문: dlg["Open"]
        #     open_btn.click()
        #     time.sleep(5)
        #     print("업로드 완료(추정)!")
        else:
            print("이 OS는 아직 자동화 로직 없음. 수동으로 선택 필요.")
            time.sleep(5)

        # ---------------------
        # 종료
        # ---------------------
        print("업로드 프로세스 끝!")
    finally:
        driver.quit()

if __name__ == "__main__":
    download_merge_upload_with_finder()