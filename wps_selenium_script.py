# -------------------------------------------------------
# 아래는 파이썬 표준 라이브러리 및 외부 라이브러리(셀레니움, 판다스 등)를 임포트하는 구문입니다.
# 각 모듈이 어떤 역할을 하는지 간략히 정리했습니다.
# -------------------------------------------------------

# 1) os : 운영체제와 상호작용하기 위한 다양한 함수들이 들어있습니다.
#    - 파일 경로를 합치거나(os.path.join), 디렉터리 생성(os.mkdir), 
#      환경 변수(os.environ), 프로세스 제어(os.kill) 등 시스템 레벨의 기능을 제공합니다.
import os

# 2) time : 시간 관련 함수가 모여있는 모듈입니다.
#    - time.sleep(초)로 특정 시간만큼 대기하거나,
#      time.time()으로 현재 시간을 초 단위로 얻는 등의 기능을 합니다.
import time

# 3) glob : 특정 패턴(예: '*.xlsx', '*.csv')에 맞는 파일들을 찾아 목록으로 반환해줍니다.
#    - 예: glob.glob('*.py') → 현재 디렉터리 내 .py 파일들을 리스트로 가져옴.
import glob

# 4) shutil : 고수준의 파일/디렉터리 작업을 위한 모듈입니다.
#    - 파일 복사(shutil.copy), 폴더 전체 복사(shutil.copytree), 이동(shutil.move), 삭제(rmtree) 등을 지원.
import shutil

# 5) sys : 파이썬 인터프리터와 관련된 정보를 제어하고 확인할 때 사용합니다.
#    - sys.argv : 명령줄 인수, sys.exit() : 인터프리터 종료, sys.path : 모듈 검색 경로 등.
import sys

# 6) subprocess : 새로운 프로세스를 생성하고 관리할 수 있는 모듈입니다.
#    - 예: subprocess.run(["ls", "-l"]) → OS 명령 실행,
#      명령의 결과, 표준 입출력 등을 파이썬에서 제어 가능.
import subprocess

# 7) pandas : 데이터 분석을 손쉽게 해주는 핵심 라이브러리입니다.
#    - 엑셀이나 CSV 파일 로드(pd.read_excel, pd.read_csv), 
#      데이터프레임을 이용한 필터링, 그룹화, 결합 등 고급 연산을 지원.
import pandas as pd

# -------------------------------------------------------
# 셀레니움(Selenium) 관련 임포트.
# 웹드라이버를 이용해 브라우저를 자동 제어(크롤링, 테스트 자동화)하는 데 쓰입니다.
# -------------------------------------------------------

# webdriver : 브라우저(Chrome, Firefox 등)를 프로그래밍적으로 실행하고 명령을 전달하기 위한 핵심 클래스.
from selenium import webdriver

# By : 셀레니움에서 요소를 찾을 때(예: By.ID, By.XPATH, By.CSS_SELECTOR) 사용할 상수/메서드를 제공.
from selenium.webdriver.common.by import By

# Options : Chrome 브라우저 실행 시 특정 옵션(예: headless 모드, 브라우저 창 크기 등)을 설정할 때 사용.
from selenium.webdriver.chrome.options import Options

# WebDriverWait, expected_conditions(EC) : 
#    웹 자동화 시, 특정 요소가 표시될 때까지 대기하거나(element_to_be_clickable),
#    텍스트가 특정 값이 될 때까지 대기하는 등의 동기화 제어를 위해 사용.
#    즉, 페이지 로딩 지연, AJAX 갱신 등을 고려해 안정적인 테스트/크롤링이 가능케 함.
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------------------------
# 경로 설정
# ---------------------------------
DOWNLOAD_DIR = "/Users/hyeonuk/Downloads"
# 로컬 판매 파일(sales.xlsx)을 기본으로 Downloads 폴더에 두도록 설정합니다.
# 필요에 따라 다른 디렉터리를 지정할 수도 있습니다.
LOCAL_SALES_PATH = os.path.join(DOWNLOAD_DIR, "sales.xlsx")

# ---------------------------------
# 크롬드라이버 생성
# ---------------------------------
def get_chrome_driver():
    """
    Selenium용 ChromeDriver 인스턴스를 생성하는 함수.

    1) chrome_options를 통해 다운로드 설정(pref) 및 헤드리스 모드 등
       다양한 브라우저 환경 옵션을 지정할 수 있습니다.
    2) "download.default_directory" 등으로 다운로드 경로를 DOWNLOAD_DIR로 설정하여,
       웹에서 다운받는 파일(예: 엑셀, CSV 등)을 지정된 폴더에 저장합니다.
    3) 필요할 경우 headless 모드(브라우저 UI 없이 백그라운드 실행)를 활성화하려면
       chrome_options.add_argument("--headless") 주석을 해제하면 됩니다.
    """

    chrome_options = Options()

    # (A) prefs를 통해 다운로드 설정을 지정.
    #     - "download.default_directory": 자동 다운로드 경로
    #     - "download.prompt_for_download": False → 다운로드 시 묻지 않고 진행
    #     - "download.directory_upgrade": True → 기존 폴더보다 상위폴더 설정 가능
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # (B) 필요할 경우 다음 옵션을 활성화하면 브라우저 UI 없이 백엔드에서 크롤링/테스트 실행 가능.
    # chrome_options.add_argument("--headless")

    # (C) webdriver.Chrome()에 준비한 chrome_options를 적용하여 드라이버 인스턴스를 생성.
    return webdriver.Chrome(options=chrome_options)

# ---------------------------------
# Explicit Wait 유틸
# ---------------------------------
def wait_for_user_center(driver, max_wait=60):
    """
    최대 max_wait초 동안 'User Center' 텍스트가 나타날 때까지 대기.

    주로 로그인 완료 후 특정 페이지가 뜨는지 확인할 때 사용합니다.
    예: <div class="header-title">User Center</div> 요소가 렌더링되면 
       로그인 절차가 성공적으로 끝났다고 판단.

    :param driver: Selenium WebDriver 객체(ChromeDriver 등).
    :param max_wait: 최대 대기 시간(초).
    :raises TimeoutException: 지정 시간 안에 해당 텍스트가 안 나타나면 예외 발생.
    """

    wait = WebDriverWait(driver, max_wait)
    wait.until(
        EC.text_to_be_present_in_element((By.CSS_SELECTOR, "div.header-title"), "User Center")
    )

def wait_for_element(driver, by, locator, max_wait=10):
    """
    특정 요소가 DOM(Document Object Model)에 나타날 때까지 최대 max_wait초 대기.

    :param driver: Selenium WebDriver 객체.
    :param by: 요소를 찾는 방식(예: By.CSS_SELECTOR, By.ID 등).
    :param locator: 찾을 선택자(예: "div.header-title" 등).
    :param max_wait: 최대 대기 시간(초).
    :return: 찾은 WebElement 객체.
    :raises TimeoutException: 지정 시간 안에 요소가 없으면 예외 발생.
    """

    wait = WebDriverWait(driver, max_wait)
    return wait.until(EC.presence_of_element_located((by, locator)))





# ---------------------------------
# 파일 다운로드 대기
# ---------------------------------
def wait_for_file_download(keyword="stock", timeout=30):
    """
    지정된 DOWNLOAD_DIR에서, 파일명에 'keyword'가 들어간 .xlsx 파일이 
    최대 timeout초 내에 생길 때까지 대기(polling)하는 함수입니다.

    1) 1초 간격으로 DOWNLOAD_DIR 내 "*.xlsx" 파일들을 확인(glob)합니다.
    2) 파일 중 이름에 keyword(대소문자 구분 없이)가 포함되어 있으면 다운로드 완료로 판단해 루프를 탈출합니다.
    3) 만약 timeout초가 지날 때까지 해당 파일을 찾지 못하면 예외(Exception)를 발생시킵니다.
    4) 찾은 파일(found_path)은 'stock.xlsx'라는 고정 파일명으로 rename하여 
       최종 경로(final_path)에 저장합니다(중복 명 처리, 통일된 파일명 관리 목적).
    5) rename한 최종 경로(final_path)를 반환합니다.

    :param keyword: 검색할 키워드(기본값: "stock"). 예: "stock", "invoice", ...
    :param timeout: 최대 대기 시간(초). 기본값 30초.
    :return: 최종적으로 rename된 파일의 전체 경로(final_path).
    :raises Exception: 주어진 시간(timeout) 내에 keyword가 들어간 .xlsx 파일을 찾지 못한 경우.
    """

    start_time = time.time()
    found_path = None
    while True:
        # glob로 DOWNLOAD_DIR 내 모든 .xlsx 파일 목록을 가져옵니다.
        xlsx_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        for f in xlsx_files:
            # 파일명에서 keyword를 찾으면 found_path에 할당하고 break
            if keyword.lower() in os.path.basename(f).lower():
                found_path = f
                break

        if found_path:
            # 해당 파일을 찾으면 대기 루프 탈출
            break

        # 아직 못 찾았고, timeout을 넘겼다면 예외 발생
        if time.time() - start_time > timeout:
            raise Exception(f"'{keyword}' 파일을 {timeout}초 내에 찾지 못했습니다.")
        
        # 1초 대기 후 다시 시도
        time.sleep(1)

    # ---------------------------------
    # 파일명을 'stock.xlsx'로 고정 rename한 후 경로 반환
    # ex) found_path = '/Downloads/stock(1).xlsx' → '/Downloads/stock.xlsx'
    # ---------------------------------
    final_path = os.path.join(DOWNLOAD_DIR, "stock.xlsx")
    if found_path != final_path:
        shutil.move(found_path, final_path)
    return final_path

# ---------------------------------
# 재고+판매 엑셀 병합
# ---------------------------------
def merge_local_sales_with_downloaded_stock(stock_file: str, sales_file: str) -> int:
    """
    다운로드 받은 재고 파일(stock_file)과 
    로컬 판매 파일(sales_file)을 병합해 재고를 차감한 뒤, 
    'stock_updated.xlsx'라는 이름으로 저장하는 함수.

    1) stock_file(엑셀)과 sales_file(엑셀)을 각각 읽어서 
       Pandas DataFrame(stock_df, sales_df)에 로드합니다.
    2) sales_df를 product_id 기준으로 판매 개수(quantity_sold) 합계를 구하여 
       stock_df와 병합.
    3) 기존 재고(stock_qty)에서 판매 수량(quantity_sold)을 차감해 updated_stock를 계산.
       음수가 되지 않도록 0으로 처리합니다.
    4) 최종 재고를 'stock_updated.xlsx' 파일에 저장하고,
       이 함수가 처리한 판매 건수(len(sales_df))를 정수로 반환합니다.

    :param stock_file: 다운로드된 재고 파일 경로(str).
    :param sales_file: 로컬 판매 파일 경로(str).
    :return: 판매 엑셀(sales_df)의 행 수(판매 건수)를 int로 반환.
    :raises FileNotFoundError: stock_file 또는 sales_file이 존재하지 않을 경우.
    """

    # ---------------------------------
    # 파일 존재 여부 체크
    # ---------------------------------
    if not os.path.exists(stock_file):
        raise FileNotFoundError("다운로드된 재고 파일이 없음.")
    if not os.path.exists(sales_file):
        raise FileNotFoundError("로컬 판매 파일(sales.xlsx)이 없음.")

    # ---------------------------------
    # 엑셀 읽기
    # ---------------------------------
    stock_df = pd.read_excel(stock_file)
    sales_df = pd.read_excel(sales_file)

    # ---------------------------------
    # 판매 수량 합산(sales_agg)
    # product_id별 quantity_sold 합계를 구하여, 
    # 그룹화 결과를 sales_agg에 저장
    # ---------------------------------
    sales_agg = sales_df.groupby("product_id")["quantity_sold"].sum().reset_index()

    # ---------------------------------
    # stock_df와 sales_agg를 병합
    # how="left"로, 재고에 없는 product_id는 무시, 
    # 재고에 있지만 판매된 적이 없다면 NaN(→0 처리)
    # ---------------------------------
    merged_df = pd.merge(stock_df, sales_agg, on="product_id", how="left")
    merged_df["quantity_sold"] = merged_df["quantity_sold"].fillna(0)

    # ---------------------------------
    # updated_stock = stock_qty - quantity_sold
    # 음수가 되면 0으로 처리하여 재고가 음수가 되지 않도록 함.
    # ---------------------------------
    merged_df["updated_stock"] = merged_df["stock_qty"] - merged_df["quantity_sold"]
    merged_df.loc[merged_df["updated_stock"] < 0, "updated_stock"] = 0

    # ---------------------------------
    # 병합 결과를 원본 stock_df에 반영한 뒤,
    # 'stock_updated.xlsx'로 저장
    # ---------------------------------
    stock_df["stock_qty"] = merged_df["updated_stock"]
    updated_path = os.path.join(DOWNLOAD_DIR, "stock_updated.xlsx")
    stock_df.to_excel(updated_path, index=False)

    # 반환값: 판매 파일의 총 행 수(= len(sales_df))
    # 참고: 실제 판매 건수(총 qty 합)는 아니라, sales_df의 행 row 개수를 의미.
    return len(sales_df)

# ---------------------------------
# AppleScript: 한 글자씩 천천히 입력
# ---------------------------------
def type_slowly_mac(path_str):
    """
    macOS 환경에서 AppleScript를 통해 path_str을 Finder에 한 글자씩 천천히 입력한 뒤,
    엔터(Enter) 키를 두 번 눌러 경로를 열기(= '열기' 버튼 클릭 효과)하는 예시 기능.

    1) AppleScript에 'tell application "System Events"' 블록을 만들어,
       keystroke 명령으로 path_str의 각 글자를 0.1초 간격(delay 0.1)으로 입력.
    2) key code 36(Enter)을 두 번 호출해 입력을 완료하고, 
       Finder 상에서 열기 동작을 실행.
    3) Windows나 Linux 환경에서는 작동하지 않으며, 
       macOS 전용 AppleScript 기능임에 주의.

    :param path_str: Finder에 입력할 경로 문자열.
    :return: None
    """

    script_lines = [
        'tell application "System Events"'
    ]

    # 1) 한 글자씩 천천히 입력
    for ch in path_str:
        if ch == '"':
            # " → \" 처리
            script_lines.append('    keystroke "\\"')
        elif ch == '\\':
            # \ → \\ 처리
            script_lines.append('    keystroke "\\\\"')
        else:
            # 일반 문자는 그대로
            script_lines.append(f'    keystroke "{ch}"')
        script_lines.append('    delay 0.1')  # 각 글자 입력 후 0.1초 대기

    # 2) 첫 번째 Enter: 경로 입력 완료
    script_lines.append('    key code 36')  # 36 = Return
    script_lines.append('    delay 1')      # 조금 대기

    # 3) 두 번째 Enter: '열기' 버튼 누르는 효과
    script_lines.append('    key code 36')
    script_lines.append('end tell')

    full_script = "\n".join(script_lines)
    subprocess.run(["osascript", "-e", full_script])





# ---------------------------------
# 메인 로직: 로그인 -> 문서 페이지 -> 다운로드 -> 병합 -> 업로드
# ---------------------------------
def download_merge_upload_with_finder():
    """
    (A) 웹 로그인 & 문서 페이지 접근:
        1) get_chrome_driver()로 Selenium WebDriver를 생성하고, WPS Docs 계정에 로그인.
        2) wait_for_user_center() 함수를 통해 'User Center' 텍스트가 화면에 뜰 때까지 대기(로그인 성공 판단).

    (B) 스프레드시트 문서 열람 & 다운로드:
        1) 특정 스프레드시트 페이지(https://sg.docs.wps.com/p/89316816846831)에 접속해 페이지 로딩을 기다림.
        2) File Operations 메뉴를 클릭하여 'Download' 항목을 선택.
        3) wait_for_file_download("stock", timeout=30)로 'stock.xlsx' 파일 다운 완료를 대기.
           - 파일명이 'stock'을 포함하고 있으면 잡아서 'stock.xlsx'라는 고정 이름으로 rename.

    (C) 로컬 판매 파일(sales.xlsx)과 다운로드된 재고 파일(stock.xlsx) 병합:
        1) merge_local_sales_with_downloaded_stock(...)로 재고와 판매 기록을 병합해 재고를 차감.
        2) 새로 만들어진 'stock_updated.xlsx'를 로컬에 저장.

    (D) WPS Docs 업로드:
        1) 업로드 페이지(https://docs.wps.com/)로 이동.
        2) "Upload" → "File" 버튼을 차례로 클릭해 Finder(또는 OS 파일 선택 창) 열기.
        3) macOS인 경우 type_slowly_mac(...)를 사용해 경로를 한 글자씩 천천히 입력 후 Enter 2회로 업로드.
           - 다른 OS면 자동화 로직이 없어, 일단 수동으로 선택하도록 5초 대기.

    (E) 종료:
        - 모든 과정이 끝나면 driver.quit()로 브라우저 세션 종료.

    :return: None
    """

    driver = get_chrome_driver()
    stock_file_path = None
    try:
        # ---------------------------------
        # (1) 로그인 과정
        # ---------------------------------
        driver.get("https://account.wps.com/")
        # 'User Center' 텍스트로 로그인 성공 판단
        wait_for_user_center(driver, max_wait=120)
        print("User Center 감지됨 → 문서 페이지로 이동")

        # ---------------------------------
        # (2) 스프레드시트 페이지
        # ---------------------------------
        driver.get("https://sg.docs.wps.com/p/89316816846831")
        time.sleep(10) # 페이지 로딩 대기(임시)

        # ---------------------------------
        # (3) File Operations 버튼 클릭
        # ---------------------------------
        file_ops_btn = wait_for_element(
            driver, By.CSS_SELECTOR, "button.kd-button.kd-button-icon", max_wait=30
        )
        file_ops_btn.click()
        print("File Operations 버튼 클릭")

        # ---------------------------------
        # (4) Download 항목 클릭
        # ---------------------------------
        download_elem = wait_for_element(
            driver, By.XPATH, '//div[@data-key="Download"]', max_wait=30
        )
        download_elem.click()
        print("Download 메뉴 클릭")

        # ---------------------------------
        # (5) 다운 완료 대기 -> 'stock.xlsx'
        # ---------------------------------
        stock_file_path = wait_for_file_download("stock", timeout=30)
        print(f"다운로드된 재고 파일: {stock_file_path}")

        # ---------------------------------
        # (6) 병합 (재고+판매)
        # ---------------------------------
        updated_path = os.path.join(DOWNLOAD_DIR, "stock_updated.xlsx")
        merge_local_sales_with_downloaded_stock(stock_file_path, LOCAL_SALES_PATH)
        print("병합 완료 →", updated_path)

        # ---------------------------------
        # (7) 업로드 페이지 이동
        # ---------------------------------
        driver.get("https://docs.wps.com/")  # 업로드 버튼 있는 URL

        # ---------------------------------
        # (8) "Upload" 버튼 클릭
        # ---------------------------------
        upload_btn = wait_for_element(driver, By.CSS_SELECTOR, "label.upload-btn-warp", max_wait=10)
        upload_btn.click()
        print("Upload 버튼 클릭")
        
        # ---------------------------------
        # (9) "File" 버튼 클릭 -> Finder 창 열림
        # ---------------------------------
        file_btn = wait_for_element(driver, By.CSS_SELECTOR, "label.upload-file", max_wait=10)
        file_btn.click()
        print("File 버튼 클릭 -> OS 대화상자 열림")

        # ---------------------------------
        # (A) macOS 환경에서 AppleScript 자동 입력
        # ---------------------------------
        if sys.platform == "darwin":
            # macOS: 한 글자씩 천천히 입력
            print(f"AppleScript로 (천천히) 경로 입력: {updated_path}")
            type_slowly_mac(updated_path)
            time.sleep(5)
            print("업로드 완료(추정)!")
        else:
            # Windows/Linux 등 macOS 외 환경은 자동화 미지원 → 수동 선택 필요
            print("이 OS는 아직 자동화 로직 없음. 수동으로 선택 필요.")
            time.sleep(5)

        print("업로드 프로세스 끝!")
    finally:
        # ---------------------------------
        # 드라이버 세션 종료
        # ---------------------------------
        driver.quit()

if __name__ == "__main__":
    download_merge_upload_with_finder()