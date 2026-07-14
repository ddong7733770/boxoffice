import os
import sys
import time
import argparse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def get_chrome_options(login_mode=False):
    options = Options()
    
    # 로컬 작업 디렉토리 내부에 전용 크롬 프로필 폴더 생성 및 지정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    chrome_profile_dir = os.path.join(current_dir, "chrome_profile")
    
    if not os.path.exists(chrome_profile_dir):
        os.makedirs(chrome_profile_dir)
        
    options.add_argument(f"--user-data-dir={chrome_profile_dir}")
    options.add_argument("--profile-directory=Default")
    
    # 자동화 탐지 우회 옵션들
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    options.add_argument("--window-size=1200,900")
    
    return options

def run_login_setup():
    # 사용자가 전용 브라우저에서 1회 수동 로그인할 수 있도록 지원하는 모드
    print("\n=== [네이버 자동 로그인 세션 생성 모드] ===")
    print("이 단계는 최초 1회만 수행하면 됩니다.")
    print("전용 크롬 브라우저가 열리면, 네이버 로그인을 직접 완료해 주세요.")
    print("로그인을 완료하고 네이버 홈 화면이 나오면 이 창(콘솔)으로 돌아와 Enter를 누르시면 됩니다.\n")
    
    options = get_chrome_options(login_mode=True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get("https://nid.naver.com/nidlogin.login")
        input("네이버 로그인을 성공적으로 마친 후, 이 콘솔 창에 엔터(Enter)를 누르세요...")
        print("로그인 세션이 안전하게 저장되었습니다.")
    except Exception as e:
        print(f"오류가 발생했습니다: {e}")
    finally:
        driver.quit()
        print("브라우저를 닫았습니다.")

def publish_post(naver_id, title, html_content_file):
    # HTML 본문 파일 로드
    if not os.path.exists(html_content_file):
        print(f"Error: HTML content file not found at {html_content_file}")
        sys.exit(1)
        
    with open(html_content_file, 'r', encoding='utf-8') as f:
        html_content = f.read()

    print("Initializing Chrome Webdriver (Using Dedicated Profile)...")
    options = get_chrome_options()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15)
    
    try:
        # 네이버 블로그 글쓰기 페이지 진입
        write_url = f"https://blog.naver.com/{naver_id}?Redirect=Write"
        print(f"Navigating to: {write_url}")
        driver.get(write_url)
        time.sleep(5)
        
        # iframe으로 전환
        try:
            driver.switch_to.frame("mainFrame")
            print("Switched to mainFrame iframe.")
        except Exception:
            print("No mainFrame found, staying in default content.")

        # 1. '이전 쓰던 글이 있습니다.' 작성 취소 팝업 대응
        try:
            cancel_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'se-popup-button-cancel') or contains(., '취소')]")))
            cancel_btn.click()
            print("Cancelled previous post recovery popup.")
            time.sleep(1)
        except Exception:
            print("No previous post popup appeared.")

        # 2. 제목 입력
        print("Locating title field...")
        title_area = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[contains(@placeholder, '제목') or contains(@class, 'se-document-title')]")))
        title_area.click()
        time.sleep(0.5)
        title_area.send_keys(title)
        print(f"Successfully entered title: {title}")
        time.sleep(1)

        # 3. 본문 입력
        print("Locating content field...")
        content_area = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'se-content') or @contenteditable='true']")))
        content_area.click()
        time.sleep(1)
        
        # HTML 직접 주입
        print("Injecting HTML content via Javascript...")
        driver.execute_script("arguments[0].innerHTML = arguments[1];", content_area, html_content)
        time.sleep(2)
        
        # React 상태 업데이트 트리거
        content_area.send_keys(Keys.END)
        content_area.send_keys(" ")
        time.sleep(0.5)
        content_area.send_keys(Keys.BACKSPACE)
        print("HTML content injection completed.")
        time.sleep(2)

        # 4. 발행 버튼 클릭 (테스트를 위해 비활성화하고 본문 작성 상태에서 멈춤)
        print("[TEST MODE] First publish buttons are disabled for safety.")
        print("Please check the opened Chrome browser to see if the title and body are correctly filled.")
        input("Press Enter in this console to close the browser after verification...")
        
    except Exception as e:
        print(f"An error occurred during automation: {e}")
        screenshot_path = "naver_publish_error.png"
        driver.save_screenshot(screenshot_path)
        print(f"Saved error screenshot to: {screenshot_path}")
        
    finally:
        driver.quit()
        print("Browser closed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Publish a post to Naver Blog via Selenium.")
    parser.add_argument("--login", action="store_true", help="Run in login setup mode to save session credentials.")
    parser.add_argument("--id", default="ddong7733770", help="Naver Blog ID (default: ddong7733770)")
    parser.add_argument("--title", help="Blog Post Title")
    parser.add_argument("--file", default="blog_post.html", help="Path to HTML content file")
    
    args = parser.parse_args()
    
    if args.login:
        run_login_setup()
    else:
        if not args.title:
            print("Error: --title is required for publishing a post.")
            sys.exit(1)
        publish_post(args.id, args.title, args.file)
