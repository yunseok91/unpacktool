# 기본 설정값들을 담는 파일
import os

# 프로젝트 루트 경로
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Webdriver 관련 설정
WEBDRIVER_PATH = {
    'edge': os.path.join(ROOT_DIR, 'webdriver_manager', 'edgedriver_133', 'msedgedriver.exe'),
    'chrome': os.path.join(ROOT_DIR, 'webdriver_manager', 'chromedriver_133', 'chromedriver.exe')
}

# UI 설정
UI_FILE_PATH = os.path.join(ROOT_DIR, 'resources', 'unpack.ui')
WINDOW_SIZE = (684, 640)

# Selenium 설정
# 브라우저 설정
BROWSER_USER_AGENTS = {
    'edge': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
    'chrome': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36'
}
PAGE_LOAD_TIMEOUT = 15
SCRIPT_TIMEOUT = 10

# 로그인 설정
LOGIN_TIMEOUT = 5
LOGIN_SELECTORS = {
    'username': '#username',
    'password': '#password',
    'submit': '#submit-button'
}

# 파일 저장 관련
RESULTS_FOLDERS = {
    'home_redirect': 'Home_Redirect',
    'digital_data': 'digitalData',
    'page_track': 'PageTrack',
    'navigation_anla': 'navigation',  
    'KV_cta': 'KV_CTA',  
    'error': 'error',
    'screenshot': 'Screenshot'
}