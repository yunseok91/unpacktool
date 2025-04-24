from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.service import Service as EdgeService 
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import SessionNotCreatedException

from webdriver_manager.chrome import ChromeDriverManager

from .config import WEBDRIVER_PATH, BROWSER_USER_AGENTS

class SeleniumManager:
    def __init__(self):
        self.driver = None

    def initialize_driver(self, browser_type='chrome', headless=False, use_local_driver=True):
        try:
            # 이전 드라이버가 있으면 종료
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None
            
            if browser_type == 'chrome':  # 기본값이므로 먼저 체크
                options = ChromeOptions()
                if use_local_driver:
                    service = ChromeService(WEBDRIVER_PATH['chrome'])
                    print(f"로컬 크롬 드라이브 :,{WEBDRIVER_PATH['chrome']}")
                else:
                    service = ChromeService(ChromeDriverManager().install())
                    print("최신 크롬 드라이브")
            else:  # edge
                options = EdgeOptions()
                service = EdgeService(WEBDRIVER_PATH['edge'])
                print(f"로컬 엣지 드라이브 :,{WEBDRIVER_PATH['edge']}")
            
                options.use_chromium = True
            
            # User-Agent 설정
            user_agent = BROWSER_USER_AGENTS.get(browser_type, BROWSER_USER_AGENTS['chrome'])
            print(f"Using user agent: {user_agent}")
            # 공통 옵션 설정 
            options.add_argument("--start-maximized")
            options.add_argument("--window-size=1920x1080")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-dev-tools")
            options.add_argument("--disable-logging")
            options.add_argument("--disable-notifications")
            options.add_argument("--incognito")
            options.add_argument("--compress")
            options.add_argument("--enable-features=OptimizeHeaders")
            options.add_argument("--user-agent=" + user_agent)
        # Edge 전용 옵션은 Edge 브라우저에만 적용
            if browser_type == 'edge':
                options.use_chromium = True
            if headless:
                options.add_argument('--headless')

            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_experimental_option("detach", False)
            # 브라우저 타입에 따라 드라이버 초기화
            if browser_type == 'chrome':
                self.driver = webdriver.Chrome(service=service, options=options)
            else:
                self.driver = webdriver.Edge(service=service, options=options)

            return self.driver

        except SessionNotCreatedException  as e:
            print(f"SessionNotCreatedException: {str(e)}")
            raise Exception(f"드라이버 초기화 실패: 브라우저 버전이 드라이버와 맞지 않습니다 {str(e)}")
        except Exception as e:
            print(f" Exception: {str(e)}")
            raise Exception(f"드라이버 초기화 실패: {str(e)}")
    def block_resources(self):
       #리소스 차단 설정#
       if self.driver:
           try:
               self.driver.execute_cdp_cmd('Network.setBlockedURLs', {
                   "urls": [
                       "*.jpg", "*.jpeg", "*.png", "*.gif", "*.css", 
                       "*.woff", "*.woff2", "*.analytics.*", 
                       "*.doubleclick.net/*", "*.adobedtm.com/*",
                       "*.google-analytics.com/*", "*.facebook.*",
                       "*.twitter.*", "*.youtube.*"
                        #파일은 차단하지 않음 "*.js" 파일은 차단하지 않음
                   ]
               })
               self.driver.execute_cdp_cmd('Network.enable', {})
           except Exception:
               # CDP 명령 실패시 드라이버 재생성
               self.quit_driver()
               self.initialize_driver()

    def quit_driver(self):
        if self.driver:
            try:
                self.driver.execute_cdp_cmd('Network.disable', {})
            except:
                pass
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
    def check_driver_initialized(self, logger=None):
        #드라이버 초기화 상태를 확인하는 함수
        #Returns:bool: 드라이버가 초기화되어 있으면 True
        #Raises: Exception: 드라이버가 초기화되지 않았을 경우 예외 발생

        # 드라이버가 초기화되어 있는지 확인
        if self.driver is None:
            error_msg = "❌ 드라이버가 초기화되지 않았습니다"
            if logger:
                logger.log(error_msg)
            raise Exception("드라이버가 초기화되지 않았습니다")
        
        # 모든 검사를 통과하면 True 반환
        return True