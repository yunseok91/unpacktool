from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication

from utils.logger import Logger
import time

class PageTrackFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget) 

    def execute(self, filename, wait_time, server, username=None, password=None, cookie=None):
        self.logger.log("Page Track 시작", "success")
        self.logger.log(f"Time Sleep : {wait_time}", "info")

        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(f"Active Sheet Name: {self.file_manager.worksheet.title}", "success")
            QApplication.processEvents()
            # 드라이버 초기화 상태 확인
            self.selenium_manager.check_driver_initialized(self.logger)
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(30)  # 타임아웃 증가
            
            # WMC 모드에서 초기 쿠키 설정
            if server == "WMC" and cookie:
                self._setup_wmc_for_domain(driver, cookie)
            self._block_resources_except_js(driver)
            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                self.logger.log("", "separator")  # 구분선 추가
                
                link = self.file_manager.worksheet['C' + str(i)].value
                if not link:
                    continue

                self.logger.log(f"{str(i-4)}번째 Link -> {str(link)}")
                QApplication.processEvents()
                try:
                     # 메모리 관리 - 변경됨: 30개마다 처리 및 WMC 모드 지원
                    if i % 30 == 0:
                        if server == "WMC" and cookie:
                            # WMC 모드에서는 login-token 쿠키 보존
                            self._clear_cookies_except_login(driver)
                        else:
                            # 일반 모드에서는 모든 메모리 정리
                            self._clear_memory(driver)
                        QApplication.processEvents()

                    # 페이지 로드
                    driver.get(link)
                    
                    # 페이지 로드 대기 추가
                    self._wait_for_page_ready(driver, timeout=15)
                    #  QA 서버인 경우 로그인 먼저 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)
                        self.logger.log("QA 서버 로그인 완료")
                        QApplication.processEvents()
                    # WMC 모드에서 에러 페이지 확인 및 처리
                    elif server == "WMC" and cookie:
                        if self._check_for_error_page(driver):
                            self._refresh_wmc_cookie(driver, cookie)
                            driver.refresh()
                            self._wait_for_page_ready(driver, timeout=10)

                    # 에러 페이지 체크
                    if self._check_error_page(i):
                        break

                    # pageTrack 데이터 추출
                    self._extract_page_track(i)

                    self.logger.log("", "separator")  # 구분선 추가
                    QApplication.processEvents()
                except Exception as e:
                    self.logger.log(f'페이지 처리 중 오류: {str(e)}', "error")
                    continue

            # 최종 저장
            self.file_manager.save_results('page_track')
            self.logger.log('작업 완료', "success")
            
        except Exception as e:
            self.logger.log(f'오류 발생: {str(e)}', "error")
            self.file_manager.save_error_file()
        
        finally:
            self.selenium_manager.quit_driver()

    def _block_resources_except_js(self, driver):
        #JS 파일을 제외한 리소스 차단 설정#
        try:
            driver.execute_cdp_cmd('Network.setBlockedURLs', {
                "urls": [
                    "*.jpg", "*.jpeg", "*.png", "*.gif", "*.css", 
                    "*.woff", "*.woff2", "*.analytics.*", 
                    "*.doubleclick.net/*", "*.adobedtm.com/*",
                    "*.google-analytics.com/*", "*.facebook.*",
                    "*.twitter.*", "*.youtube.*"
                    # JS 파일은 차단하지 않음
                ]
            })
            driver.execute_cdp_cmd('Network.enable', {})
            self.logger.log("리소스 차단 설정 완료 (JS 파일 유지)", "info")
        except Exception as e:
            self.logger.log(f"리소스 차단 설정 오류: {str(e)}", "warning")

    # 추가된 메서드: WMC 모드 초기화
    def _setup_wmc_for_domain(self, driver, cookie):
        #실제 도메인에서 WMC 모드 쿠키 설정#
        self.logger.log("WMC 모드 초기화 중...", "info")
        try:
            # 먼저 타겟 도메인으로 이동 (예: samsung.com)
            initial_url = "https://www.samsung.com/sg"
            self.logger.log(f"초기 도메인 접속 중: {initial_url}")
            driver.get(initial_url)
            
            # 페이지 로드 대기
            self._wait_for_page_ready(driver)
            
            # 쿠키 값 형식 처리
            cookie_value = cookie.strip()
            if "login-token=" in cookie_value:
                # "login-token=값" 형식에서 값만 추출
                cookie_value = cookie_value.split("login-token=")[1].split(";")[0].strip()
                
            # 쿠키 설정
            cookie_dict = {
                'name': 'login-token',
                'value': cookie_value,
                'domain': '.samsung.com',  # 도메인에 맞게 수정
                'path': '/'
            }
            
            # 기존 쿠키 삭제 후 새 쿠키 추가
            driver.delete_all_cookies()
            driver.add_cookie(cookie_dict)
            
            # 페이지 새로고침으로 쿠키 적용
            driver.refresh()
            self._wait_for_page_ready(driver)
            
            self.logger.log("WMC 모드 초기화 완료", "success")
            return True
        except Exception as e:
            self.logger.log(f"WMC 모드 초기화 오류: {str(e)}", "error")
            return False

    # 추가된 메서드: 로그인 쿠키 보존 메모리 정리
    def _clear_cookies_except_login(self, driver):
        #login-token을 제외한 모든 쿠키 삭제#
        try:
            # 현재 모든 쿠키 가져오기
            cookies = driver.get_cookies()
            
            # login-token 쿠키만 별도 저장
            login_cookies = [c for c in cookies if c['name'] == 'login-token']
            
            if login_cookies:
                # 쿠키 삭제 전 개수 로깅
                self.logger.log(f"쿠키 정리 전: {len(cookies)}개", "info")
                
                # 모든 쿠키 삭제
                driver.delete_all_cookies()
                
                # login-token 쿠키만 다시 추가
                for cookie in login_cookies:
                    driver.add_cookie(cookie)
                
                # 스토리지 정리 (로그인 관련 항목 제외)
                driver.execute_script("""
                    // 세션 스토리지에서 로그인 관련 키를 제외한 모든 항목 제거
                    let sessionKeys = Object.keys(sessionStorage);
                    for (let i = 0; i < sessionKeys.length; i++) {
                        const key = sessionKeys[i];
                        if (!key.includes('login') && !key.includes('auth') && !key.includes('token')) {
                            sessionStorage.removeItem(key);
                        }
                    }
                    
                    // 로컬 스토리지에서 로그인 관련 키를 제외한 모든 항목 제거
                    let localKeys = Object.keys(localStorage);
                    for (let i = 0; i < localKeys.length; i++) {
                        const key = localKeys[i];
                        if (!key.includes('login') && !key.includes('auth') && !key.includes('token')) {
                            localStorage.removeItem(key);
                        }
                    }
                """)
                
                self.logger.log("불필요한 쿠키 정리 완료 (login-token 유지) 🧹", "info")
            else:
                # login-token이 없으면 쿠키를 유지하고 경고 로그
                self.logger.log("login-token 쿠키가 없습니다. 새로 설정이 필요할 수 있습니다.", "warning")
                
        except Exception as e:
            self.logger.log(f"쿠키 정리 중 오류: {str(e)}", "warning")

    # 추가된 메서드: 페이지 로드 대기
    def _wait_for_page_ready(self, driver, timeout=15):
        #페이지 로드 완료 대기#
        try:
            # 문서 상태 확인
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            return True
        except Exception as e:
            self.logger.log(f"페이지 로드 대기 중 오류: {str(e)}", "warning")
            return False
            
    def _clear_memory(self, driver):
        #메모리 관리
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료", "info")
        except:
            pass

    def _handle_qa_login(self, wait, username, password):
        # QA 서버 로그인 처리#
        try:
            login_elements = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#username, #password")
                )
            )

            if len(login_elements) == 2 and username and password:
                self.selenium_manager.driver.execute_script(
                    """
                    document.querySelector('#username').value = arguments[0];
                    document.querySelector('#password').value = arguments[1];
                    document.querySelector('#submit-button').click();
                    """,
                    username,
                    password,
                )
                time.sleep(1)
        except TimeoutException:
            self.logger.log('로그인 요소를 찾을 수 없습니다.', "error")
            pass

    def _check_error_page(self, row):
       #에러 페이지 확인#
       driver = self.selenium_manager.driver
       if "404 Not Found" in driver.title:
           self.logger.log("Error 404: Page not found", "error")
           self.file_manager.worksheet['F'+str(row)] = "404 Error"
           return True
       if "500 Internal Server Error" in driver.title:
           self.logger.log("Error 500: Internal server error", "error")
           self.file_manager.worksheet['F'+str(row)] = "500 Error"
           return True
       return False
    
    # 추가된 메서드: WMC 쿠키 재설정
    def _refresh_wmc_cookie(self, driver, cookie):
        #WMC 쿠키 재설정#
        try:
            # 쿠키 값 형식 처리
            cookie_value = cookie.strip()
            if "login-token=" in cookie_value:
                # "login-token=값" 형식에서 값만 추출
                cookie_value = cookie_value.split("login-token=")[1].split(";")[0].strip()
                
            # 쿠키 설정
            cookie_dict = {
                'name': 'login-token',
                'value': cookie_value,
                'domain': '.samsung.com',  # 도메인에 맞게 수정
                'path': '/'
            }
            
            # 기존 쿠키 삭제 후 새 쿠키 추가
            driver.delete_all_cookies()
            driver.add_cookie(cookie_dict)
            
            self.logger.log("WMC 쿠키 재설정 완료", "success")
            return True
        except Exception as e:
            self.logger.log(f"WMC 쿠키 재설정 오류: {str(e)}", "error")
            return False

    # 추가된 메서드: 에러 페이지 확인
    def _check_for_error_page(self, driver):
        #에러 페이지 확인#
        try:
            page_title = driver.title
            page_content = driver.page_source[:1000]  # 처음 1000자만 확인
            
            error_indicators = [
                "400 Bad Request", "Error", "Access Denied", 
                "Forbidden", "Unauthorized", "Denied", "Login Required"
            ]
            
            # 타이틀에서 에러 확인
            for indicator in error_indicators:
                if indicator in page_title:
                    self.logger.log(f"에러 페이지 감지: {page_title}", "warning")
                    return True
                    
            # 페이지 내용에서 에러 확인
            for indicator in error_indicators:
                if indicator in page_content:
                    self.logger.log(f"에러 내용 감지: {indicator}", "warning")
                    return True
                    
            return False
        except Exception as e:
            self.logger.log(f"에러 페이지 확인 중 오류: {str(e)}", "warning")
            return False
            
    def _extract_page_track(self, row):
       #페이지 트랙 데이터 추출#
       driver = self.selenium_manager.driver
       js_code = "return window.digitalData;"
       digital_data = driver.execute_script(js_code)
       
       if not digital_data:
           self.logger.log("Error: digitalData not found", "error")
           self.file_manager.worksheet['D'+str(row)] = "digitalData error"
           return

       try:
           page_track = digital_data['page']['pageInfo']['pageTrack']
           if not page_track:
               self.logger.log('pageTrack: 없음', "warning")
               page_track = "pageTrack: 없음"
           else:
               self.logger.log(f"pageTrack: {page_track}", "info")
           
           self.file_manager.worksheet['D'+str(row)] = page_track
           
       except KeyError:
           self.logger.log('pageTrack: key not found', "error")
           self.file_manager.worksheet['D'+str(row)] = "error"