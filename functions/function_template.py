from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
import time
import traceback

class NewFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)
        self.max_extracted_items = 0  # 최대 추출 항목 수 추적

    def execute(self, filename, wait_time, server, username=None, password=None, extra_param=None, cookie=None):
        self.logger.log("새 기능 실행")
        self.logger.log(f"Time Sleep: {wait_time}")
        QApplication.processEvents()
        
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
            self._block_resources(driver)

            # WMC 모드에서 초기 쿠키 설정
            if server == "WMC" and cookie:
                self._setup_wmc_for_domain(driver, cookie)

            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                self.logger.log("─" * 50)  # 구분선
                QApplication.processEvents()
                
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue
                
                self.logger.log(f"{str(i-4)}번째 Link -> {str(link)}")
                QApplication.processEvents()
                
                try:
                    # 메모리 관리 - 30개마다 처리
                    if i % 30 == 0:
                        if server == "WMC" and cookie:
                            # WMC 모드에서는 login-token 쿠키 보존
                            self._clear_cookies_except_login(driver)
                        else:
                            # 일반 모드에서는 모든 메모리 정리
                            self._clear_memory(driver)
                    
                    # 페이지 로드
                    try:
                        driver.get(link)
                    except Exception as load_e:
                        self.logger.log(f"페이지 로드 중 오류: {str(load_e)}", "error")
                        continue
                    
                    time.sleep(wait_time)  # 페이지 로드 대기
                    
                    # QA 서버인 경우 로그인 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)
                        self.logger.log("QA 서버 로그인 완료")
                        QApplication.processEvents()
                    elif server == "WMC" and cookie:
                        # WMC 모드에서 에러 페이지 확인 및 처리
                        if self._check_for_error_page(driver):
                            self._refresh_wmc_cookie(driver, cookie)
                            driver.refresh()
                            time.sleep(wait_time)
                    
                    # 데이터 추출
                    self._extract_data(i, extra_param)
                    self.logger.log("", "separator")  # 구분선 추가
                    QApplication.processEvents()
                
                except Exception as page_e:
                    self.logger.log(f"페이지 처리 중 오류: {str(page_e)}", "error")
                    self.logger.log(traceback.format_exc(), "error")
                    continue
            
            # 결과 저장
            self.file_manager.save_results("new_function")
            self.logger.log('작업 완료', "success")
        
        except Exception as e:
            self.logger.log(f'오류 발생: {str(e)}', "error")
            self.logger.log(traceback.format_exc(), "error")
            self.file_manager.save_error_file()
        
        finally:
            # 드라이버 종료
            if self.selenium_manager and self.selenium_manager.driver:
                try:
                    self.selenium_manager.quit_driver()
                except Exception as quit_e:
                    self.logger.log(f"드라이버 종료 중 오류: {str(quit_e)}", "error")

    def _handle_qa_login(self, wait, username, password):
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

    def _clear_memory(self, driver):
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료", "info")
        except Exception as e:
            self.logger.log(f"메모리 정리 중 오류: {str(e)}", "warning")

    def _extract_data(self, row, extra_param=None):
        driver = self.selenium_manager.driver

        try:
            # TODO: 데이터 추출 JavaScript 코드 작성
            extracted_data = driver.execute_script("""
                const results = [];
                // 데이터 추출 로직 구현
                return results;
            """)

            # 추출된 데이터 없을 경우 처리
            if not extracted_data:
                self.logger.log("데이터를 찾을 수 없습니다.", "warning")
                return

            # 최대 추출 항목 수 업데이트
            if len(extracted_data) > self.max_extracted_items:
                self.max_extracted_items = len(extracted_data)
                # 필요시 엑셀 헤더 업데이트 로직 추가

            # 로그에 추출된 데이터 수 출력
            self.logger.log(f"추출된 데이터 항목 수: {len(extracted_data)}", "info")

            # 엑셀에 데이터 저장 (D열부터)
            base_col_index = 3  # D열부터 시작

            # 데이터 저장 로직
            for idx, item in enumerate(extracted_data):
                text_col = chr(65 + base_col_index + (idx * 2))  # D, F, H, ...
                label_col = chr(65 + base_col_index + (idx * 2) + 1)  # E, G, I, ...

                # 엑셀에 저장
                self.file_manager.worksheet[f"{text_col}{row}"] = item.get('text', '')
                self.file_manager.worksheet[f"{label_col}{row}"] = item.get('label', '')

                # 로그 출력
                self.logger.log(f"항목 {idx+1}: {item.get('text', '')} - {item.get('label', '')}", "info")

        except Exception as e:
            self.logger.log(f"데이터 추출 중 오류: {str(e)}", "error")
            self.logger.log(traceback.format_exc(), "error")

    def _block_resources(self, driver):
        try:
            driver.execute_cdp_cmd('Network.setBlockedURLs', {
                "urls": [
                    "*.jpg", "*.jpeg", "*.png", "*.gif", "*.css", 
                    "*.woff", "*.woff2", "*.analytics.*", 
                    "*.doubleclick.net/*", "*.adobedtm.com/*",
                    "*.google-analytics.com/*", "*.facebook.*",
                    "*.twitter.*", "*.youtube.*"
                ]
            })
            driver.execute_cdp_cmd('Network.enable', {})
            self.logger.log("리소스 차단 설정 완료", "info")
        except Exception as e:
            self.logger.log(f"리소스 차단 설정 오류: {str(e)}", "warning")

    def _setup_wmc_for_domain(self, driver, cookie):
        self.logger.log("WMC 모드 초기화 중...", "info")
        try:
            initial_url = "https://www.samsung.com/sg"
            self.logger.log(f"초기 도메인 접속 중: {initial_url}")
            driver.get(initial_url)
            
            # 페이지 로드 대기
            self._wait_for_page_ready(driver)
            
            # 쿠키 값 형식 처리
            cookie_value = cookie.strip()
            if "login-token=" in cookie_value:
                cookie_value = cookie_value.split("login-token=")[1].split(";")[0].strip()
                
            # 쿠키 설정
            cookie_dict = {
                'name': 'login-token',
                'value': cookie_value,
                'domain': '.samsung.com',
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

    def _wait_for_page_ready(self, driver, timeout=15):
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            return True
        except Exception as e:
            self.logger.log(f"페이지 로드 대기 중 오류: {str(e)}", "warning")
            return False

    def _clear_cookies_except_login(self, driver):
        try:
            cookies = driver.get_cookies()
            login_cookies = [c for c in cookies if c['name'] == 'login-token']
            
            if login_cookies:
                self.logger.log(f"쿠키 정리 전: {len(cookies)}개", "info")
                
                driver.delete_all_cookies()
                
                for cookie in login_cookies:
                    driver.add_cookie(cookie)
                
                self.logger.log("불필요한 쿠키 정리 완료 (login-token 유지)", "info")
            else:
                self.logger.log("login-token 쿠키가 없습니다.", "warning")
                
        except Exception as e:
            self.logger.log(f"쿠키 정리 중 오류: {str(e)}", "warning")

    def _refresh_wmc_cookie(self, driver, cookie):
        try:
            cookie_value = cookie.strip()
            if "login-token=" in cookie_value:
                cookie_value = cookie_value.split("login-token=")[1].split(";")[0].strip()
                
            cookie_dict = {
                'name': 'login-token',
                'value': cookie_value,
                'domain': '.samsung.com',
                'path': '/'
            }
            
            driver.delete_all_cookies()
            driver.add_cookie(cookie_dict)
            
            self.logger.log("WMC 쿠키 재설정 완료", "success")
            return True
        except Exception as e:
            self.logger.log(f"WMC 쿠키 재설정 오류: {str(e)}", "error")
            return False

    def _check_for_error_page(self, driver):
        try:
            page_title = driver.title
            page_content = driver.page_source[:1000]
            
            error_indicators = [
                "400 Bad Request", "Error", "Access Denied", 
                "Forbidden", "Unauthorized", "Denied", "Login Required"
            ]
            
            for indicator in error_indicators:
                if indicator in page_title or indicator in page_content:
                    self.logger.log(f"에러 페이지 감지: {indicator}", "warning")
                    return True
                    
            return False
        except Exception as e:
            self.logger.log(f"에러 페이지 확인 중 오류: {str(e)}", "warning")
            return False