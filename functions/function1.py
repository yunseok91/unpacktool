from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger

import time, os


class HomeRedirectFunction:
    def __init__(self, selenium_manager, file_manager, text_widget, captureCheck=None):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.text_widget = text_widget
        self.logger = Logger(text_widget)
        self.captureCheck = captureCheck

    def execute(self, filename, wait_time, server, username=None, password=None, cookie=None):
        self.logger.log("Home redirect 시작", "info")
        self.logger.log(f"Time Sleep : {wait_time}초")
        QApplication.processEvents()  # UI 업데이트
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(
                f"Active Sheet Name : {self.file_manager.worksheet.title}","success"
            )
            QApplication.processEvents()
            # 드라이버 초기화 상태 확인 - 새로 추가한 함수 호출
            self.selenium_manager.check_driver_initialized(self.logger)
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(30)



            # WMC 모드에서 초기 쿠키 설정
            if server == "WMC" and cookie:
                self._setup_wmc_for_domain(driver, cookie)
            # 리소스 차단
            self._block_resources_except_js(driver)
            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
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
                        QApplication.processEvents()

                    # 페이지 이동 - 기존 방식 그대로 유지
                    driver.execute_script(f"window.location.href = '{link}'")
                    
                    # 페이지 로드 대기 추가
                    self._wait_for_page_ready(driver, timeout=15)
                    #  QA 서버인 경우 로그인 먼저 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)
                        self.logger.log("QA 서버 로그인 완료")
                        QApplication.processEvents()
                    # WMC 모드에서 페이지 로드 후 에러 확인 및 처리
                    elif server == "WMC" and cookie:
                        if self._check_and_handle_errors(driver, cookie):
                            self._refresh_wmc_cookie(driver, cookie)
                            driver.refresh()
                            self._wait_for_page_ready(driver)
                    # 에러 페이지 체크
                    if self._check_error_page(i):
                        break

                    # 스크린샷 체크박스 여부 
                    if self.captureCheck.isChecked():
                        self.captureCheckBox()
                    # URL 저장
                    current = driver.current_url
                    self.file_manager.worksheet["E" + str(i)] = current

                    # 현재 URL에 "unpacked"가 포함되어 있는지 확인
                    if "unpacked" in current.lower():
                        self.file_manager.worksheet["D" + str(i)] = "O"
                        self.logger.log(f"URL 비교 결과: 🟢 Pass (unpacked 포함)")
                        self.logger.log(f"엑셀 URL: {link}")
                        self.logger.log(f"현재 URL: {current}")
                    else:
                        self.file_manager.worksheet["D" + str(i)] = "X"
                        self.logger.log(f"URL 비교 결과: ❌ Fail (unpacked 없음)")
                        self.logger.log(f"엑셀 URL: {link}")
                        self.logger.log(f"현재 URL: {current}")

                    QApplication.processEvents()
                    self.logger.log("", "separator")
                    # 주기적 저장
                    if i % 50 == 0:
                        self.file_manager.save_results("home_redirect")
                        self.logger.log("★★★임시 저장 완료★★★")
                        QApplication.processEvents()

                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}", "error")
                    self.file_manager.worksheet["F" + str(i)] = "Error"
                    QApplication.processEvents()
                    continue

            # 최종 저장
            self.file_manager.save_results("home_redirect")
            self.logger.log("최종 저장 완료", "success")
            QApplication.processEvents()

        except Exception as e:
            self.logger.log(f"최종 저장 중 오류 발생: {e}", "error")
            self.file_manager.save_error_file()
            QApplication.processEvents()

        finally:
            self.selenium_manager.quit_driver()

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
            
    def captureCheckBox(self):
        try:
            driver = self.selenium_manager.driver
            if driver:
                time.sleep(1)
                folder_name = "Screenshot"
                current_directory = os.getcwd()
                new_folder_path = os.path.join(current_directory, folder_name)
                if not os.path.exists(new_folder_path):
                    os.makedirs(new_folder_path)
                
                # 올바른 드라이버 변수 사용
                site_code = driver.find_element(By.CSS_SELECTOR,"meta[name='sitecode']").get_attribute("content")
                timestamp = time.strftime("%Y%m%d_%H%M%S")

                file_name = f"capture_{site_code}_{timestamp}.png"
                file_path = os.path.join(new_folder_path, file_name)

                # 올바른 드라이버 변수 사용
                driver.save_screenshot(file_path)
                self.logger.log(f"캡쳐완료 {file_name}","success")
                QApplication.processEvents()
                time.sleep(0.5)
        except NoSuchElementException:
            site_code = "unknown"  # 요소를 찾을 수 없는 경우 기본값
            self.logger.log("사이트 코드를 찾을 수 없습니다", "warning")
        except Exception as e:
            self.logger.log(f"캡쳐 에러: {e}", "error")
            
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
    def _setup_wmc_for_domain(self, driver, cookie):
        #WMC 모드 초기 설정#
        self.logger.log("WMC 모드 초기화 중...", "info")
        try:
            # 먼저 타겟 도메인으로 이동 (예: samsung.com)
            initial_url = "https://www.samsung.com/sg/"
            self.logger.log(f"초기 도메인 접속 중: {initial_url}")
            driver.get(initial_url)
            
            # 페이지 로드 대기
            self._wait_for_page_ready(driver)
            
            # 로그인 토큰 쿠키 설정 (직접 add_cookie 사용)
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
        

    def _apply_wmc_cookie(self, driver, cookie_value):
        #WMC 쿠키 적용#
        try:
            # cookie_value가 JavaScript 코드 형식인지 확인
            if "document.cookie=" in cookie_value:
                # JavaScript 코드에서 쿠키 값 추출
                cookie_content = cookie_value.split('document.cookie="')[1].split('";')[0]
            elif "login-token=" in cookie_value:
                # 이미 "login-token=값" 형식인 경우 그대로 사용
                cookie_content = cookie_value
            else:
                # 값만 입력한 경우 "login-token=" 프리픽스 추가
                cookie_content = f"login-token={cookie_value}; path=/"

            # 쿠키가 올바른 형식인지 확인
            if not "login-token=" in cookie_content:
                self.logger.log("쿠키 형식이 잘못되었습니다. 'login-token=' 형식으로 입력해주세요.", "error")
                return False

            # 쿠키 설정을 JavaScript로 직접 실행
            driver.execute_script(f"document.cookie='{cookie_content}'")
            self.logger.log(f"WMC 쿠키 설정 완료:{cookie_content}", "success")
            return True
        except Exception as e:
            self.logger.log(f"쿠키 설정 오류: {str(e)}", "error")
            return False

    def _check_wmc_login(self, driver):
        #WMC 로그인 상태 확인#
        try:
            # 쿠키 확인
            cookies = driver.get_cookies()
            for cookie in cookies:
                if cookie['name'] == 'login-token':
                    return True
            return False
        except:
            return False

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

    def _clear_memory(self, driver):
        #일반 메모리 정리 (모든 쿠키와 스토리지 삭제)#
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료 🧹", "info")
        except Exception as e:
            self.logger.log(f"메모리 정리 오류: {str(e)}", "warning")
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
    def _check_and_handle_errors(self, driver, cookie):
        #400 및 접근 거부 에러 감지 및 처리#
        try:
            # 페이지 타이틀이나 내용에서 오류 감지
            page_title = driver.title
            
            # 400 에러 또는 접근 거부 감지
            error_indicators = [
                "400", "Bad Request", "Error", "Access Denied", 
                "Forbidden", "Unauthorized", "Denied"
            ]
            
            for indicator in error_indicators:
                if indicator in page_title:
                    self.logger.log(f"에러 감지: {page_title}", "warning")
                    
                    # 모든 쿠키 삭제 후 로그인 쿠키만 다시 설정
                    driver.delete_all_cookies()
                    
                    if cookie:
                        self._apply_wmc_cookie(driver, cookie)
                        
                    # 페이지 새로고침
                    driver.refresh()
                    self.logger.log("쿠키 재설정 및 페이지 새로고침 완료", "info")
                    return True
                    
            return False  # 에러 없음
        except Exception as e:
            self.logger.log(f"에러 체크 중 오류: {str(e)}", "warning")
            return False

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