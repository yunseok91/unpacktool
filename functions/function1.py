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

    def execute(self, filename, wait_time, server, username=None, password=None):
        self.logger.log("▶️ Home redirecta")
        self.logger.log(f"Time Sleep : {wait_time}초")
        QApplication.processEvents()  # UI 업데이트
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(
                f"🟢Active Sheet Name : {self.file_manager.worksheet.title}","success"
            )
            QApplication.processEvents()
            # 드라이버 초기화 상태 확인 - 새로 추가한 함수 호출
            self.selenium_manager.check_driver_initialized(self.logger)
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(30)
            # 리소스 차단
            self.block_resources_home()

            #  QA 서버인 경우 로그인 먼저 처리
            if server == "QA":
                wait = WebDriverWait(driver, 5)
                self._handle_qa_login(wait, username, password)
                self.logger.log("QA 서버 로그인 완료")
                QApplication.processEvents()

            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue
                self.logger.log(f"{str(i-4)}번째 Link ➡️ {str(link)}")
                QApplication.processEvents()

                try:
                    # 페이지 로드를 기다리지 않고 즉시 다음 작업 실행 가능 성능 최적화를 위해 JavaScript 방식
                    driver.execute_script(f"window.location.href = '{link}'")

                    # 메모리 관리
                    if i % 30 == 0:
                        driver.execute_script("window.sessionStorage.clear();")
                        driver.execute_script("window.localStorage.clear();")
                        driver.delete_all_cookies()
                        self.logger.log("🧹 메모리 정리 완료")
                        QApplication.processEvents()

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
            
    def block_resources_home(self):
        # 리소스 차단 설정
        driver = self.selenium_manager.driver  
        if driver:
            try:
                driver.execute_cdp_cmd('Network.setBlockedURLs', {
                    "urls": [
                        "*.woff", "*.woff2", "*.analytics.*", 
                        "*.doubleclick.net/*", "*.adobedtm.com/*",
                        "*.google-analytics.com/*", "*.facebook.*",
                        "*.twitter.*", "*.youtube.*"
                    ]
                })
                # 올바른 드라이버 변수 사용
                driver.execute_cdp_cmd('Network.enable', {})
            except Exception as e:
                self.logger.log(f"리소스 차단 설정 오류: {e}", "error")
                # CDP 명령 실패시 드라이버 재생성
                self.selenium_manager.quit_driver()
                self.selenium_manager.initialize_driver()