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

    def execute(self, filename, wait_time, server, username=None, password=None):
        self.logger.log("▶️Page Track 시작", "success")
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
            self.selenium_manager.block_resources()  # block_resources 추가

            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                self.logger.log("", "separator")  # 구분선 추가
                
                link = self.file_manager.worksheet['C' + str(i)].value
                if not link:
                    continue

                self.logger.log(f"{str(i-4)}번째 Link➡️{str(link)}")
                QApplication.processEvents()
                try:
                    # 페이지 로드
                    driver.get(link)
                    
                    # 메모리 관리
                    if i % 50 == 0:
                        self._clear_memory(driver)

                    time.sleep(wait_time)

                    # QA 로그인 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)

                    # 에러 페이지 체크
                    if self._check_error_page(i):
                        continue

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

    def _clear_memory(self, driver):
        #메모리 관리
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료 🧹", "info")
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