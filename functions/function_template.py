from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
import time


class NewFunction:
    #새로운 기능을 위한 클래스 템플릿 
    def __init__(self, selenium_manager, file_manager, text_widget):
        #초기화 메서드#
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)
        # 필요한 추가 변수 초기화

    def execute(self, filename, wait_time, server, username=None, password=None, extra_param=None):
        #기능 메서드
            # filename: 엑셀 파일 경로
            # wait_time: 페이지 로드 대기 시간
            # server: 서버 타입 ('WWW' 또는 'QA')
            # username: QA 서버 로그인 아이디 (선택적)
            # password: QA 서버 로그인 비밀번호 (선택적)
            # extra_param: 추가 매개변수 (필요한 경우)

        self.logger.log("▶️ 새 기능 실행")  # 기능 이름 수정
        self.logger.log(f"Time Sleep: {wait_time}")
        QApplication.processEvents()
        
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(f"🟢 Active Sheet Name: {self.file_manager.worksheet.title}")
            QApplication.processEvents()
            
            # 드라이버 초기화 상태 확인
            self.selenium_manager.check_driver_initialized(self.logger)
            
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(15)
            self.selenium_manager.block_resources()
            
            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                self.logger.log("─" * 50)  # 구분선
                QApplication.processEvents()
                
                # C열에서 URL 가져오기 (필요에 따라 열 변경)
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue
                
                self.logger.log(f"{str(i-4)}번째 Link➡️{str(link)}")
                QApplication.processEvents()
                
                try:
                    # 페이지 로드
                    driver.get(link)
                    time.sleep(wait_time)  # 페이지 로드 대기
                    
                    # 메모리 관리
                    if i % 50 == 0:
                        self._clear_memory(driver)
                    
                    # QA 로그인 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)
                    
                    # 데이터 추출 - 이 부분을 기능에 맞게 구현
                    self._extract_data(i, extra_param)
                    
                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}")
                    continue
            
            # 결과 저장 - 기능에 맞는 폴더 이름으로 변경
            self.file_manager.save_results("new_function")  
            self.logger.log('작업 완료', "success")
            
        except Exception as e:
            self.logger.log(f'오류 발생: {str(e)}', "error")
            self.file_manager.save_error_file()
        
        finally:
            # 드라이버 종료
            if self.selenium_manager and self.selenium_manager.driver:
                self.selenium_manager.quit_driver()
    
    def _handle_qa_login(self, wait, username, password):
        #QA 서버 로그인 처리#
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
        #메모리 관리#
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료 🧹", "info")
        except:
            pass
    
    def _extract_data(self, row, extra_param=None):

        driver = self.selenium_manager.driver
        
        # TODO: 데이터 추출 로직 구현
        # JavaScript를 사용하여 원하는 데이터 추출
        extracted_data = driver.execute_script("""   """)
        
        # 데이터가 없는 경우 처리
        if not extracted_data:
            self.logger.log("❌ 데이터를 찾을 수 없습니다.")
            return
        
        # 로그에 추출된 데이터 수 출력
        self.logger.log(f"추출된 데이터 항목 수: {len(extracted_data)}")
        QApplication.processEvents()
        
        # 엑셀에 데이터 저장 (D열부터)
        col_index = 3  # D열부터 시작
    
        
