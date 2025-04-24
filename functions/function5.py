from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
import time


class KVFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)

    def execute(self, filename, wait_time, server, username=None, password=None):
        self.logger.log("▶️ KV 실행")
        self.logger.log(f"Time Sleep: {wait_time}")
        QApplication.processEvents()
        
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(f"🟢 Active Sheet Name: {self.file_manager.worksheet.title}","success")
            QApplication.processEvents()
            
            # 드라이버 초기화 상태 확인
            self.selenium_manager.check_driver_initialized(self.logger)
            
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(15)
            self.selenium_manager.block_resources()
            
            # URL 처리
            for i in range(9, self.file_manager.worksheet.max_row + 1):
                self.logger.log("─" * 50)  # 구분선
                QApplication.processEvents()
                
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue
                
                self.logger.log(f"{str(i-8)}번째 Link➡️{str(link)}")
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
                    
                    # KV 데이터 추출 - 여기에 KV 관련 기능 구현
                    self._extract_kv_data(i)
                    
                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}")
                    continue
            
            # 결과 저장
            self.file_manager.save_results("KV_cta")
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
    
    def _extract_kv_data(self, row):
        #데이터 추출
        driver = self.selenium_manager.driver

        try:
            # JS를 사용한 데이터 추출 예시
            kv_data = driver.execute_script("""
                const results = [];
                const slide = document.querySelector('.home-kv-carousel__slide[data-swiper-slide-index="0"]');
                if (slide) {
                    slide.querySelectorAll(".home-kv-carousel__cta-wrap a").forEach((link, index) => {
                        results.push({
                            index: index + 1,
                            anCa: link.getAttribute("an-ca") || "",
                            anLa: link.getAttribute("an-la") || "",
                            text: link.textContent.trim()
                        });
                    });
                }
                return results;
            """)
            
            # 추출한 데이터가 없을 경우
            if not kv_data or len(kv_data) == 0:
                self.logger.log("❌ KV 데이터를 찾을 수 없습니다.")
                return
            # 결과 로깅 및 엑셀에 저장
            self.logger.log(f"KV 슬라이드에서 {len(kv_data)}개의 CTA 버튼 발견","info")

            # CTA 버튼 정보 저장
            for cta in kv_data:
                idx = cta['index']
                anLa = cta['anLa']
                anCa = cta['anCa']
                text = cta['text']
                
                self.logger.log(f"CTA {idx}: {text}")
                self.logger.log(f"  an-ca: {anCa}")
                self.logger.log(f"  an-la: {anLa}")
                self.logger.log("-" * 40)
                
                # 엑셀에 저장
                try:
                    if idx == 1:  # 첫 번째 CTA
                        self.file_manager.worksheet[f"E{row}"] = anCa
                        self.file_manager.worksheet[f"F{row}"] = anLa
                    elif idx == 2:  # 두 번째 CTA
                        self.file_manager.worksheet[f"G{row}"] = anCa
                        self.file_manager.worksheet[f"H{row}"] = anLa
                except Exception as e:
                    self.logger.log(f"엑셀 저장 중 오류: {str(e)}")
            
            self.logger.log(f"행 {row}에 CTA 정보 저장 완료")

        except Exception as e:
            self.logger.log(f"KV 데이터 추출 중 오류: {str(e)}","error")
            self.file_manager.save_error_file()
