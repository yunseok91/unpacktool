from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from PyQt5.QtWidgets import QApplication

from utils.logger import Logger
import datetime
import time
import os


class DigitalDataFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)

    def execute(self, filename, wait_time, server, username=None, password=None):
        self.logger.log("▶️ Digital Data")
        self.logger.log(f"Time Sleep : {wait_time}초")
        
        start_time = datetime.datetime.now()
        ymds = start_time.strftime("%Y/%m/%d/%H/%M/%S")
        self.logger.log("시작: " + str(ymds))
        QApplication.processEvents()
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(f"Active Sheet Name: {self.file_manager.worksheet.title}초","success")
            # 드라이버 초기화 상    태 확인 - 새로 추가한 함수 호출
            self.selenium_manager.check_driver_initialized(self.logger)
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(20)
            driver.set_script_timeout(20)
            self.selenium_manager.block_resources()

            # URL 처리
            for i in range(5, self.file_manager.worksheet.max_row + 1):
                self.logger.log("", "separator")  # 구분선 추가

                link = self.file_manager.worksheet["E" + str(i)].value
                if not link:
                    continue

                self.logger.log(f"{str(i-4)}번째 Link➡️{str(link)}")

                # 최대 3번 재시도
                max_retries = 3
                retry_count = 0

                while retry_count < max_retries:
                    try:
                        # 메모리 관리
                        if i % 30 == 0:
                            self._clear_memory(driver)

                        # 페이지 이동
                        driver.get(link)

                        # QA 로그인 처리
                        if server == "QA":
                            wait = WebDriverWait(driver, 5)
                            self._handle_qa_login(wait, username, password)

                        # 에러 페이지 체크
                        if self._check_error_page(driver, i):
                            break  # while 루프 탈출하고 다음 URL로 이동

                        # digitalData 추출 및 처리
                        digital_data = self._extract_digital_data(driver)
                        if not digital_data:
                            self.logger.log("❌ Error: digitalData not found", "error")
                            self.file_manager.worksheet["F" + str(i)] = "key Unknown"
                            break

                        # 데이터 처리
                        self._process_digital_data(digital_data, i)
                        break

                    except TimeoutException:
                        retry_count += 1
                        self.logger.log(
                            f"재시도 {retry_count}/{max_retries}...", "warning"
                        )
                        if retry_count == max_retries:
                            self._handle_timeout_error(i)
                            break # 최대 재시도 횟수 초과 시 다음 URL로 이동
                    except Exception as e:
                        self.logger.log(f"페이지 처리 중 오류: {str(e)}", "error")
                        pass

            # 최종 결과 저장
            self.file_manager.save_results('digital_data')
            self.logger.log('작업 완료', "success")
            

        except Exception as e:
            self.logger.log(f"오류 발생: {str(e)}", "error")
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

    def _check_error_page(self, driver, row):
        #에러 페이지 체크#
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

    def _extract_digital_data(self, driver):
        #digitalData 추출#
        script = """
            let maxAttempts = 20;
            let attempt = 0;
            let checkData = function() {
                if (window.digitalData) {
                    return window.digitalData;
                }
                if (attempt >= maxAttempts) {
                    return null;
                }
                attempt++;
                return undefined;
            };
            return checkData();
        """
        wait = WebDriverWait(driver, 10) #최대 10초 동안 script를 실행하며 반복 시도

        return wait.until(lambda x: x.execute_script(script))

    def _process_digital_data(self, digital_data, row):
        #digitalData 처리 및 저장#
        page_info = digital_data.get("page", {}).get("pageInfo", {})
        pathIndicator = digital_data.get("page", {}).get("pathIndicator", {})
        product_info = digital_data.get("product", {})

        # pageName 처리
        page_name = self._get_page_name(page_info, pathIndicator)

        # 데이터 매핑 및 저장
        data_mapping = {
            "site_code": (page_info.get("siteCode"), "site_code"),
            "page_name": (page_name, "F"),
            "site_section": (page_info.get("siteSection"), "G"),
            "page_track": (page_info.get("pageTrack"), "H"),
            "category": (product_info.get("category"), "I"),
            "model_code": (product_info.get("model_code"), "J"),
            "model_name": (product_info.get("model_name"), "K"),
        }

        for key, (value, col) in data_mapping.items():
            if value == "" or value is None:
                self.logger.log(f"{key}: 없음", "warning")
                value = f"{key}: 없음"
            else:
                self.logger.log(f"{key}: {value}", "info")

            if col in "FGHIJK":
                self.file_manager.worksheet[col + str(row)] = value

    def _get_page_name(self, page_info, pathIndicator):
        #pageName 추출 및 처리#
        try:
            if "pageName" in page_info:
                page_name = page_info["pageName"]

                if not page_name or page_name.strip() == "":
                    depth_values = []
                    for depth in range(2, 6):
                        depth_key = f"depth_{depth}"
                        depth_value = pathIndicator.get(depth_key)
                        if (
                            depth_value
                            and isinstance(depth_value, str)
                            and depth_value.strip()
                        ):
                            depth_values.append(depth_value.strip())

                    if depth_values:
                        site_code = page_info.get("siteCode", "")
                        page_name = f"{site_code}:" + ":".join(depth_values)
                        self.logger.log(f"pathIndicator: {page_name}", "info")
                    else:
                        page_name = "pathIndicator 없음"
                        self.logger.log("pageName: pathIndicator도 없음", "error")
                else:
                    self.logger.log(f"pageName from pageInfo: {page_name}", "info")
            else:
                page_name = "pageName 없음"
                self.logger.log("pageName: 속성 없음", "error")

        except Exception as e:
            page_name = "pagename error"
            self.logger.log(f"pageName 처리 중 오류: {str(e)}", "error")

        return page_name

    def _save_temp_file(self, ymds):
        #임시 파일 저장#
        try:
            temp_save_path = os.path.join(
                os.getcwd(), "digitalData", f'temp_{ymds.replace("/", "-")}.xlsx'
            )
            self.file_manager.save_results(temp_save_path)
            self.logger.log("임시 저장 완료", "success")
        except Exception as e:
            self.logger.log(f"임시 저장 중 오류: {e}", "error")

    def _handle_timeout_error(self, row):
        #타임아웃 에러 처리#
        self.logger.log("최대 재시도 횟수 초과", "error")
        for col in "FGHIJK":
            self.file_manager.worksheet[col + str(row)] = "Timeout Error"

    def _reconnect_driver(self):
        #드라이버 재연결#
        if self.selenium_driver:
            try:
                self.selenium_driver.quit()
            except:
                pass
        self.selenium_driver = None
        self.seleniumFunction()
