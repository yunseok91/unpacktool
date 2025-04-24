from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
import time


class NavigationExtractFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)
        self.max_menu_items = 0  # 최대 메뉴 항목 수 추적

    # 여기선 class_name 필드 값 가져와야함
    def execute(
        self, filename, wait_time, server, username=None, password=None, class_name=None
    ):
        self.logger.log("▶️ Navigation 추출 ")
        self.logger.log(f"Time Sleep: {wait_time}")
        QApplication.processEvents()
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(
                f"🟢 Active Sheet Name: {self.file_manager.worksheet.title}","success"
            )
            QApplication.processEvents()

            if not class_name or class_name.strip() == "":
                class_name = "floating-navigation__menu-item"  # 기본값

            self.logger.log(f"🔍 Target Class: {class_name}")
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
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue

                self.logger.log(f"{str(i-4)}번째 Link➡️{str(link)}")
                QApplication.processEvents()
                try:
                    # 페이지 로드
                    driver.get(link)
                    time.sleep(wait_time)  # 페이지 로드 대기
                    # 메모리 
                    if i % 50 == 0:
                        self._clear_memory(driver)
                    
                    # QA 로그인 처리
                    if server == "QA":
                        wait = WebDriverWait(driver, 5)
                        self._handle_qa_login(wait, username, password)
                    # 네비게이션 메뉴 추출
                    self._extract_navigation_menu(i, class_name)

                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}")
                    continue

            # 결과 저장 - 여기서 저장
            QApplication.processEvents()
            self.file_manager.save_results("navigation_anla")
            self.logger.log('작업 완료', "success")
        except Exception as e:
            self.logger.log(f'오류 발생: {str(e)}', "error")
            self.file_manager.save_error_file()

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

    def _clear_memory(self, driver):
        #메모리 관리
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료 🧹", "info")
        except:
            pass

    def _extract_navigation_menu(self, row, class_name):
        # 네비게이션 메뉴 추출#
        driver = self.selenium_manager.driver

        # visible 메뉴 항목만 추출
        menu_items = driver.execute_script(
            f"""
            const items = [];
            
            // 보이는 li 요소만 선택 (사용자 지정 클래스 사용)
            document.querySelectorAll('{class_name}').forEach((li) => {{
                const a = li.querySelector('a');
                if (a && window.getComputedStyle(a).display !== 'none') {{
                    const label = a.getAttribute('an-la') || '';
                    const text = a.innerText.trim() || '';
                    items.push({{label, text}});
                }}
            }});
            
            return items;
            """
        )

        if not menu_items:
            self.logger.log(
                f"❌ 클래스명 '{class_name}'에 해당하는 요소를 찾을 수 없습니다."
            )
            return

        # 추출된 메뉴 항목 수
        num_items = len(menu_items)
        self.logger.log(f"네비게이션 메뉴 항목 수: {num_items}","info")
        QApplication.processEvents()
        base_col_index = 3 # D열부터 시작
        
        # 최대 메뉴 항목 수 업데이트
        if num_items > self.max_menu_items:
            self.max_menu_items = num_items
            # 최대값이 업데이트되면 헤더도 업데이트

            # 기존 헤더 셀의 스타일 참조
            ref_text_cell = self.file_manager.worksheet['D4']
            ref_label_cell = self.file_manager.worksheet['E4']
            #헤더 추가 for문
            for idx in range(self.max_menu_items):
                # idx = 0 각 항목마다 2칸식 이동해야지 D F H 
                text_col = chr(65 + base_col_index + (idx * 2))  # D, F, H, ...
                label_col = chr(65 + base_col_index + (idx * 2) + 1)  # E, G, I, ...
                
                # 헤더 값 설정
                text_cell = self.file_manager.worksheet[f"{text_col}4"]
                label_cell = self.file_manager.worksheet[f"{label_col}4"]
                
                text_cell.value = f"Menu Text {idx+1}"
                label_cell.value = f"Menu Label {idx+1}"
                
                # 스타일 복사 시도
                try:
                    # 텍스트 셀 스타일 복사
                    text_cell.font = ref_text_cell.font.copy()
                    text_cell.fill = ref_text_cell.fill.copy()
                    text_cell.border = ref_text_cell.border.copy()
                    text_cell.alignment = ref_text_cell.alignment.copy()
                    
                    # 라벨 셀 스타일 복사
                    label_cell.font = ref_label_cell.font.copy()
                    label_cell.fill = ref_label_cell.fill.copy()
                    label_cell.border = ref_label_cell.border.copy()
                    label_cell.alignment = ref_label_cell.alignment.copy()
                except:
                    pass  # 스타일 복사 실패해도 계속 진행
        #두 번째 for문 (데이터 저장 for문)
        for idx, item in enumerate(menu_items):
            # 메뉴 텍스트를 홀수 열에, an-la 값을 짝수 열에 저장
            # ASCII : char(65 = A B C D=68
            # idx = 0일 때: chr(65 + 3 + (0 * 2)) = chr(68) = 'D'

            text_col = chr(65 + base_col_index + (idx * 2))  # D, F, H, ...
            label_col = chr(65 + base_col_index + (idx * 2) + 1)  # E, G, I, ...

            # 엑셀에 저장
            self.file_manager.worksheet[f"{text_col}{row}"] = item["text"]
            self.file_manager.worksheet[f"{label_col}{row}"] = item["label"]

            # 로그 출력
            self.logger.log(f"항목 {idx+1}: {item['text']} - {item['label']}","info")
            QApplication.processEvents()