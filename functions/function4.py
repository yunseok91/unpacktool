from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
from datetime import datetime
import time


class NavigationExtractFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)
        self.max_menu_items = 0  # 최대 메뉴 항목 수 추적

    # 여기선 class_name 필드 값 가져와야함
    def execute(self, filename, wait_time, server, username=None, password=None, class_name=None, cookie=None):
        start_time = datetime.now()
        self.logger.log(f"🚀 Navigation 시작: {start_time.strftime('%Y년 %m월 %d일 %H시 %M분 %S초')}", "info")
        self.logger.log(f"Time Sleep: {wait_time}")
        QApplication.processEvents()
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(
                f"Active Sheet Name: {self.file_manager.worksheet.title}","success"
            )
            QApplication.processEvents()

            if not class_name or class_name.strip() == "":
                class_name = ".floating-navigation__menu-item"  # 기본값

            self.logger.log(f" Target Class: {class_name}")
            QApplication.processEvents()
            # 드라이버 초기화 상태 확인 
            self.selenium_manager.check_driver_initialized(self.logger)
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(30)  # 타임아웃 증가

            self._block_resources_except_js(driver)

            if server == "WMC" and cookie:
                # WMC 모드 쿠키 설정
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
                    if i % 30 == 0:
                        if server == "WMC" and cookie:
                            # WMC 모드에서는 login-token 쿠키 보존
                            self._clear_cookies_except_login(driver)
                        else:
                            # 일반 모드에서는 모든 메모리 정리
                            self._clear_memory(driver)
                    # 페이지 로드
                    driver.get(link)
                    time.sleep(wait_time)  # 페이지 로드 대기
                    # QA 로그인 처리
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
                            
                    # 네비게이션 메뉴 추출
                    self.logger.log("CLASS NAME:",class_name)
                    self._extract_navigation_menu(i, class_name)
                    self.logger.log("", "separator")  # 구분선 추가
                    QApplication.processEvents()
                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}")
                    continue

            # 결과 저장 - 여기서 저장
            self.file_manager.save_results("navigation_anla")
            end_time = datetime.now()
            total_s = end_time - start_time
            
            self.logger.log("최종 저장 완료", "success")
            self.logger.log(f"소요시간: {str(total_s).split('.')[0]} (시:분:초)", "success")

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

    def _setup_wmc_for_domain(self, driver, cookie):
        #실제 도메인에서 WMC 모드 쿠키 설정#
        self.logger.log("WMC 모드 초기화 중...", "info")
        try:
            # 먼저 타겟 도메인으로 이동 (예: samsung.com)
            initial_url = "https://www.samsung.com/sg/"
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
        
    def _clear_memory(self, driver):
        #메모리 관리
        try:
            driver.execute_script("window.sessionStorage.clear();")
            driver.execute_script("window.localStorage.clear();")
            driver.delete_all_cookies()
            self.logger.log("메모리 정리 완료", "info")
        except:
            pass

    def _extract_navigation_menu(self, row, class_name):
        # 네비게이션 메뉴 추출#
        driver = self.selenium_manager.driver
        self.logger.log(f"클래스명 검색 시도: {class_name}", "info")
        # visible 메뉴 항목만 추출
        menu_items = driver.execute_script(
            f"""
            const items = [];
            
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
                f"  클래스명 '{class_name}'에 해당하는 요소를 찾을 수 없습니다."
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
                self.logger.log("login-token 쿠키가 없습니다.", "warning")
                
        except Exception as e:
            self.logger.log(f"쿠키 정리 중 오류: {str(e)}", "warning")

        # 추가: WMC 쿠키 재설정
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
                    # JS 파일은 차단하지 않음s
                ]
            })
            driver.execute_cdp_cmd('Network.enable', {})
            self.logger.log("리소스 차단 설정 완료 (JS 파일 유지)", "info")
        except Exception as e:
            self.logger.log(f"리소스 차단 설정 오류: {str(e)}", "warning")

        # 추가: 에러 페이지 확인
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
