from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from PyQt5.QtWidgets import QApplication
from utils.logger import Logger
from datetime import datetime
import time


class KVFunction:
    def __init__(self, selenium_manager, file_manager, text_widget):
        self.selenium_manager = selenium_manager
        self.file_manager = file_manager
        self.logger = Logger(text_widget)

    def execute(self, filename, wait_time, server, username=None, password=None,cookie=None):
        start_time = datetime.now()
        self.logger.log(f"🚀 KV 시작: {start_time.strftime('%Y년 %m월 %d일 %H시 %M분 %S초')}", "info")
        self.logger.log(f"Time Sleep : {wait_time}초")
        QApplication.processEvents()
        
        try:
            # 파일 로드
            self.file_manager.load_excel(filename)
            self.logger.log(f" Active Sheet Name: {self.file_manager.worksheet.title}","success")
            QApplication.processEvents()
            
            # 드라이버 초기화 상태 확인
            self.selenium_manager.check_driver_initialized(self.logger)
            
            # 드라이버 설정
            driver = self.selenium_manager.driver
            driver.set_page_load_timeout(30)  # 타임아웃 증가
            self._block_resources_except_js(driver)

            # WMC 모드에서 초기 쿠키 설정
            if server == "WMC" and cookie:
                self._setup_wmc_for_domain(driver, cookie)

            # URL 처리
            for i in range(10, self.file_manager.worksheet.max_row + 1):
                self.logger.log("─" * 50)  # 구분선
                QApplication.processEvents()
                
                link = self.file_manager.worksheet["C" + str(i)].value
                if not link:
                    continue
                
                self.logger.log(f"{str(i-9)}번째 Link ->{str(link)}")
                QApplication.processEvents()
                
                try:
                     # 메모리 관리 - 30개마다 처리
                    if i % 30 == 0:
                        if server == "WMC" and cookie:
                            # WMC 모드에서는 login-token 쿠키 보존
                            self._clear_cookies_except_login(driver)
                            self.logger.log("WMC 모드에서 쿠키 정리 완료 (login-token 유지)", "info")
                        else:
                            # 일반 모드에서는 모든 메모리 정리
                            self._clear_memory(driver)
                    # 페이지 로드
                    driver.get(link)
                    time.sleep(wait_time)  # 페이지 로드 대기
                    #  QA 서버인 경우 로그인 먼저 처리
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
                            time.sleep(wait_time)  # 페이지 로드 대기
                    
                    # KV 데이터 추출 - 여기에 KV 관련 기능 구현
                    self._extract_kv_data(i)
                    self.logger.log("", "separator")  # 구분선 추가
                    QApplication.processEvents()
                except Exception as e:
                    self.logger.log(f"페이지 처리 중 오류: {str(e)}")
                    continue
            
            # 결과 저장
            self.file_manager.save_results("KV_cta")
            end_time = datetime.now()
            total_s = end_time - start_time
            
            self.logger.log("최종 저장 완료", "success")
            self.logger.log(f"소요시간: {str(total_s).split('.')[0]} (시:분:초)", "success")

            QApplication.processEvents()
            
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
            self.logger.log("메모리 정리 완료 ", "info")
        except:
            pass
    
    def _extract_kv_data(self, row):
        #데이터 추출
        driver = self.selenium_manager.driver

        try:
            # JS를 사용한 데이터 추출 - 수정된 버전
            kv_data = driver.execute_script("""
                const results = [];
                
                // 첫 번째 컴포: hd08-hero-kv-home__cta-wrap a (Button 1, Button 2)
                const ctaLinks1 = document.querySelectorAll('.hd08-hero-kv-home__cta-wrap a');
                
                // 두 번째 컴포 : 첫 번째 co76-feature-kv__container의 cta-wrap a만 선택
                const firstCo76Container = document.querySelector('.co76-feature-kv__container');
                const ctaLinks2 = firstCo76Container ? firstCo76Container.querySelectorAll('.co76-feature-kv__cta-wrap a') : [];
                
                // Banner 클릭 - url-to-kv 클래스가 있는 것만
                // 1. hd08-hero-kv-home__cta-wrap-clone a 중에서 url-to-kv 클래스가 있는 것만
                const hd08CloneLinks = document.querySelectorAll('.hd08-hero-kv-home__cta-wrap-clone a');
                const bannerLinks1 = Array.from(hd08CloneLinks).filter(link => link.classList.contains('url-to-kv'));
                
                // 2. 첫 번째 co76-feature-kv__container의 cta-wrap-clone a 중에서 url-to-kv 클래스가 있는 것만
                const co76CloneLinks = firstCo76Container ? firstCo76Container.querySelectorAll('.co76-feature-kv__cta-wrap-clone a') : [];
                const bannerLinks2 = Array.from(co76CloneLinks).filter(link => link.classList.contains('url-to-kv'));
                
                // hd08 데이터 처리 (Button 1, Button 2)
                if (ctaLinks1.length > 0) {
                    ctaLinks1.forEach((link, index) => {
                        const hasDisplayOnClass = link.classList.contains('cta-kv-display-on');
                        const hasText = link.textContent.trim().length > 0;
                        const hasValidHref = link.getAttribute("href") && 
                                            link.getAttribute("href") !== "" && 
                                            link.getAttribute("href") !== "null";
                        
                        if (hasDisplayOnClass || (hasText && hasValidHref)) {
                            if (hasText) {
                                results.push({
                                    index: index + 1,
                                    type: 'hd08_button',
                                    anCa: link.getAttribute("an-ca") || "",
                                    anAc: link.getAttribute("an-ac") || "",
                                    anLa: link.getAttribute("an-la") || "",
                                    text: link.textContent.trim(),
                                    source: 'hd08-hero-kv-home'
                                });
                            }
                        }
                    });
                }
                
                // co76 데이터 처리 (Button 1, Button 2)
                if (ctaLinks2.length > 0) {
                    ctaLinks2.forEach((link, index) => {
                        const hasDisplayOnClass = link.classList.contains('cta-kv-display-on');
                        const hasText = link.textContent.trim().length > 0;
                        const hasValidHref = link.getAttribute("href") && 
                                            link.getAttribute("href") !== "" && 
                                            link.getAttribute("href") !== "null";
                        
                        if (hasDisplayOnClass || (hasText && hasValidHref)) {
                            if (hasText) {
                                results.push({
                                    index: index + 1,
                                    type: 'co76_button',
                                    anCa: link.getAttribute("an-ca") || "",
                                    anAc: link.getAttribute("an-ac") || "",
                                    anLa: link.getAttribute("an-la") || "",
                                    text: link.textContent.trim(),
                                    source: 'co76-feature-kv'
                                });
                            }
                        }
                    });
                }
                
                // hd08 Banner 처리
                if (bannerLinks1.length > 0) {
                    bannerLinks1.forEach((link, index) => {
                        const hasText = link.textContent.trim().length > 0;
                        const hasValidHref = link.getAttribute("href") && 
                                            link.getAttribute("href") !== "" && 
                                            link.getAttribute("href") !== "null";
                        
                        if (hasText && hasValidHref) {
                            results.push({
                                index: index + 1,
                                type: 'hd08_banner',
                                anCa: link.getAttribute("an-ca") || "",
                                anAc: link.getAttribute("an-ac") || "",
                                anLa: link.getAttribute("an-la") || "",
                                text: link.textContent.trim(),
                                source: 'hd08-hero-kv-home'
                            });
                        }
                    });
                }
                
                // co76 Banner 처리
                if (bannerLinks2.length > 0) {
                    bannerLinks2.forEach((link, index) => {
                        const hasText = link.textContent.trim().length > 0;
                        const hasValidHref = link.getAttribute("href") && 
                                            link.getAttribute("href") !== "" && 
                                            link.getAttribute("href") !== "null";
                        
                        if (hasText && hasValidHref) {
                            results.push({
                                index: index + 1,
                                type: 'co76_banner',
                                anCa: link.getAttribute("an-ca") || "",
                                anAc: link.getAttribute("an-ac") || "",
                                anLa: link.getAttribute("an-la") || "",
                                text: link.textContent.trim(),
                                source: 'co76-feature-kv'
                            });
                        }
                    });
                }
                
                return results;
            """)
            
            # 추출한 데이터가 없을 경우
            if not kv_data or len(kv_data) == 0:
                self.logger.log("❌ KV 데이터를 찾을 수 없습니다.", "warning")
                return
                
            # 결과 로깅 및 엑셀에 저장
            self.logger.log(f"KV에서 {len(kv_data)}개의 CTA 요소 발견","info")

            # 각 타입별로 분류
            hd08_buttons = [item for item in kv_data if item['type'] == 'hd08_button']
            co76_buttons = [item for item in kv_data if item['type'] == 'co76_button']
            hd08_banners = [item for item in kv_data if item['type'] == 'hd08_banner']
            co76_banners = [item for item in kv_data if item['type'] == 'co76_banner']
            
            self.logger.log(f"hd08 버튼: {len(hd08_buttons)}개, co76 버튼: {len(co76_buttons)}개", "info")
            self.logger.log(f"hd08 배너: {len(hd08_banners)}개, co76 배너: {len(co76_banners)}개", "info")
            
            # hd08 Button 1, Button 2 저장 (E, F, G / H, I, J 열)
            for idx, button in enumerate(hd08_buttons[:2]):  # 최대 2개의 버튼만 처리
                anCa = button['anCa']
                anAc = button['anAc']
                anLa = button['anLa']
                text = button['text']
                
                self.logger.log(f"hd08 Button {idx+1}: {text}")
                self.logger.log(f"  an-ca: {anCa}")
                self.logger.log(f"  an-ac: {anAc}")
                self.logger.log(f"  an-la: {anLa}")
                
                try:
                    if idx == 0:  # hd08 Button 1 - E, F, G 열
                        self.file_manager.worksheet[f"E{row}"] = anCa
                        self.file_manager.worksheet[f"F{row}"] = anAc
                        self.file_manager.worksheet[f"G{row}"] = anLa
                    elif idx == 1:  # hd08 Button 2 - H, I, J 열
                        self.file_manager.worksheet[f"H{row}"] = anCa
                        self.file_manager.worksheet[f"I{row}"] = anAc
                        self.file_manager.worksheet[f"J{row}"] = anLa
                except Exception as e:
                    self.logger.log(f"hd08 버튼 엑셀 저장 중 오류: {str(e)}", "error")
            
            # hd08 Banner 저장 (K, L, M 열)
            if hd08_banners:
                banner = hd08_banners[0]  # 첫 번째 배너만 사용
                anCa = banner['anCa']
                anAc = banner['anAc']
                anLa = banner['anLa']
                text = banner['text']
                
                self.logger.log(f"hd08 Banner: {text}")
                self.logger.log(f"  an-ca: {anCa}")
                self.logger.log(f"  an-ac: {anAc}")
                self.logger.log(f"  an-la: {anLa}")
                
                try:
                    # hd08 Banner - K, L, M 열
                    self.file_manager.worksheet[f"K{row}"] = anCa
                    self.file_manager.worksheet[f"L{row}"] = anAc
                    self.file_manager.worksheet[f"M{row}"] = anLa
                except Exception as e:
                    self.logger.log(f"hd08 배너 엑셀 저장 중 오류: {str(e)}", "error")
            
            # co76 Button 1, Button 2 저장 (N, O, P / Q, R, S 열)
            for idx, button in enumerate(co76_buttons[:2]):  # 최대 2개의 버튼만 처리
                anCa = button['anCa']
                anAc = button['anAc']
                anLa = button['anLa']
                text = button['text']
                
                self.logger.log(f"co76 Button {idx+1}: {text}")
                self.logger.log(f"  an-ca: {anCa}")
                self.logger.log(f"  an-ac: {anAc}")
                self.logger.log(f"  an-la: {anLa}")
                
                try:
                    if idx == 0:  # co76 Button 1 - N, O, P 열
                        self.file_manager.worksheet[f"N{row}"] = anCa
                        self.file_manager.worksheet[f"O{row}"] = anAc
                        self.file_manager.worksheet[f"P{row}"] = anLa
                    elif idx == 1:  # co76 Button 2 - Q, R, S 열
                        self.file_manager.worksheet[f"Q{row}"] = anCa
                        self.file_manager.worksheet[f"R{row}"] = anAc
                        self.file_manager.worksheet[f"S{row}"] = anLa
                except Exception as e:
                    self.logger.log(f"co76 버튼 엑셀 저장 중 오류: {str(e)}", "error")
            
            # co76 Banner 저장 (T, U, V 열)
            if co76_banners:
                banner = co76_banners[0]  # 첫 번째 배너만 사용
                anCa = banner['anCa']
                anAc = banner['anAc']
                anLa = banner['anLa']
                text = banner['text']
                
                self.logger.log(f"co76 Banner: {text}")
                self.logger.log(f"  an-ca: {anCa}")
                self.logger.log(f"  an-ac: {anAc}")
                self.logger.log(f"  an-la: {anLa}")
                
                try:
                    # co76 Banner - T, U, V 열
                    self.file_manager.worksheet[f"T{row}"] = anCa
                    self.file_manager.worksheet[f"U{row}"] = anAc
                    self.file_manager.worksheet[f"V{row}"] = anLa
                except Exception as e:
                    self.logger.log(f"co76 배너 엑셀 저장 중 오류: {str(e)}", "error")
            
            self.logger.log("-" * 50)
            total_items = len(hd08_buttons) + len(co76_buttons) + len(hd08_banners) + len(co76_banners)
            self.logger.log(f"행 {row}에 KV 정보 저장 완료 (총 {total_items}개)", "success")

        except Exception as e:
            self.logger.log(f"KV 데이터 추출 중 오류: {str(e)}","error")
            self.file_manager.save_error_file()

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

    # 추가: 로그인 쿠키 보존 메모리 정리
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
                self.logger.log("login-token 쿠키가 없습니다. ", "warning")
                
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
                