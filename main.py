import sys
import os
import subprocess

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QButtonGroup
from PyQt5.uic import loadUi
from utils.logger import Logger
from core.selenium_manager import SeleniumManager
from core.file_manager import FileManager
from functions.function1 import HomeRedirectFunction
from functions.function2 import DigitalDataFunction
from functions.function3 import PageTrackFunction
from functions.function4 import NavigationExtractFunction
from functions.function5 import KVFunction
from core.config import UI_FILE_PATH, WINDOW_SIZE


class MainWindow(QMainWindow):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        try:
            if not os.path.exists(UI_FILE_PATH):
                raise FileNotFoundError(f"UI file not found at {UI_FILE_PATH}")

            loadUi(UI_FILE_PATH, self)
            self.setFixedSize(*WINDOW_SIZE)
        except Exception as e:
            print(f"UI 파일 로드 오류: {str(e)}")
            sys.exit(1)
        # Logger 초기화
        self.logger = Logger(self.textEdit)
        
        # 초기 설정
        self.radioButton_2.setChecked(True)
        self.uploaded_filename = None
        self.selenium_manager = None
        self.file_manager = FileManager()
        self.className_label.hide()
        self.lineEdit_class.hide()

        # 초기값 설정
        self.comboBox_Select()
        username = self.lineEdit_id.text()
        password = self.lineEdit_password.text()
        self.logger.log("ID: " + username)
        self.logger.log("Password: " + password)

        # 서버 선택 버튼 그룹
        self.server_group = QButtonGroup()
        self.server_group.addButton(self.radioButton_1)
        self.server_group.addButton(self.radioButton_2)

        # 브라우저 선택 버튼 그룹
        self.browser_group = QButtonGroup()
        self.browser_group.addButton(self.edge_radioButton)  # Edge
        self.browser_group.addButton(self.chrome_radioButton)  # Chrome
        self.browser_group.addButton(self.chromelatest_radioButton)  # Latest Chrome

        # 기능 선택 버튼 그룹
        self.function_group = QButtonGroup()
        self.function_group.addButton(self.functionBtn_1)
        self.function_group.addButton(self.functionBtn_2)
        self.function_group.addButton(self.functionBtn_3)
        self.function_group.addButton(self.functionBtn_4)  # 추가
        self.function_group.addButton(self.functionBtn_5)  # 추가


        # 버튼 연결
        self.runButton.clicked.connect(self.run_function)
        self.uploadButton.clicked.connect(self.upload_file)
        self.radioButton_1.toggled.connect(self.toggle_input_fields)
        self.radioButton_2.toggled.connect(self.toggle_input_fields)
        self.removeBtn.clicked.connect(self.clear_text)
        self.exitBtn.clicked.connect(self.exitButton)
        self.functionBtn_1.toggled.connect(self.radio_function_button)
        self.functionBtn_2.toggled.connect(self.radio_function_button)
        self.functionBtn_3.toggled.connect(self.radio_function_button)
        self.functionBtn_4.toggled.connect(self.radio_function_button)  # 추가
        self.functionBtn_5.toggled.connect(self.radio_function_button)  # 추가
        self.timeBox.currentIndexChanged.connect(self.comboBox_Select)

        # 브라우저 선택 라디오 버튼 연결
        self.edge_radioButton.toggled.connect(self.browser_selection_changed)
        self.chrome_radioButton.toggled.connect(self.browser_selection_changed)
        self.chromelatest_radioButton.toggled.connect(self.browser_selection_changed)

    # 브라우저 선택 변경 이벤트 처리 메서드 추가
    def browser_selection_changed(self):
        sender = self.sender()
        if not sender.isChecked():
            return
        if self.edge_radioButton.isChecked():
            self.logger.log("Selected Browser: Edge (Local Driver)")
        elif self.chrome_radioButton.isChecked():
            self.logger.log("Selected Browser: Chrome (Local Driver)")
        elif self.chromelatest_radioButton.isChecked():
            self.logger.log("Selected Browser: Chrome (Latest Version)")

    def comboBox_Select(self):
        # 대기 시간 선택#
        selectsec = int(self.timeBox.currentText())
        self.logger.log(f"Selected Sec: {selectsec}")

    def radio_function_button(self):
        # 기능 버튼 상태 변경#
        if self.functionBtn_1.isChecked():
            self.captureCheck.setEnabled(True)
        else:
            self.captureCheck.setEnabled(False)
        # 클래스명 입력 필드 표시/숨김
        if self.functionBtn_4.isChecked():
            self.className_label.show()
            self.lineEdit_class.show()
        else:
            self.className_label.hide()
            self.lineEdit_class.hide()

    def toggle_input_fields(self):
        # ID/PW 입력 필드 토글#
        if self.radioButton_2.isChecked():
            self.lineEdit_id.show()
            self.lineEdit_password.show()
        else:
            self.lineEdit_id.hide()
            self.lineEdit_password.hide()

    def run_function(self):
        # 기능 실행#
        selectsec = int(self.timeBox.currentText())
        self.comboBox_Select()

        if not self.uploaded_filename:
            self.logger.log("❌ Please upload a file ❌")
            return
        # 브라우저 타입 확인
        if self.edge_radioButton.isChecked():  # Edge
            browser_type = "edge"
            use_local_driver = True
        elif self.chrome_radioButton.isChecked():  # Chrome (로컬)
            browser_type = "chrome"
            use_local_driver = True
        else:  # Chrome (최신)
            browser_type = "chrome"
            use_local_driver = False
        
        self.logger.log(f"브라우저 설정: {browser_type} {'(로컬 드라이버)' if use_local_driver else '(최신 버전)'}")
        QApplication.processEvents()
        # Selenium 매니저 초기화 - 
        try:
            # 기존 매니저가 있으면 종료
            if hasattr(self, 'selenium_manager') and self.selenium_manager:
                self.logger.log("기존 드라이버 종료 중...")
                self.selenium_manager.quit_driver()
                
            # 새 매니저 생성
            self.logger.log("새 드라이버 매니저 생성 중...")
            self.selenium_manager = SeleniumManager()
            
            # 드라이버 초기화 - 명시적으로 성공 여부 확인
            driver = self.selenium_manager.initialize_driver(
                browser_type=browser_type,
                headless=self.webDrivercheckBox.isChecked(),
                use_local_driver=use_local_driver,
            )
            
            if driver is None:
                self.logger.log("❌ 드라이버 초기화 실패! 브라우저를 확인해주세요.")
                self.enable_buttons()
                return
                
            self.logger.log("✅ 드라이버 초기화 성공!")
            
        except Exception as e:
            self.logger.log(f"❌ 드라이버 초기화 중 오류: {str(e)}")
            self.enable_buttons()
            return
        

        # 서버 타입 확인
        if self.radioButton_1.isChecked():
            server_type = "WWW"
            username = None
            password = None
        else:
            server_type = "QA"
            username = self.lineEdit_id.text()
            password = self.lineEdit_password.text()
            if not username or not password:
                self.logger.log("❌ 아이디, 패스워드 입력해주세요 ❌")
                return

        self.logger.log(f"Selected option: {server_type}")

        # 기능 선택 및 실행
        try:
            self.disable_buttons()

            if self.functionBtn_1.isChecked():
                function = HomeRedirectFunction(
                    self.selenium_manager, self.file_manager, self.textEdit,self.captureCheck
                )
                self.logger.log("Selected Function: Home Redirection")
            elif self.functionBtn_2.isChecked():
                function = DigitalDataFunction(
                    self.selenium_manager, self.file_manager, self.textEdit
                )
                self.logger.log("Selected Function: Digital Data")
            elif self.functionBtn_3.isChecked():
                function = PageTrackFunction(
                    self.selenium_manager, self.file_manager, self.textEdit
                )
                self.logger.log("Selected Function: Page Track")
            elif self.functionBtn_4.isChecked():
                # 클래스명 입력 확인
                class_name = self.lineEdit_class.text()
                if not class_name.strip():
                    self.logger.log("❌ 추출할 클래스명을 입력해주세요 ❌")
                    return
                function = NavigationExtractFunction(
                    self.selenium_manager, self.file_manager, self.textEdit
                )
                self.logger.log("Selected Function: Navigation an-la")
                self.logger.log(f"Target Class: {class_name}")

                function.execute(
                    self.uploaded_filename,
                    selectsec,
                    server_type,
                    username,
                    password,
                    class_name,
                )
                return
            elif self.functionBtn_5.isChecked():
                function = KVFunction(
                    self.selenium_manager, self.file_manager, self.textEdit
                )
                self.logger.log("Selected Function: KV CTA")
            else:
                self.logger.log("❌ Please select a function ❌")
                return

            # 기능 실행
            function.execute(
                self.uploaded_filename, selectsec, server_type, username, password
            )

        except Exception as e:
            self.logger.log(f"실행 중 오류 발생: {str(e)}")
            # 강제 종료 로직 추가
            if self.selenium_manager and self.selenium_manager.driver:
                try:
                    self.selenium_manager.quit_driver()
                except:
                    pass
            self.selenium_manager = None
        finally:
            self.enable_buttons()

    def upload_file(self):
        # 파일 업로드#
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "Excel Files (*.xlsx)"
        )

        if filename:
            self.uploaded_filename = filename
            self.logger.log(f"Uploaded file: {filename}")
        else:
            self.logger.log(".xlsx 파일형식만 업로드해주세요.")

    def clear_text(self):
        # 텍스트 창 초기화#
        self.textEdit.clear()

    def closeEvent(self, event):
        # 프로그램 종료 시 실행#
        try:
            if self.selenium_manager and self.selenium_manager.driver:
                self.selenium_manager.quit_driver()

            if hasattr(self, "file_manager"):
                self.file_manager.close()

            event.accept()
        except Exception as e:
            print(f"종료 중 오류 발생: {e}")
            event.accept()

    def terminate_if_running(self, process_name):
        # 프로세스가 실행 중인 경우에만 종료#
        try:
            # 프로세스 존재 여부 확인
            check_cmd = f'tasklist /FI "IMAGENAME eq {process_name}" /NH'
            output = subprocess.check_output(check_cmd, shell=True).decode("utf-8")

            # 프로세스가 존재하는 경우에
            if process_name in output:
                subprocess.call(f"taskkill /F /IM {process_name}", shell=True)
                return True
            return False
        except Exception:
            return False

    def exitButton(self):
        # 종료 버튼 클릭 시 실행#
        try:
            if self.selenium_manager and self.selenium_manager.driver:
                self.selenium_manager.quit_driver()
            # 백그라운드 프로세스 강제 종료
            self.terminate_if_running("msedgedriver.exe")
            self.terminate_if_running("chromedriver.exe")
            if hasattr(self, "file_manager"):
                self.file_manager.close()

            self.close()
            QApplication.quit()  # 이 부분 추가

        except Exception as e:
            print(f"종료 버튼 처리 중 오류 발생: {e}")
            self.close()
            QApplication.quit()

    def disable_buttons(self):
        # 버튼 비활성화#
        buttons = [
            self.runButton,
            self.uploadButton,
            self.radioButton_1,
            self.radioButton_2,
            self.edge_radioButton,
            self.chrome_radioButton,
            self.chromelatest_radioButton,
            self.functionBtn_1,
            self.functionBtn_2,
            self.functionBtn_3,
            self.functionBtn_4,
            self.functionBtn_5,
            self.captureCheck,
            self.removeBtn,
            self.webDrivercheckBox,
            self.timeBox,
        ]
        for button in buttons:
            button.setEnabled(False)

    def enable_buttons(self):
        # 버튼 활성화#
        buttons = [
            self.runButton,
            self.uploadButton,
            self.radioButton_1,
            self.radioButton_2,
            self.edge_radioButton,
            self.chrome_radioButton,
            self.chromelatest_radioButton,
            self.functionBtn_1,
            self.functionBtn_2,
            self.functionBtn_3,
            self.functionBtn_4,
            self.functionBtn_5,
            self.captureCheck,
            self.removeBtn,
            self.webDrivercheckBox,
            self.timeBox,
        ]
        for button in buttons:
            button.setEnabled(True)

def main():
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.logger.log("Application starting...")
        window.show()
        sys.exit(app.exec_())

    except Exception as e:
        print(f"Error in main: {str(e)}")
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Unhandled exception: {str(e)}")
