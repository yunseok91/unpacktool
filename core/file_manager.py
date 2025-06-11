import openpyxl
import os
from datetime import datetime
from .config import RESULTS_FOLDERS

class FileManager:
    def __init__(self):
        self.workbook = None
        self.worksheet = None
        self.current_file = None

    def load_excel(self, filepath):
        #Excel 파일 로드#
        try:
            self.workbook = openpyxl.load_workbook(filepath)
            self.worksheet = self.workbook.active
            self.current_file = filepath
            return True
        except Exception as e:
            raise Exception(f"Excel 파일 로드 실패: {str(e)}")

    def save_results(self, folder_type, start_time=None):
        #결과 저장#
        if not self.workbook:
            return False

        try:
            # 저장 폴더 생성
            folder_name = RESULTS_FOLDERS.get(folder_type, 'default')
            save_folder = os.path.join(os.getcwd(), folder_name)
            os.makedirs(save_folder, exist_ok=True)

            # 파일명 생성
            if start_time:
                timestamp = start_time.strftime('%Y-%m-%d-%H-%M-%S')
            else:
                timestamp = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            
            filename = f"{timestamp}_{folder_type}.xlsx"
            save_path = os.path.join(save_folder, filename)

            # 파일 저장
            self.workbook.save(save_path)
            return True
        except Exception as e:
            raise Exception(f"결과 저장 실패: {str(e)}")

    def save_error_file(self):
        #에러 발생 시 파일 저장#
        if not self.workbook or not self.current_file:
            return False
        try:
            # 에러 폴더 생성 확인
            error_folder = os.path.join(os.getcwd(), 'error')
            if not os.path.exists(error_folder):
                os.makedirs(error_folder)
            
            error_filename = os.path.join(error_folder, '비정상_' + os.path.basename(self.current_file))
            self.workbook.save(error_filename)
            return True
        except Exception as e:
            print(f"에러 파일 저장 실패: {str(e)}")
            return False

    def close(self):
        #파일 정리#
        if self.workbook:
            try:
                self.workbook.close()
            except Exception:
                pass
            finally:
                self.workbook = None
                self.worksheet = None