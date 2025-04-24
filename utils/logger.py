from PyQt5.QtWidgets import QApplication  # GUI 이벤트 처리를 위한 Qt Application 클래스 import

class Logger:
   #로그 메시지를 GUI에 출력
   
   def __init__(self, text_widget):
       self.text_widget = text_widget  # 로그 출력에 사용할 텍스트 위젯 저장

   def log(self, message, type='normal'):
       # type (str): 메시지 타입. 기본값은 'normal' 
       # 가능한 값: 'normal', 'success', 'error', 'warning', 'info', 'separator'
       # 메시지 타입별 텍스트 색상 정의
       styles = {
           'normal': 'color: #ffffff',    # 기본 텍스트 색상 
           'success': 'color: #2ecc71',   # 성공 메시지 (녹색)
           'error': 'color: #e74c3c',     # 에러 메시지 (빨간색)
           'warning': 'color: #f39c12',   # 경고 메시지 (주황색)
           'info': 'color: #3498db',      # 정보 메시지 (파란색)
           'separator': 'color: #ffffff'   # 구분선 (회색)
       }

       # 지정된 타입의 스타일 가져오기 (없으면 normal 스타일 사용)
       style = styles.get(type, styles['normal'])
       # 기본 폰트 스타일
       base_style = f'font-family: Consolas; font-size: 9pt; {style}'

       if type == 'separator':
           # 구분선 
           message = "─" * 100
       else:
           emojis = {
               'success': '✅ ',   # 성공 표시
               'error': '❌ ',     # 에러 표시
               'warning': '⚠️ ',   # 경고 표시
               'info': 'ℹ️ ',     # 정보 표시
               'normal': ''       # 기본 메시지는 이모지 없음
           }
           # 메시지 앞에 해당 타입의 이모지 추가
           message = f"{emojis.get(type, '')}{message}"

       # HTML 스타일이 적용된 메시지를 텍스트 위젯에 추가
       self.text_widget.append(f'<span style="{base_style}">{message}</span>')
       
       # GUI 이벤트 즉시 처리 (실시간 업데이트를 위해)
       QApplication.processEvents()