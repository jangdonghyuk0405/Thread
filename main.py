import requests
import importlib
import os
import sys

# PyInstaller가 의존성을 인식하도록 명시적 임포트
import tkinter
import tkinter.ttk
import tkinter.messagebox
import pyperclip

# 1. GitHub Raw 코드 주소 (본인의 주소로 교체)
URL = "https://raw.githubusercontent.com/jangdonghyuk0405/Thread/refs/heads/main/screw_input_app1-6.py"
FILE_NAME = "current_logic.py"

def update_and_run():
    try:
        # 서버에서 최신 코드 다운로드
        response = requests.get(URL)
        if response.status_code == 200:
            with open(FILE_NAME, "w", encoding="utf-8") as f:
                f.write(response.text)
            
            # 다운로드된 파일 실행
            if os.getcwd() not in sys.path:
                sys.path.append(os.getcwd())

            import current_logic
            importlib.reload(current_logic) # 이미 실행 중이면 새로고침
            current_logic.run_program()
        else:
            print("업데이트 서버에 연결할 수 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    update_and_run()