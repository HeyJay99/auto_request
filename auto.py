import pyautogui
import pandas as pd
import time
import keyboard
from tkinter import *
from tkinter import filedialog, messagebox, ttk

window = Tk()

#윈도우 기본 설정
window.title("ERP auto_request")
window.geometry("400x200")
window.resizable(True,True)

file_label = Label(window, text="불러온 파일 없음")
file_label.pack(pady=10)

selected_file_path = ""
options = ["2층", "3층"]

# 변수 지정
selected_option = StringVar()
selected_option.set(options[0])  # 기본 선택값

# 드롭다운 만들기
dropdown = ttk.Combobox(window, textvariable=selected_option, values=options, state="readonly")
dropdown.pack(pady=20)

#파일 불러오기 함수
def open_file():
    global selected_file_path
    path = filedialog.askopenfilename()
    filetypes=[("자재 파일", "*.xlsx")]
    print("선택한 파일", path)
    if path:
        selected_file_path = path
        file_label.config(text=f"불러온 파일: {path}")

def print_selection():
    print("선택된 옵션:", selected_option.get())
   


#파일선택 버튼 생성
btn = Button(window) 
btn.config(text="파일 선택")
btn.config(command = open_file)
btn.pack()


#매크로 실행
def run_macro():
    
    if not selected_file_path:
        messagebox.showwarning("실행 오류", "먼저 파일을 선택해주세요.")
        return
    
    global stop
    stop = False
    global completed
    completed = True


    #중단 함수 생성
    def stop_macro():
        global stop
        stop = True
        completed = False
        temp = Toplevel()
        temp.withdraw()               # 새 창은 숨기고
        temp.attributes("-topmost", True)
        messagebox.showwarning("자동화 중단 요청됨", "매크로가 중단 되었습니다.")
        temp.destroy()  

    #키보드 q를 눌렀을때 중단 함수가 호출 될수 있도록 설정
    keyboard.add_hotkey('q', stop_macro)

    #데이터 프레임 불러오기, 시트와 읽어올 행 설정
    df = pd.read_excel(selected_file_path, sheet_name = "Sheet2", skiprows=2)

    #매크로 반복문 설정
    for index, row in df.iterrows():

        #매크로로 인해 중단될 경우 메세지 출력 후 중단
        if stop:
            print("자동화 수동 중단됨.")
            break

        #아이템 코드 읽어오기 + 공백제거
        item_code = str(row['일흥품번']).strip()

        # 비고란에서 키워드 포함 여부 확인
        remark = str(row.get('비고', '')).strip()  # '비고' 열이 없으면 기본값은 ''
        keywords = [selected_option.get()]

        if not any(k in remark for k in keywords):
            continue

        pyautogui.moveTo(x=368, y=619)
        pyautogui.click()
        time.sleep(0.5)

        pyautogui.moveTo(x=656, y=323)
        pyautogui.doubleClick()
        time.sleep(0.5)

        pyautogui.write(item_code)
        time.sleep(0.2)

        pyautogui.moveTo(x=1361, y=407)
        pyautogui.click()
        time.sleep(1)

        pyautogui.moveTo(x=574, y=495)
        pyautogui.click()
        time.sleep(1)

        pyautogui.moveTo(x=1272, y=788)
        pyautogui.click()
        time.sleep(0.5)

        # if completed:
        #     temp = Toplevel()
        #     temp.withdraw()
        #     temp.attributes("-topmost", True)
        #     messagebox.showinfo("작업 완료", "자재 등록이 정상적으로 완료되었습니다!", parent=temp)
        #     temp.destroy()



#매크로 실행 버튼 생성
run_btn = Button(window, text="매크로 실행", command=run_macro)
run_btn.pack(pady=10)

window.mainloop()

