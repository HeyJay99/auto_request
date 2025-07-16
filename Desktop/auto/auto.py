import pyautogui
import pandas as pd
import time
import keyboard

stop = False

#중단 함수 생성
def stop_macro():
    global stop
    stop = True
    print("Q 키 감지됨: 자동화 중단 요청됨")

#키보드 q를 눌렀을때 중단 함수가 호출 될수 있도록 설정
keyboard.add_hotkey('q', stop_macro)

#데이터 프레임 불러오기, 시트와 읽어올 행 설정
df = pd.read_excel("C:/Users/MyCom/Desktop/auto/자재업로드.xlsx", sheet_name="2층", skiprows=2)

#매크로 반복문 설정
for index, row in df.iterrows():

    #매크로로 인해 중단될 경우 메세지 출력 후 중단
    if stop:
        print("자동화 수동 중단됨.")
        break

    #아이템 코드 읽어오기 + 공백제거
    item_code = str(row['일흥품번']).strip()
    
    #아이템 코드를 다읽고 빈칸이 나오면 프로그램 종료
    if item_code == None:
        print("작업 완료")
        break

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
