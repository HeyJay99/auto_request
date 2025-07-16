import pyautogui
import pandas as pd
import time
import keyboard

stop = False

def stop_macro():
    global stop
    stop = True
    print("Q 키 감지됨: 자동화 중단 요청됨")

keyboard.add_hotkey('q', stop_macro)

df = pd.read_excel("C:/Users/MyCom/Desktop/auto/자재업로드.xlsx", sheet_name="3층", skiprows=2)

for index, row in df.iterrows():
    if stop:
        print("자동화 수동 중단됨.")
        break

    item_code = str(row['일흥품번']).strip()
    
    if pd.isna(row['일흥품번']):
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
