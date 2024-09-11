import pyautogui

# 특정 윈도우 내에서 이미지 찾기
notepad = pyautogui.getWindowsWithTitle('메모장')[0]
button_location = pyautogui.locateOnScreen('button.png', region=(notepad.left, notepad.top, notepad.width, notepad.height))

# 찾은 위치 클릭
if button_location:
    pyautogui.click(button_location)
    
    
    
# 활성 윈도우 제목 얻기
active_window = pyautogui.getActiveWindow()
print(active_window.title)

# 특정 제목의 윈도우 찾기
notepad = pyautogui.getWindowsWithTitle('메모장')[0]

# 윈도우 활성화
notepad.activate()

# 윈도우 크기 조절
notepad.resizeTo(500, 500)

# 윈도우 이동
notepad.moveTo(100, 100)

# 윈도우 최대화, 최소화, 복원
notepad.maximize()
notepad.minimize()
notepad.restore()

# 윈도우 닫기
notepad.close()


# 안전장치 설정 (마우스를 왼쪽 상단 구석으로 이동하면 동작 중지)
pyautogui.FAILSAFE = True

# 일정 시간 동안 일시 정지
pyautogui.sleep(2)

import pyautogui

# 텍스트 입력
pyautogui.write('Hello, World!')

# 특정 키 누르기
pyautogui.press('enter')

# 여러 키 동시에 누르기
pyautogui.hotkey('ctrl', 'c')



# 마우스를 특정 위치로 이동
pyautogui.moveTo(100, 100, duration=1)

# 현재 위치에서 클릭
pyautogui.click()

# 특정 위치에서 더블 클릭
pyautogui.doubleClick(x=200, y=200)



# 전체 화면 캡처
screenshot = pyautogui.screenshot()
screenshot.save('screen.png')

# 특정 영역 캡처
region_screenshot = pyautogui.screenshot(region=(0, 0, 300, 400))
region_screenshot.save('region.png')


# 화면에서 이미지 찾기
location = pyautogui.locateOnScreen('button.png')

# 찾은 이미지 클릭
if location:
    pyautogui.click(location)
    
    
import pyautogui
import subprocess

# 메모장 실행
subprocess.Popen('notepad.exe')

# 실행된 메모장 찾기
notepad = pyautogui.getWindowsWithTitle('제목 없음')[0]

# 메모장에 텍스트 입력
notepad.activate()
pyautogui.write('Hello, PyAutoGUI!')


import pyautogui

# 특정 윈도우의 스크린샷 찍기
notepad = pyautogui.getWindowsWithTitle('메모장')[0]
screenshot = pyautogui.screenshot(region=(notepad.left, notepad.top, notepad.width, notepad.height))
screenshot.save('notepad_screenshot.png')