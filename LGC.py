import pyautogui
import pygetwindow
import pywinauto
import cv2
import pytesseract
import pandas as pd
import time
from PIL import ImageGrab

# https://github.com/PowercoderJr/lords-gifts-counter/blob/master/analyze.py
# LDPlayer 960*540 160dpi english

# 변수 선언
tesseract_config = '-c tessedit_char_whitelist=" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" --oem 3 --psm 7'
rarities = ['Common', 'Uncommon', 'Rare', 'Epic', 'Legendary']
nickname_list = []
df = pd.DataFrame(columns=rarities)
error_count = 0
filename = time.strftime('%y%m%d-%H%M') # 시작 시간으로 파일 이름 저장
i = 0
tmpname = 1

# LDPlayer 창 활성화하기
win = pygetwindow.getWindowsWithTitle('LDPlayer')[0]
if not win.isActive:
    pywinauto.application.Application().connect(handle=win._hWnd).top_window().set_focus()
win.activate()
fore = pyautogui.getActiveWindow()
pyautogui.moveTo(fore.left, fore.top)
pyautogui.move(830, 260)

# 전체 상자 수 확인
pyautogui.hotkey('alt', 'printscreen')
ImageGrab.grabclipboard().crop((700, 200, 787, 222)).save('gift.png', 'png')
gift_gray = cv2.imread('gift.png', cv2.IMREAD_GRAYSCALE)
gift_gray = cv2.threshold(gift_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
gift_text = pytesseract.image_to_string(gift_gray, lang='eng', config=tesseract_config)
gift_text = int(gift_text[6:len(gift_text)])
print('전체 상자 수:', gift_text)

while gift_text + error_count != 0:
    i += 1
    # 활성화 창 클립보드로 스샷 찍고 가공 후 저장
    pyautogui.hotkey('alt', 'printscreen')
    im1 = ImageGrab.grabclipboard()
    im1.crop((480, 300, 1130, 400)).save(('./data/' + str(tmpname) + '.png'), 'png')
    im1.crop((492, 300, 900, 330)).save('rarity.png', 'png')
    im1.crop((640, 330, 900, 360)).save('nickname.png', 'png')

    # 전처리 & OCR
    rarity_gray = cv2.imread('rarity.png', cv2.IMREAD_GRAYSCALE)
    rarity_gray = cv2.threshold(rarity_gray, 120, 255, cv2.THRESH_BINARY_INV)[1]
    rarity_text = pytesseract.image_to_string(rarity_gray, lang='eng', config=tesseract_config)
    rarity_text = rarity_text.split()[0]
    nickname_gray = cv2.imread('nickname.png', cv2.IMREAD_GRAYSCALE)
    nickname_gray = cv2.threshold(nickname_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
    nickname_text = pytesseract.image_to_string(nickname_gray, lang='eng', config=tesseract_config).strip('\n')

    print(str(i) + '번 째:', nickname_text, '/', rarity_text)

    # 데이터프레임에 넣기
    if nickname_text in nickname_list:
        pass
    else:
        nickname_list.append(nickname_text)
        df2 = pd.DataFrame(index=[nickname_text], columns=rarities, data=[[0, 0, 0, 0, 0]])
        df = pd.concat([df, df2])

    if rarity_text in rarities:
        if rarity_text == "Common":
            df.loc[nickname_text, 'Common'] += 1
        elif rarity_text == "Uncommon":
            df.loc[nickname_text, 'Uncommon'] += 1
        elif rarity_text == "Rare":
            df.loc[nickname_text, 'Rare'] += 1
        elif rarity_text == "Epic":
            df.loc[nickname_text, 'Epic'] += 1
        else:
            df.loc[nickname_text, 'Legendary'] += 1
        pyautogui.click()
        time.sleep(1)
        pyautogui.click()
        time.sleep(1)
        tmpname += 1
        if error_count > 0:
            error_count -= 1
        else:
            gift_text -= 1
    else:
        error_count += 1
        time.sleep(2)

print('전체 실행 횟수:', i )

# 데이터프레임 가공하기 (총 상자 수, 총 포인트)
Total_gifts = (df['Common'].sum(), df['Uncommon'].sum(), df['Rare'].sum(), df['Epic'].sum(), df['Legendary'].sum())
df3 = pd.DataFrame(index=['Total'], columns=rarities, data=[Total_gifts])
df = pd.concat([df, df3])
df['Points'] = df['Common']*1 + df['Uncommon']*5 + df['Rare']*15 + df['Epic']*30 + df['Legendary']*50

# 엑셀로 저장하기
df.to_excel(filename+'.xlsx')
