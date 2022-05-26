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
gift = ImageGrab.grabclipboard().crop((700, 200, 787, 222)).save('gift.png', 'png')
gift_gray = cv2.imread('gift.png', cv2.IMREAD_GRAYSCALE)
gift_gray = cv2.threshold(gift_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
gift_text = pytesseract.image_to_string(gift_gray, lang='eng', config=tesseract_config)
gift_text = int(gift_text[6:len(gift_text)])
print('Gifts:', gift_text)

for i in range(gift_text):
    print(i, '번 째')
    # 활성화 창 클립보드로 스샷 찍고 가공 후 저장
    pyautogui.hotkey('alt', 'printscreen')
    im1 = ImageGrab.grabclipboard()
    rarity = im1.crop((492, 300, 900, 330)).save('rarity.png', 'png')
    nickname = im1.crop((640, 330, 900, 360)).save('nickname.png', 'png')

    # 전처리 & OCR
    rarity_gray = cv2.imread('rarity.png', cv2.IMREAD_GRAYSCALE)
    rarity_gray = cv2.threshold(rarity_gray, 120, 255, cv2.THRESH_BINARY_INV)[1]
    rarity_text = pytesseract.image_to_string(rarity_gray, lang='eng', config=tesseract_config)
    rarity_text = rarity_text.split()[0]
    nickname_gray = cv2.imread('nickname.png', cv2.IMREAD_GRAYSCALE)
    nickname_gray = cv2.threshold(nickname_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
    nickname_text = pytesseract.image_to_string(nickname_gray, lang='eng', config=tesseract_config).strip('\n')

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
        time.sleep(0.7)
        pyautogui.click()
        time.sleep(1)

    else:
        # df.to_excel('test(error).xlsx')
        # print('ERROR', rarity_text, nickname_text, end='')
        pass
filename = time.strftime(%Y-%m-%d %H:%M)
df.to_excel(filename+'.xlsx')
