import pyautogui
import pygetwindow
import pywinauto
import cv2
import pytesseract
import pandas as pd
import time
import sys
from PIL import ImageGrab

# https://github.com/PowercoderJr/lords-gifts-counter/blob/master/analyze.py
# LDPlayer 960*540 160dpi english

# 변수 선언
tesseract_config = '-c tessedit_char_whitelist=" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" --oem 3 --psm 7'
rarities = ['Common', 'Uncommon', 'Rare', 'Epic', 'Legendary']
rarity_text = ''
nickname_text = ''
nickname_list = []

df = pd.DataFrame(columns=rarities)
error_count = 0
filename = time.strftime('%y%m%d-%H%M') # 시작 시간으로 파일 이름 저장
i = 0

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
ImageGrab.grabclipboard().save('check.png', 'png')
ImageGrab.grabclipboard().crop((550, 160, 630, 182)).save('gift.png', 'png')
gift_gray = cv2.imread('gift.png', cv2.IMREAD_GRAYSCALE)
gift_gray = cv2.threshold(gift_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
gift_text = pytesseract.image_to_string(gift_gray, lang='eng', config=tesseract_config)
gift_text = int(gift_text[6:len(gift_text)])

if gift_text > 0:
    print('전체 상자 수:', gift_text)
else:
    print('상자가 없음')
    sys.exit()

while True:
    if rarity_text == 'Be' and nickname_text == 'Be':
        break

    time.sleep(0.8)

    # 상자 남았는 지 체크
    pyautogui.hotkey('alt', 'printscreen')
    ImageGrab.grabclipboard().crop((550, 300, 730, 500)).save('check.png', 'png')
    check_gray = cv2.imread('check.png', cv2.IMREAD_GRAYSCALE)
    check_gray = cv2.threshold(check_gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
    check_text = pytesseract.image_to_string(check_gray, lang='eng', config=tesseract_config).strip('\n')

    if check_text == 'No Guild Gifts':
        print('더이상 상자가 없음')
        break
    else:
        i += 1

    # 활성화 창 클립보드로 스샷 찍고 가공 후 저장
    pyautogui.hotkey('alt', 'printscreen')
    im1 = ImageGrab.grabclipboard()
    im1.crop((380, 240, 680, 320)).save(('./data/' + str(i) + '.png'), 'png')
    im1.crop((385, 242, 700, 266)).save('rarity.png', 'png')
    im1.crop((515, 268, 700, 293)).save('nickname.png', 'png')

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
        time.sleep(0.7)
        pyautogui.click()
    else:
        time.sleep(2)


print('전체 실행 횟수:', i )

# 데이터프레임 가공하기 (총 상자 수, 총 포인트)
Total_gifts = (df['Common'].sum(), df['Uncommon'].sum(), df['Rare'].sum(), df['Epic'].sum(), df['Legendary'].sum())
df3 = pd.DataFrame(index=['Total'], columns=rarities, data=[Total_gifts])
df = pd.concat([df, df3])
df['Points'] = df['Common']*1 + df['Uncommon']*5 + df['Rare']*15 + df['Epic']*30 + df['Legendary']*50

# 엑셀로 저장하기
df.to_excel(filename+'.xlsx')
