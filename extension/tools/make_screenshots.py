"""把 e2e_extension.py（SHOTS=目錄）擷取的模擬畫面合成商店截圖（1280×800）。"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

SHOTS = sys.argv[1]
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'store')
FB = '/System/Library/Fonts/STHeiti Medium.ttc'
FL = '/System/Library/Fonts/STHeiti Light.ttc'
NAVY = (26, 58, 92)


def font(path, size):
    return ImageFont.truetype(path, size)


def frame(img, w):
    h = int(img.height * w / img.width)
    img = img.resize((w, h), Image.LANCZOS)
    box = Image.new('RGB', (w + 4, h + 4), (180, 190, 205))
    box.paste(img, (2, 2))
    return box


def note(d, top=False):
    xy, anchor = ((1260, 14), 'ra') if top else ((1260, 785), 'rs')
    d.text(xy, '示意畫面：模擬 AF 頁面與虛構資料', font=font(FL, 15), fill=(120, 130, 145), anchor=anchor)


def overview():
    c = Image.new('RGB', (1280, 800), (244, 247, 251))
    d = ImageDraw.Draw(c)
    d.rectangle([0, 0, 1280, 120], fill=NAVY)
    d.text((40, 38), 'AF → aftool 一鍵匯入', font=font(FB, 40), fill='white')
    d.text((42, 88), '在 AF 用人費用管理資訊系統按一下，校對清冊(A) 直接進 aftool，免下載、免拖檔', font=font(FL, 20), fill=(220, 230, 240))
    left = frame(Image.open(os.path.join(SHOTS, 'shot-popup-done.png')), 590)
    right = frame(Image.open(os.path.join(SHOTS, 'shot-aftool-imported.png')), 590)
    c.paste(left, (36, 190)); c.paste(right, (650, 190))
    d.text((36, 150), '① AF 系統「報表列印」：自動選校對清冊(A)、EXCEL、填密碼', font=font(FB, 19), fill=NAVY)
    d.text((650, 150), '② aftool：自動解開並放進 AF 資料檔', font=font(FB, 19), fill=NAVY)
    y = 190 + max(left.height, right.height) + 28
    for i, t in enumerate(['只在 fund.dgpa.gov.tw 與 aftool.leotsai.me 運作', '不讀取帳號、密碼與 cookie，不儲存任何資料', '多校：切換「兼辦機關」後再按一次']):
        d.ellipse([40, y + 9 + i * 34, 50, y + 19 + i * 34], fill=(42, 110, 63))
        d.text((60, y + i * 34), t, font=font(FL, 19), fill=(40, 60, 80))
    note(d)
    c.save(os.path.join(OUT, 'screenshot-1-overview-1280x800.png'))


def single(src, dst, caption):
    img = Image.open(os.path.join(SHOTS, src)).convert('RGB').resize((1280, 800))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 1280, 0], fill=NAVY)
    d.rounded_rectangle([430, 300, 1250, 360], radius=10, fill=(26, 58, 92))
    d.text((840, 330), caption, font=font(FB, 22), fill='white', anchor='mm')
    note(d, top=True)
    img.save(os.path.join(OUT, dst))


if __name__ == '__main__':
    overview()
    single('shot-popup-done.png', 'screenshot-2-af-1280x800.png', '按右下角按鈕，報表選項與密碼自動填好並取得檔案')
    img = Image.open(os.path.join(SHOTS, 'shot-aftool-imported.png')).convert('RGB').resize((1280, 800))
    d = ImageDraw.Draw(img); note(d)
    img.save(os.path.join(OUT, 'screenshot-3-aftool-1280x800.png'))
    print('ok')
