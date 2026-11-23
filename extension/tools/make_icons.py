"""產生擴充功能圖示（16/48/128）與商店 logo（300×300）。需 Pillow。"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT = '/System/Library/Fonts/Supplemental/Arial Bold.ttf'


def icon(size):
    s = size * 4
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    grad = Image.new('RGBA', (s, s))
    gd = ImageDraw.Draw(grad)
    for y in range(s):
        k = y / s
        gd.line([(0, y), (s, y)], fill=(int(26 + 20 * k), int(58 + 52 * k), int(92 + 70 * k), 255))
    mask = Image.new('L', (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22), fill=255)
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)
    if size <= 16:
        f = ImageFont.truetype(FONT, int(s * 0.62))
        d.text((s / 2, s / 2), 'AF', font=f, fill='white', anchor='mm')
    else:
        f = ImageFont.truetype(FONT, int(s * 0.46))
        d.text((s / 2, s * 0.42), 'AF', font=f, fill='white', anchor='mm')
        # 下方箭頭：代表「送進 aftool」
        y = s * 0.76; x0, x1 = s * 0.24, s * 0.76; w = max(4, int(s * 0.06))
        d.line([(x0, y), (x1 - s * 0.04, y)], fill=(255, 214, 102, 255), width=w)
        d.polygon([(x1, y), (x1 - s * 0.13, y - s * 0.09), (x1 - s * 0.13, y + s * 0.09)], fill=(255, 214, 102, 255))
    return img.resize((size, size), Image.LANCZOS)


if __name__ == '__main__':
    for n in (16, 48, 128):
        icon(n).save(os.path.join(ROOT, 'edge', 'icons', f'icon{n}.png'))
    icon(300).save(os.path.join(ROOT, 'store', 'logo-300x300.png'))
    print('ok')
