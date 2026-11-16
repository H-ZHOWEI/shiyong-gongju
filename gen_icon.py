# -*- coding: utf-8 -*-
"""生成「实用工具」应用图标：深蓝绿渐变圆角底 + 白色齿轮（工具合集）。"""
import math, os
from PIL import Image, ImageDraw

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
S = 1024  # 高清母版尺寸

def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))

# 渐变底：左上深蓝灰（界面主色 #2c3e50）→ 右下青绿（强调色 #4CAF50）
c_top = (44, 62, 80)
c_bottom = (38, 145, 120)
base = Image.new('RGBA', (S, S), (0, 0, 0, 0))
grad = Image.new('RGB', (S, S))
px = grad.load()
for y in range(S):
    for x in range(S):
        t = max(0.0, min(1.0, (x + y) / (2 * S)))
        px[x, y] = lerp(c_top, c_bottom, t ** 1.15)
base.paste(grad, (0, 0))

# 圆角方形遮罩
mask = Image.new('L', (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle((52, 52, S - 52, S - 52), radius=225, fill=255)
base.putalpha(mask)

# 齿轮（8 齿），白色
d = ImageDraw.Draw(base)
cx = cy = S // 2
r_out, r_in, r_hole = 335, 255, 112
teeth = 8
step = 2 * math.pi / teeth
pts = []
for j in range(teeth):
    a = j * step - math.pi / 2
    for frac, r in ((0.00, r_in), (0.16, r_out), (0.34, r_out), (0.50, r_in)):
        ang = a + frac * step
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
d.polygon(pts, fill=(255, 255, 255, 255))
# 中心孔：透出底色
d.ellipse((cx - r_hole, cy - r_hole, cx + r_hole, cy + r_hole),
          fill=lerp(c_top, c_bottom, 0.5) + (255,))
# 中心绿点，呼应软件主题色
d.ellipse((cx - 52, cy - 52, cx + 52, cy + 52), fill=(76, 175, 80, 255))

# 母版 PNG
png_path = os.path.join(OUT_DIR, 'app_icon.png')
base.save(png_path)

# Windows ICO：多尺寸
sizes = [16, 24, 32, 48, 64, 128, 256]
ico_path = os.path.join(OUT_DIR, 'app_icon.ico')
base.save(ico_path, format='ICO', sizes=[(s, s) for s in sizes])
print('saved:', png_path, ico_path, os.path.getsize(ico_path))
