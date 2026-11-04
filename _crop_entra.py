import sys
from PIL import Image

src, dest = sys.argv[1], sys.argv[2]
im = Image.open(src).convert("RGB")
w, h = im.size
px = im.load
top = left = None
for y in range(80, 280):
 dark = sum(
 1
 for x in range(0, w, 8)
 if px[x, y][0] < 35 and px[x, y][1] < 35 and px[x, y][2] < 35
 )
 if dark > 100:
 top = y
 for x in range(w):
 r, g, b = px[x, y]
 if r < 30 and g < 30 and b < 30:
 left = x
 break
 break
if top is None or left is None:
 raise SystemExit("entra window not found")
bottom = h - 40
for y in range(h - 1, top, -1):
 light = sum(1 for x in range(left, w, 12) if sum(px[x, y]) > 400)
 if light > 40:
 bottom = y
 break
im.crop((left, top, w - 4, bottom + 1)).save(dest)
print("cropped", left, top, w - 4, bottom)
