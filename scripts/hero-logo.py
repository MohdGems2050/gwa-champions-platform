"""
Replaces the small GEMS badge burned into the LED screen of public/hero-student.png
with the official GEMS World Academy Dubai logo (public/gems-logo.png).

Placement (1408x768 source): the badge sits on the LED screen in the upper-right
area, above the "ARABIC & ISLAMIC CHAMPIONS" headline — old bounding box
x 938..1160, y 106..190. The official logo is pasted on a clean white rounded
plate centred on that box; the logo artwork itself is never recoloured, cropped
or stretched (uniform scale, original colours preserved).

Run:  python3 scripts/hero-logo.py
"""
from PIL import Image, ImageDraw, ImageFilter
import numpy as np

SRC   = "public/hero-student.png"
LOGO  = "public/gems-logo.png"
OUT   = "public/hero-student.png"

bx0, by0, bx1, by1 = 938, 106, 1160, 190
cx, cy = (bx0 + bx1) // 2, (by0 + by1) // 2

hero = Image.open(SRC).convert("RGB")
logo = Image.open(LOGO).convert("RGB")

# 1) erase the old badge by blending the screen gradient
pad = 6
fx0, fy0, fx1, fy1 = bx0 - pad, by0 - pad, bx1 + pad, by1 + pad
arr = np.asarray(hero).astype(float).copy()
for y in range(fy0, fy1 + 1):
    left = arr[y, fx0 - 16:fx0 - 5].mean(axis=0)
    right = arr[y, fx1 + 6:fx1 + 17].mean(axis=0)
    span = fx1 - fx0
    for i in range(span + 1):
        t = i / span
        arr[y, fx0 + i] = left * (1 - t) + right * t
tmp = Image.fromarray(arr.clip(0, 255).astype("uint8"))
base = tmp.copy()
base.paste(tmp.crop((fx0 - 3, fy0 - 3, fx1 + 4, fy1 + 4)).filter(ImageFilter.GaussianBlur(2.0)),
           (fx0 - 3, fy0 - 3))

# 2) white plate + official logo
pw, ph = (bx1 - bx0) + 12, (by1 - by0) + 8
box = (cx - pw // 2, cy - ph // 2)

shadow = Image.new("RGBA", (pw + 56, ph + 56), (0, 0, 0, 0))
ImageDraw.Draw(shadow).rounded_rectangle((28, 32, 28 + pw, 32 + ph), radius=15, fill=(0, 8, 22, 120))
shadow = shadow.filter(ImageFilter.GaussianBlur(11))

glow = Image.new("RGBA", (pw + 56, ph + 56), (0, 0, 0, 0))
ImageDraw.Draw(glow).rounded_rectangle((28, 28, 28 + pw, 28 + ph), radius=18, fill=(120, 220, 240, 55))
glow = glow.filter(ImageFilter.GaussianBlur(10))

plate = Image.new("RGBA", (pw, ph), (0, 0, 0, 0))
dp = ImageDraw.Draw(plate)
dp.rounded_rectangle((0, 0, pw - 1, ph - 1), radius=14, fill=(255, 255, 255, 255))
dp.rounded_rectangle((1, 1, pw - 2, ph - 2), radius=13, outline=(214, 236, 244, 255), width=1)

content = logo.crop((62, 150, 1348, 619))          # artwork only (white margin trimmed)
scale = min((pw - 16) / content.width, (ph - 12) / content.height)
nw, nh = max(1, round(content.width * scale)), max(1, round(content.height * scale))
art = content.resize((nw, nh), Image.LANCZOS)
alpha = np.clip((252 - np.asarray(art).astype(int).min(axis=2)) / 18.0 * 255, 0, 255).astype("uint8")
art_rgba = art.convert("RGBA"); art_rgba.putalpha(Image.fromarray(alpha))
plate.paste(art_rgba, ((pw - nw) // 2, (ph - nh) // 2), art_rgba)

out = base.convert("RGBA")
out.alpha_composite(shadow, (box[0] - 28, box[1] - 28))
out.alpha_composite(glow,   (box[0] - 28, box[1] - 28))
out.alpha_composite(plate,  box)
out.convert("RGB").save(OUT, optimize=True)
print("written", OUT, out.size)
