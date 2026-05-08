from PIL import Image, ImageDraw
import os

def make_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Draw gradient background row by row
    for y in range(size):
        t = y / size
        g = int(0x78 + t * (0x5a - 0x78))
        b = int(0xd4 + t * (0x9e - 0xd4))
        draw.line([(0, y), (size, y)], fill=(0, g, b, 255))

    # Apply rounded corners mask
    r = int(size * 0.17)
    mask = Image.new("L", (size, size), 0)
    mdraw = ImageDraw.Draw(mask)
    mdraw.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=255)
    img.putalpha(mask)

    draw = ImageDraw.Draw(img)
    s = size / 128.0

    def rr(x, y, w, h, rx, color, alpha=255):
        col = color + (alpha,)
        draw.rounded_rectangle(
            [x * s, y * s, (x + w) * s, (y + h) * s],
            radius=max(1, int(rx * s)),
            fill=col
        )

    # Left column (source env)
    rr(18, 28, 36, 7, 3, (255, 255, 255), 242)
    for y in [42, 52, 62, 72, 82]:
        rr(18, y, 36, 5, 2, (255, 255, 255), 140)
    rr(18, 92, 28, 5, 2, (255, 255, 255), 89)

    # Right column (target env)
    rr(74, 28, 36, 7, 3, (255, 255, 255), 242)
    rr(74, 42, 36, 5, 2, (255, 255, 255), 140)
    rr(74, 52, 36, 5, 2, (125, 211, 252), 217)   # blue = diff
    rr(74, 62, 36, 5, 2, (255, 255, 255), 140)
    rr(74, 72, 36, 5, 2, (252, 165, 165), 217)   # red = missing
    rr(74, 82, 36, 5, 2, (255, 255, 255), 140)
    rr(74, 92, 28, 5, 2, (255, 255, 255), 89)

    # Centre double-headed arrow
    cx = 64 * s
    cy = 64 * s
    aw = 8 * s
    lw = max(1, int(2.5 * s))
    draw.line([(cx - aw, cy), (cx + aw, cy)], fill=(255, 255, 255, 220), width=lw)
    draw.polygon([(cx + aw - 4*s, cy - 4*s), (cx + aw + 2*s, cy), (cx + aw - 4*s, cy + 4*s)],
                 fill=(255, 255, 255, 220))
    draw.polygon([(cx - aw + 4*s, cy - 4*s), (cx - aw - 2*s, cy), (cx - aw + 4*s, cy + 4*s)],
                 fill=(255, 255, 255, 220))

    return img


outdir = os.path.join(os.path.dirname(__file__))
for size in [16, 32, 48, 128]:
    ico = make_icon(size)
    path = os.path.join(outdir, f"icon{size}.png")
    ico.save(path, "PNG")
    print(f"Saved {path}")

print("Done.")
