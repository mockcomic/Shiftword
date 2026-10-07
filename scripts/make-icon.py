#!/usr/bin/env python3
"""Generate FreeWord's application icons into build/.

electron-builder is pointed at build/icon.png (see electron-builder.yml); this
script is the source of that asset so the artwork is reproducible instead of a
binary blob someone once pasted in. Run with the Pillow that ships with the
agent/runtime environment:

    python scripts/make-icon.py

Outputs:
  build/icon.png          1024x1024 master (mac + generic)
  build/icon.ico          16..256 px, Windows installer + taskbar
  build/icon.icns         best-effort macOS icon (needs iconutil on macOS; skipped elsewhere)
  build/icons/{16,24,32,48,64,128,256,512}.png
                          Linux / hicolor sizes
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SIZE = 1024
BLUE_TOP = (43, 87, 151, 255)      # Word-ish blue, matching the app's accent
BLUE_BOTTOM = (23, 55, 102, 255)
WHITE = (255, 255, 255, 255)
RULE = (191, 208, 232, 255)

BUILD = Path(__file__).resolve().parent.parent / "build"
ICONS = BUILD / "icons"

FONT_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/segoeuib.ttf",
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def load_font(size: int) -> ImageFont.FreeTypeFont:
    for candidate in FONT_CANDIDATES:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size)
    # Pillow's bitmap default is unusable at 1024 px; better to fail loudly.
    raise SystemExit(
        "No usable TrueType font found. Looked for:\n  " + "\n  ".join(FONT_CANDIDATES)
    )


def rounded_mask(size: int, radius: int) -> Image.Image:
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return mask


def render() -> Image.Image:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

    # Vertical gradient inside a rounded square, 6% of padding left around it
    # so installers that add their own chrome still show the shape whole.
    pad = int(SIZE * 0.06)
    box = SIZE - 2 * pad
    gradient = Image.new("RGBA", (box, box), (0, 0, 0, 0))
    for y in range(box):
        t = y / max(1, box - 1)
        colour = tuple(round(a + (b - a) * t) for a, b in zip(BLUE_TOP, BLUE_BOTTOM))
        ImageDraw.Draw(gradient).line([(0, y), (box, y)], fill=colour)  # type: ignore[arg-type]
    gradient.putalpha(rounded_mask(box, int(box * 0.19)))
    img.paste(gradient, (pad, pad), gradient)

    draw = ImageDraw.Draw(img)

    # The ruled "document" lines: three short bars under the letter, which read
    # as text at 16 px where the letter alone turns to mush.
    bar_x = pad + int(box * 0.26)
    bar_w = int(box * 0.48)
    bar_h = int(box * 0.045)
    for i in range(3):
        y = pad + int(box * (0.66 + 0.085 * i))
        width = bar_w if i < 2 else int(bar_w * 0.62)
        draw.rounded_rectangle(
            [bar_x, y, bar_x + width, y + bar_h], radius=bar_h // 2, fill=RULE
        )

    # Letter F, centred in the upper two thirds.
    font = load_font(int(box * 0.52))
    text = "F"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(
        (pad + (box - tw) / 2 - bbox[0], pad + int(box * 0.40) - th / 2 - bbox[1]),
        text,
        font=font,
        fill=WHITE,
    )
    return img


def main() -> int:
    BUILD.mkdir(parents=True, exist_ok=True)
    ICONS.mkdir(parents=True, exist_ok=True)

    master = render()
    master.save(BUILD / "icon.png")
    master.save(BUILD / "icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    for px in (16, 24, 32, 48, 64, 128, 256, 512):
        master.resize((px, px), Image.LANCZOS).save(ICONS / f"{px}.png")

    print(f"wrote {BUILD / 'icon.png'}")
    print(f"wrote {BUILD / 'icon.ico'}")
    print(f"wrote {ICONS / '*.png'} (8 sizes)")
    print("note: build/icon.icns is produced by electron-builder from icon.png on macOS hosts")
    return 0


if __name__ == "__main__":
    sys.exit(main())
