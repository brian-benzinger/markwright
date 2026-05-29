"""Generate the Markwright ribbon icons.

Renders a rounded-square Word-blue tile with a white M and a downward
arrow underneath — combining Markdown's familiar M-down lockup with
Word's brand color. Outputs the four sizes the manifest references
(16, 32, 64, 80). Run from the repo root:

    python3 scripts/make-icons.py
"""

from __future__ import annotations

import pathlib
from PIL import Image, ImageDraw, ImageFont

OUT = pathlib.Path(__file__).resolve().parent.parent / "src" / "assets"

# Use a font that ships on common dev environments. Any clean bold sans
# would do; the binary-search step below sizes it to fit regardless.
FONT_PATH = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"

WORD_BLUE = (43, 87, 154, 255)  # #2B579A
WHITE = (255, 255, 255, 255)
SIZES = (16, 32, 64, 80)


def fit_font(target_height: int, font_path: str) -> ImageFont.FreeTypeFont:
    """Return the largest font size whose 'M' glyph fits target_height."""
    lo, hi = 4, target_height * 2
    best = ImageFont.truetype(font_path, lo)
    while lo <= hi:
        mid = (lo + hi) // 2
        font = ImageFont.truetype(font_path, mid)
        bbox = font.getbbox("M")
        h = bbox[3] - bbox[1]
        if h <= target_height:
            best = font
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def make_icon(size: int) -> Image.Image:
    # Oversample so the rounded corners and triangle edges stay crisp
    # after downscaling.
    scale = 8 if size <= 32 else 4
    big = size * scale
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Rounded square background with a small transparent gutter so the
    # corners read well on light and dark ribbon themes.
    pad = max(1, int(big * 0.04))
    d.rounded_rectangle(
        [pad, pad, big - pad, big - pad],
        radius=int(big * 0.18),
        fill=WORD_BLUE,
    )

    # Sized M centered horizontally, sitting in the upper half so the
    # arrow has room below.
    target_m_height = int(big * 0.50)
    font = fit_font(target_m_height, FONT_PATH)
    bbox = font.getbbox("M")
    text_w = bbox[2] - bbox[0]
    x = (big - text_w) // 2 - bbox[0]
    y = int(big * 0.16) - bbox[1]
    d.text((x, y), "M", fill=WHITE, font=font)

    # Downward triangle (the "down" half of the M-down lockup).
    arrow_w = int(big * 0.30)
    arrow_h = int(big * 0.18)
    arrow_cx = big // 2
    arrow_top_y = int(big * 0.70)
    d.polygon(
        [
            (arrow_cx - arrow_w // 2, arrow_top_y),
            (arrow_cx + arrow_w // 2, arrow_top_y),
            (arrow_cx, arrow_top_y + arrow_h),
        ],
        fill=WHITE,
    )

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for size in SIZES:
        path = OUT / f"icon-{size}.png"
        make_icon(size).save(path, optimize=True)
        print(f"wrote {path}")


if __name__ == "__main__":
    main()
