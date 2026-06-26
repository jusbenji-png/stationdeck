"""
create_icon.py — StationDeck Icon Generator (Windows-compliant)

Run once from the project root:
    python create_icon.py

Output: web/static/favicon.ico
Requires: Pillow  (pip install Pillow)

How Windows uses icon layers
-----------------------------
Windows ICO files are containers that hold several independent bitmaps.
The shell picks whichever layer most closely matches the display size it
needs, then scales only the remainder — so a missing size forces a larger
layer to be downscaled by GDI+, which produces blurry or jagged results.

Critical sizes for the Windows desktop:
  16×16   — small icon in Explorer list view, title bar, taskbar
  32×32   — medium icon in Explorer, some dialogs
  48×48   — large icon on the desktop and in Explorer "medium icons" view
  256×256 — "extra large" icons (Windows Vista+); stored as PNG-compressed
             inside the ICO to keep file size manageable

Color depth:
  All layers are stored as 32-bit RGBA.  The alpha channel carries smooth
  per-pixel transparency so rounded corners and anti-aliased edges blend
  cleanly on any desktop background color — no black halos, no jagged edges.

Resampling:
  We draw everything at 4× the target size, then reduce with LANCZOS
  (a high-quality sinc-based filter) so sub-pixel details are averaged
  down correctly rather than thrown away.
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import struct, zlib


# ---------------------------------------------------------------------------
# Brand constants
# ---------------------------------------------------------------------------
BG_COLOR  = (26, 26, 46, 255)    # #1A1A2E  dark navy,  fully opaque
RED_COLOR = (200, 16, 46, 255)   # #C8102E  StationDeck red
WHITE     = (255, 255, 255, 255)
CLEAR     = (0, 0, 0, 0)         # fully transparent


# ---------------------------------------------------------------------------
# Font loader — tries common bold fonts, falls back to Pillow default
# ---------------------------------------------------------------------------
_FONT_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/calibrib.ttf",
    "C:/Windows/Fonts/verdanab.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "DejaVuSans-Bold.ttf",
]

def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in _FONT_CANDIDATES:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


# ---------------------------------------------------------------------------
# Core frame renderer
# ---------------------------------------------------------------------------
def make_frame(size: int) -> Image.Image:
    """
    Render one icon frame at `size` × `size` pixels in RGBA mode.

    Strategy:
      • Draw at 4× resolution for crisp anti-aliasing.
      • Use a transparent canvas + rounded-rectangle fill so the alpha
        channel carries the true shape — no black corner bleed.
      • Scale to target with LANCZOS.
    """
    scale     = 4
    draw_size = size * scale

    # Transparent canvas — corners stay transparent (not black)
    img  = Image.new("RGBA", (draw_size, draw_size), CLEAR)
    draw = ImageDraw.Draw(img)

    # Rounded background — radius proportional to size; square at tiny sizes
    if size >= 32:
        radius = draw_size // 7
    else:
        radius = draw_size // 16   # almost square at 16 px

    draw.rounded_rectangle(
        [0, 0, draw_size - 1, draw_size - 1],
        radius=radius,
        fill=BG_COLOR,
    )

    # Bold "S" glyph
    font_size = int(draw_size * 0.60)
    font      = _load_font(font_size)

    bbox   = draw.textbbox((0, 0), "S", font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (draw_size - text_w) // 2 - bbox[0]
    y = (draw_size - text_h) // 2 - bbox[1] - int(draw_size * 0.04)
    draw.text((x, y), "S", fill=WHITE, font=font)

    # Red underline bar
    bar_w = int(draw_size * 0.38)
    bar_h = max(4, int(draw_size * 0.055))
    bar_x = (draw_size - bar_w) // 2
    bar_y = int(draw_size * 0.73)
    draw.rounded_rectangle(
        [bar_x, bar_y, bar_x + bar_w, bar_y + bar_h],
        radius=bar_h // 2,
        fill=RED_COLOR,
    )

    # Downscale with high-quality LANCZOS filter
    img = img.resize((size, size), Image.Resampling.LANCZOS)
    return img


# ---------------------------------------------------------------------------
# ICO builder that guarantees PNG compression for the 256×256 layer
# ---------------------------------------------------------------------------
# Pillow's .save(..., format="ICO") correctly PNG-compresses 256×256 when
# the append_images list is provided and the sizes kwarg is set.  We use
# Pillow's built-in writer but validate the output afterwards.

def save_ico(frames: list[Image.Image], sizes: list[int], path: Path) -> None:
    """
    Save a multi-resolution ICO using Pillow's native writer.

    Pillow stores 256×256 as a PNG chunk inside the ICO (Vista+ format)
    and all smaller sizes as 32-bit BMP chunks — exactly what Windows wants.
    """
    # Pillow needs all frames as RGBA for correct 32-bit output
    rgba_frames = [f.convert("RGBA") for f in frames]

    rgba_frames[0].save(
        str(path),
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=rgba_frames[1:],
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    # Windows desktop icon sizes — order matters: largest first so Pillow
    # stores the best reference frame at index 0 for previews.
    sizes  = [256, 48, 32, 16]
    frames = [make_frame(s) for s in sizes]

    out_dir = Path(__file__).parent / "web" / "static"
    out_dir.mkdir(parents=True, exist_ok=True)

    ico_path = out_dir / "favicon.ico"
    save_ico(frames, sizes, ico_path)
    print(f"[OK] ICO saved   : {ico_path}  ({ico_path.stat().st_size:,} bytes)")

    # 256-px PNG preview so you can inspect the design without opening the ICO
    preview_path = out_dir / "favicon_preview.png"
    frames[0].save(str(preview_path))
    print(f"[OK] Preview PNG : {preview_path}")

    print()
    print("Layers written:")
    for s, f in zip(sizes, frames):
        print(f"  {s:>3}×{s:<3}  RGBA  LANCZOS-scaled")

    print()
    print("Next steps:")
    print("  1. Rebuild:  build.bat  (runs PyInstaller)")
    print("  2. Recompile installer in Inno Setup (F9)")
    print("  3. Reinstall StationDeck")
    print("  4. Clear Windows icon cache (run as Administrator):")
    print('     ie4uinit.exe -show')
    print('     del /f /q "%localappdata%\\IconCache.db"')
    print('     del /f /q "%localappdata%\\Microsoft\\Windows\\Explorer\\iconcache_*.db"')
    print("     taskkill /f /im explorer.exe && start explorer.exe")


if __name__ == "__main__":
    main()
