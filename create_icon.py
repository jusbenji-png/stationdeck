"""
create_icon.py — StationDeck Icon Generator
Run once from the project root:
    python create_icon.py

Output: web/static/favicon.ico
Requires: Pillow (already installed in venv)
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path


def make_frame(size: int) -> Image.Image:
    """
    Draw at 4x resolution, scale down with LANCZOS for clean antialiasing.
    RGB mode — no alpha channel — prevents dark halo on Windows desktop.
    Rounded corners drawn directly (not via mask) to avoid black corner bleed.
    """
    scale     = 4
    draw_size = size * scale
    bg_color  = (26, 26, 46)    # #1A1A2E — dark navy
    red_color = (200, 16, 46)   # #C8102E — StationDeck red

    if size >= 64:
        # Start with black canvas, paint rounded navy rect on top
        # Corners outside the rounded rect stay black — clean hard edge
        img  = Image.new("RGB", (draw_size, draw_size), (0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle(
            [0, 0, draw_size - 1, draw_size - 1],
            radius=draw_size // 7,
            fill=bg_color,
        )
    else:
        # Small sizes: full navy square — no rounded corners at small sizes
        img  = Image.new("RGB", (draw_size, draw_size), bg_color)
        draw = ImageDraw.Draw(img)

    # Bold "S" — try Windows fonts first, fall back to DejaVu
    font_size  = int(draw_size * 0.60)
    font       = None
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf",
        "arialbd.ttf",
        "C:/Windows/Fonts/calibrib.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "DejaVuSans-Bold.ttf",
    ]
    for candidate in candidates:
        try:
            font = ImageFont.truetype(candidate, font_size)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()

    bbox   = draw.textbbox((0, 0), "S", font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (draw_size - text_w) // 2 - bbox[0]
    y = (draw_size - text_h) // 2 - bbox[1] - int(draw_size * 0.04)
    draw.text((x, y), "S", fill=(255, 255, 255), font=font)

    # Red underline bar
    bar_w = int(draw_size * 0.38)
    bar_h = max(4, int(draw_size * 0.055))
    bar_x = (draw_size - bar_w) // 2
    bar_y = int(draw_size * 0.73)
    draw.rounded_rectangle(
        [bar_x, bar_y, bar_x + bar_w, bar_y + bar_h],
        radius=bar_h // 2,
        fill=red_color,
    )

    # Scale down with LANCZOS for smooth antialiasing
    img = img.resize((size, size), Image.LANCZOS)
    return img


def main():
    sizes  = [16, 24, 32, 48, 64, 128, 256]
    frames = [make_frame(s) for s in sizes]

    out_dir = Path(__file__).parent / "web" / "static"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_path = out_dir / "favicon.ico"
    frames[0].save(
        str(out_path),
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=frames[1:],
    )
    print(f"[OK] Icon saved:  {out_path}")

    preview_path = out_dir / "favicon_preview.png"
    frames[-1].save(str(preview_path))
    print(f"[OK] Preview:     {preview_path}")
    print()
    print("Next steps:")
    print("  1. Run:  rmdir /s /q build dist")
    print("  2. Run:  build.bat")
    print("  3. Press F9 in Inno Setup")
    print("  4. Delete old desktop shortcut")
    print("  5. Reinstall StationDeck")
    print("  6. Clear icon cache (run as Administrator):")
    print("     taskkill /f /im explorer.exe")
    print('     del /f /q "%localappdata%\\IconCache.db"')
    print('     del /f /q "%localappdata%\\Microsoft\\Windows\\Explorer\\iconcache_*.db"')
    print('     del /f /q "%localappdata%\\Microsoft\\Windows\\Explorer\\thumbcache_*.db"')
    print("     ie4uinit.exe -show")
    print("     start explorer.exe")


if __name__ == "__main__":
    main()