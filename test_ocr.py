# test_ocr.py — C:\Users\LENOVO\stationdeck\test_ocr.py
# Run: python test_ocr.py

import sys, os, base64, io, json, platform
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pathlib import Path
from PIL import Image, ImageDraw

# ── Fix tesseract path ────────────────────────────────────────
import pytesseract

TESS_PATH = r"C:\Users\LENOVO\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESS_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESS_PATH
    print(f"Tesseract: FOUND at {TESS_PATH}")
else:
    print(f"Tesseract: NOT FOUND at {TESS_PATH}")
    sys.exit(1)

# ── Verify tesseract actually works ───────────────────────────
try:
    ver = pytesseract.get_tesseract_version()
    print(f"Tesseract version: {ver}")
except Exception as e:
    print(f"Tesseract ERROR: {e}")
    sys.exit(1)

print()

# ── OpenAI ────────────────────────────────────────────────────
from openai import OpenAI
from config.settings import OPENAI_API_KEY, OPENAI_MODEL
client = OpenAI(api_key=OPENAI_API_KEY)
print(f"OpenAI model   : {OPENAI_MODEL}")
print(f"API key present: {'YES (' + OPENAI_API_KEY[:8] + '...)' if OPENAI_API_KEY else 'NO'}")
print()

# ── Find all photos ───────────────────────────────────────────
ocr_temp = Path("data/ocr_temp")
photos = sorted(ocr_temp.glob("2026-06-17_Day_*.jpg"))
print(f"Photos found: {len(photos)}")

# ── Detect each sheet (v6 two-pass logic) ────────────────────
from src.ocr_reader import _classify_text

sheet_map = {}
for p in photos:
    try:
        img = Image.open(str(p))
        w, h = img.size

        # Pass 1: top third
        top = img.crop((0, 0, w, h // 3)).convert("L")
        text = pytesseract.image_to_string(top, config="--psm 3").lower()
        stype = _classify_text(text)

        # Pass 2: top 60% if still unknown
        if not stype:
            top2 = img.crop((0, 0, w, int(h * 0.60))).convert("L")
            text2 = pytesseract.image_to_string(top2, config="--psm 3").lower()
            stype = _classify_text(text2) or "unknown"

        sheet_map[stype] = p
        print(f"  {p.name}  ->  {stype}")
    except Exception as e:
        print(f"  {p.name}  ->  ERROR: {e}")

print()

# ── Test on PMS photo ─────────────────────────────────────────
pms_photo = sheet_map.get("pms")
if not pms_photo:
    print("ERROR: No PMS sheet detected among the photos.")
    print("Sheet detection is not working correctly.")
    sys.exit(1)

print(f"PMS photo identified: {pms_photo.name}")
img = Image.open(pms_photo)
W, H = img.size
sx, sy = W / 1927, H / 2560
print(f"Size: {W}x{H}  scale: {sx:.4f},{sy:.4f}")
print()

# ── Crop and save PMS1 fields ─────────────────────────────────
pms1_coords = {
    "pms1_opening_dip":      ( 280, 565,  700, 657),
    "pms1_closing_dip":      ( 750, 565, 1170, 657),
    "pms1_return_tank":      (1220, 565, 1900, 657),
    "pms1_updf_receipt":     ( 280, 742,  700, 832),
    "pms1_updf_consumption": ( 750, 742, 1170, 832),
    "pms1_cost_price":       ( 280, 800,  700, 900),
    "pms1_selling_price":    ( 600, 800, 1920, 900),
}

print("Saving individual crop files — open these to verify coordinates:")
for name, (x1, y1, x2, y2) in pms1_coords.items():
    cx1,cy1 = int(x1*sx), int(y1*sy)
    cx2,cy2 = int(x2*sx), int(y2*sy)
    crop = img.crop((cx1, cy1, cx2, cy2))
    out  = crop.resize((crop.width*3, crop.height*3), Image.LANCZOS)
    fname = f"test_crop_{name}.jpg"
    out.save(fname)
    print(f"  {fname}")

# ── Build tiled image ─────────────────────────────────────────
TILE_W, CROP_H, LABEL_H, GAP = 500, 90, 22, 4
fid_list = list(pms1_coords.keys())
tile_h = (CROP_H + LABEL_H + GAP) * len(fid_list)
tiled = Image.new("RGB", (TILE_W, tile_h), (255, 255, 255))
draw  = ImageDraw.Draw(tiled)
y = 0
for fid in fid_list:
    x1,y1,x2,y2 = pms1_coords[fid]
    crop = img.crop((int(x1*sx),int(y1*sy),int(x2*sx),int(y2*sy))).convert("RGB")
    draw.text((4, y+4), fid, fill=(180,0,0))
    y += LABEL_H
    tiled.paste(crop.resize((TILE_W, CROP_H), Image.LANCZOS), (0, y))
    y += CROP_H + GAP

tiled.save("test_tiled_pms1.jpg")
print()
print("Saved test_tiled_pms1.jpg — open this to verify crops show correct areas")
print()

# ── Call API ──────────────────────────────────────────────────
print("Calling gpt-4o-mini vision API...")
buf = io.BytesIO()
tiled.save(buf, format="JPEG", quality=95)
b64 = base64.b64encode(buf.getvalue()).decode()

prompt = (
    "This image shows cropped handwritten digit fields from a fuel station "
    "meter book (PMS sheet). Each crop is labelled in red with its field name.\n\n"
    "Read the handwritten digits in each crop. Return 0 if blank.\n"
    "Digits only — no commas, spaces, or letters.\n\n"
    "Fields:\n" + "\n".join(f"- {f}" for f in fid_list) + "\n\n"
    "Reply ONLY with JSON. No markdown, no backticks.\n"
    'Example: {"pms1_opening_dip": "315", "pms1_closing_dip": "84"}'
)

try:
    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        max_tokens=400,
        messages=[{"role":"user","content":[
            {"type":"image_url","image_url":{
                "url":f"data:image/jpeg;base64,{b64}","detail":"high"}},
            {"type":"text","text":prompt}
        ]}]
    )
    raw = response.choices[0].message.content
    print(f"Raw response: {raw}")
    print()

    clean = raw.strip().strip("`").strip()
    if clean.lower().startswith("json"): clean = clean[4:].strip()
    values = json.loads(clean)

    expected = {
        "pms1_opening_dip": "315", "pms1_closing_dip": "84",
        "pms1_return_tank": "0", "pms1_updf_receipt":"0",
        "pms1_updf_consumption":"0", "pms1_cost_price": "6467",
        "pms1_selling_price": "6550",
    }
    print("RESULTS:")
    all_ok = True
    for fid in fid_list:
        got = str(values.get(fid, "MISSING"))
        exp = expected.get(fid,"?")
        blank_ok = exp == "0" and got in ["0","","MISSING"]
        ok = got == exp or blank_ok
        if not ok: all_ok = False
        print(f"  {'✓' if ok else '✗'}  {fid:35s}  got={got:12s}  expected={exp}")
    print()
    print("ALL CORRECT ✓" if all_ok else "Coordinate adjustment needed for wrong values.")

except Exception as e:
    print(f"API ERROR: {e}")