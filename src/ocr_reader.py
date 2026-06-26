# =============================================================
# src/ocr_reader.py
# StationDeck — OCR Capture Sheet Reader  (v12)
#
# Sheet detection  : pytesseract reads the printed sheet title
# Digit extraction : gpt-4o-mini vision via openai library
#
# Coordinate history:
#   v3: initial estimate
#   v4: +207px y offset applied to all meter book fields
#   v5: fine-tuned per-row based on real crop inspection:
#       - top 3 fields (dips) confirmed correct, unchanged
#       - updf row: +30px down
#       - price row: +28px down, selling_price x extended to 1870
#   v12: REVERTED v11 per-box transcription — forcing one char per box made GPT
#        hallucinate digits in blank boxes (empty PMS2/3/4 tanks showed fake
#        numbers). Back to "read the handwritten digits" but with explicit
#        "blank boxes are not zeros / do not pad" rules. Kept: image enhancement,
#        box-count overflow flag (digits > boxes → warn), tighter validation.
#   v11: per-box transcription (reverted — see v12).
#   v10: AGO cost_price x1 500→300, return_tank x1 1220→1080, selling_price y+x fixed;
#        PMS return_tank x1 1220→1080, selling_price x1 extended;
#        TILE_W 500→700, CROP_H 90→110; prompt completeness rules; better validation notes
#   v9: all sheet coordinates calibrated and passing (PMS/AGO/cash2/products/expenses)
#   v8: AGO calibrated from WA0036 — different y offsets than PMS (zoomed-out photo)
#   v7: fix updf row coords (was hitting Cost Price label row),
#       selling_price x start moved 600→750 (removes left-margin GPT confusion)
#   v6: improved _detect_sheet_type():
#       - two-pass crop: top-third first, then top-60% if still unknown
#       - loosened keyword matching (words checked independently)
#       - "expenses" alone now matches Daily Expenses sheet
#       - "product"+"sales" checked independently (handles partial title)
#       - "lubricants" added as secondary products keyword
# =============================================================

import base64
import io
import json
import logging
import os
import platform
from pathlib import Path

logger = logging.getLogger(__name__)

# ── Tesseract path ────────────────────────────────────────────
def _find_tesseract():
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\LENOVO\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
    ]
    appdata = os.environ.get("LOCALAPPDATA", "")
    if appdata:
        candidates.append(
            os.path.join(appdata, "Programs", "Tesseract-OCR", "tesseract.exe")
        )
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

# ── Pillow + pytesseract ──────────────────────────────────────
try:
    import pytesseract
    from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

    if platform.system() == "Windows":
        tess_path = _find_tesseract()
        if tess_path:
            pytesseract.pytesseract.tesseract_cmd = tess_path
            logger.info(f"Tesseract found at: {tess_path}")
        else:
            logger.warning("tesseract.exe not found — sheet detection will fail.")

    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    logger.warning("pytesseract or Pillow not available — OCR disabled.")

# ── OCR backend ───────────────────────────────────────────────
# Primary path: route the vision call through the StationDeck server so the
# OpenAI key lives ONLY on the server. Stations never need their own key.
# Fallback path: if no station credentials are supplied (e.g. local dev /
# test_ocr.py) and a local OPENAI_API_KEY exists, call OpenAI directly.
OCR_SERVER_URL = "https://web-production-46077.up.railway.app/ocr"

try:
    import requests
    _REQUESTS_OK = True
except ImportError:
    _REQUESTS_OK = False

try:
    from openai import OpenAI
    from config.settings import OPENAI_API_KEY, OPENAI_MODEL
    _openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
except Exception as e:
    _openai_client = None
    logger.warning(f"Local OpenAI client not available (will use server OCR): {e}")

# Vision works if we can reach the server (requests) OR call OpenAI locally.
VISION_AVAILABLE = _REQUESTS_OK or (_openai_client is not None)

# Station credentials for server-side OCR — set per call by read_capture_sheets().
_OCR_STATION_NAME = None
_OCR_MACHINE_ID   = None


# =============================================================
# REFERENCE RESOLUTION
# =============================================================
REF_W = 1927
REF_H = 2560

# =============================================================
# FIELD COORDINATE MAPS
# =============================================================
# (x1, y1, x2, y2) in a 1927×2560 reference image.
# Calibrated from real TotalEnergies Rwizi photos, v5 final.

TANK_STEP_PMS = 395
TANK_STEP_AGO = 395

_PMS_TANK1 = {
    # Dip readings row — CONFIRMED CORRECT from crop inspection
    "opening_dip":      ( 280, 565,  700, 657),
    "closing_dip":      ( 750, 565, 1170, 657),
    "return_tank":      (1080, 565, 1920, 657),  # x1 extended left: 1220→1080
    # UPDF row — civilian station, always blank; anchor to confirmed-blank
    # zone inside return_tank row so GPT sees empty boxes, not printed text
    "updf_receipt":     (1300, 565, 1600, 657),
    "updf_consumption": (1600, 565, 1900, 657),
    # Price row — cost_price unchanged (confirmed ✅), selling_price x start
    # moved from 600 to 750 to drop left-margin content that caused GPT misread
    "cost_price":       ( 280, 800,  700, 900),
    "selling_price":    ( 900, 800, 1920, 900),  # x1 extended left: 1050→900
}

# AGO sheet — calibrated from capture template photos (2026-06-17).
# Photo is wider/zoomed-out vs PMS, so all rows are ~110px higher.
# v10 fixes: return_tank x1 1220→1080, cost_price x1 500→300,
#            selling_price y corrected (was 650-740, now 685-775 to match cost row)
#            selling_price x1 extended 1050→900
_AGO_TANK1 = {
    "opening_dip":      ( 200, 460,  800, 555),
    "closing_dip":      ( 800, 460, 1200, 555),
    "return_tank":      (1080, 460, 1920, 555),  # x1 extended left: 1220→1080
    "updf_receipt":     (1300, 460, 1600, 555),  # blank — civilian station
    "updf_consumption": (1600, 460, 1870, 555),  # blank — civilian station
    "cost_price":       ( 300, 685,  840, 775),  # x1 extended left: 500→300
    "selling_price":    ( 900, 685, 1920, 775),  # x1 900, y aligned with cost row
}

def _meter_book_coords(product: str) -> dict:
    base = _PMS_TANK1 if product == "pms" else _AGO_TANK1
    step = TANK_STEP_PMS if product == "pms" else TANK_STEP_AGO
    coords = {}
    for tank in range(1, 5):
        dy = (tank - 1) * step
        for field, (x1, y1, x2, y2) in base.items():
            coords[f"{product}{tank}_{field}"] = (x1, y1 + dy, x2, y2 + dy)
    return coords


def _meter_book_coords_tank(product: str, tank: int) -> dict:
    """Return coord map for a single tank only (used for per-tank API calls)."""
    base = _PMS_TANK1 if product == "pms" else _AGO_TANK1
    step = TANK_STEP_PMS if product == "pms" else TANK_STEP_AGO
    dy = (tank - 1) * step
    return {f"{product}{tank}_{field}": (x1, y1 + dy, x2, y2 + dy)
            for field, (x1, y1, x2, y2) in base.items()}

_CASH1_COORDS = {
    "cs_lubes_ltrs":        ( 310,  600,  700,  680),
    "cs_lubes_ugx":         ( 955,  600, 1870,  680),
    "cs_tba_ugx":           ( 310,  708,  960,  798),
    "cs_plus_card_credits": ( 955,  708, 1870,  798),
    "cs_lpg_kgs":           ( 310,  820,  700,  910),
    "cs_lpg_ugx":           ( 955,  820, 1870,  910),
    "cs_shop_sales":        ( 310,  935,  960, 1025),
    "cs_other_payment":     ( 955,  935, 1870, 1025),
    "cs_tyre_ugx":          ( 310, 1060,  960, 1150),
    "cs_plus_pms":          ( 310, 1230,  960, 1315),
    "cs_plus_ago":          ( 955, 1230, 1870, 1315),
    "cs_plus_others":       ( 310, 1345,  960, 1435),
    "cs_momo":              ( 955, 1345, 1870, 1435),
    "cs_airtel":            ( 310, 1470,  960, 1555),
    "cs_visa":              ( 955, 1470, 1870, 1555),
    "cs_debtors_credit":    ( 310, 1590,  960, 1680),
}

_CASH2_COORDS = {
    "cs_exp_umeme":          ( 310,  745,  960,  835),
    "cs_exp_water":          ( 955,  772, 1870,  862),
    "cs_exp_security":       ( 310,  862,  960,  972),
    "cs_exp_stationery":     ( 955,  896, 1870,  986),
    "cs_exp_generator":      ( 310,  999,  960, 1089),
    "cs_exp_meals":          ( 955, 1019, 1870, 1109),
    "cs_exp_transport":      ( 310, 1135,  960, 1205),
    "cs_exp_salaries":       ( 955, 1143, 1870, 1233),
    "cs_exp_sanitary":       ( 310, 1250,  960, 1340),
    "cs_exp_airtime":        ( 955, 1266, 1870, 1356),
    "cs_exp_misc":           ( 310, 1374,  960, 1464),
    "cs_exp_shop_packaging": ( 955, 1389, 1870, 1479),
    "cs_stock_tba":          ( 310, 1554,  960, 1644),
    "cs_stock_lpg_acc":      ( 955, 1566, 1870, 1656),
    "cs_stock_shop":         ( 310, 1678,  960, 1768),
    "cs_cash_banked":        ( 290, 1855,  960, 1945),
}

_PRODUCTS_COORDS = {
    "pt_lubricants": ( 310,  515,  960,  625),
    "pt_tba":        ( 955,  515, 1870,  625),
    "pt_lpg_acc":    ( 310,  650,  960,  760),
    "pt_lpg":        ( 955,  650, 1870,  760),
    "pt_car_wash":   ( 310,  780,  960,  890),
    "pt_shop":       ( 955,  780, 1870,  890),
    "pt_solar":      ( 310,  905,  960, 1015),
}

_EXPENSES_COORDS = {
    "dx_meals":          ( 310,  695,  960,  775),
    "dx_generator":      ( 955,  695, 1870,  775),  # confirmed 1800 on 2026-06-17
    "dx_electricity":    ( 310,  805,  960,  885),
    "dx_water":          ( 955,  805, 1870,  885),
    "dx_salaries":       ( 310,  915,  960,  995),
    "dx_stationery":     ( 955,  915, 1870,  995),
    "dx_security":       ( 310, 1025,  960, 1105),
    "dx_sanitation":     ( 955, 1025, 1870, 1105),
    "dx_airtime":        ( 310, 1135,  960, 1215),
    "dx_transport":      ( 955, 1135, 1870, 1215),
    "dx_nssf":           ( 310, 1245,  960, 1325),
    "dx_sundries":       ( 955, 1245, 1870, 1325),  # confirmed ~500 on 2026-06-17
    "dx_maintenance":    ( 310, 1355,  960, 1435),
    "dx_vat":            ( 955, 1355, 1870, 1435),
    "dx_photocopy":      ( 310, 1495,  960, 1575),
    "dx_tax_compliance": ( 955, 1495, 1870, 1575),
}


# =============================================================
# BOX-COUNT MAP  (from capture_template_builder.py groups)
# =============================================================
# Every printed field is a fixed grid of digit boxes:
#   litre/kg fields  → [3,3]   = 6 boxes
#   UGX money fields → [3,3,2] = 8 boxes
# Knowing the exact box count lets us force GPT to transcribe each box
# individually (digit or blank) instead of guessing "the number" — this
# eliminates dropped/duplicated digits (e.g. 620000 misread as 6200000).

# Fields that are 6-box (litre / kg) rather than the 8-box money default.
_SIX_BOX_FIELDS = {"cs_lubes_ltrs", "cs_lpg_kgs"}

def _box_count(fid: str) -> int:
    # Meter book: dips & return-to-tank are litre (6); prices are money (8).
    if any(s in fid for s in ("opening_dip", "closing_dip", "return_tank")):
        return 6
    if "cost_price" in fid or "selling_price" in fid:
        return 8
    if fid in _SIX_BOX_FIELDS:
        return 6
    # All cash / product / expense money fields are 8-box.
    return 8


# =============================================================
# PUBLIC ENTRY POINT
# =============================================================

def read_capture_sheets(photo_paths: list, entry_date: str, shift: str,
                        station_name: str = None, machine_id: str = None) -> dict:
    # Store station credentials so _call_vision_api can route through the
    # StationDeck server (server holds the OpenAI key).
    global _OCR_STATION_NAME, _OCR_MACHINE_ID
    _OCR_STATION_NAME = station_name
    _OCR_MACHINE_ID   = machine_id

    if not OCR_AVAILABLE:
        return _all_unread("pytesseract or Pillow is not installed on this machine.")
    if not VISION_AVAILABLE:
        return _all_unread("OCR backend unavailable — cannot read digits.")

    def _is_meaningful(v):
        """True when v carries a confirmed non-zero digit reading."""
        if isinstance(v, dict):
            val = v.get("value", "")
            return val != "" and val != "0"
        return v != 0

    results = {}
    for path in photo_paths:
        try:
            sheet_type = _detect_sheet_type(path)
            logger.info(f"Detected sheet type '{sheet_type}' for {path.name}")

            if sheet_type in ("pms", "ago"):
                # Process one tank at a time so each API tile has only ~7 fields.
                # All 4 tanks at once (~20 fields) makes tiles too small for GPT
                # to read reliably.
                sheet_results = {}
                for tank in range(1, 5):
                    tank_coords = _meter_book_coords_tank(sheet_type, tank)
                    tank_results = _read_sheet_with_vision(
                        path, tank_coords, f"{sheet_type} tank {tank}")
                    sheet_results.update(tank_results)
            elif sheet_type == "cash1":
                sheet_results = _read_sheet_with_vision(
                    path, _CASH1_COORDS, "cash1")
            elif sheet_type == "cash2":
                sheet_results = _read_sheet_with_vision(
                    path, _CASH2_COORDS, "cash2")
            elif sheet_type == "products":
                sheet_results = _read_sheet_with_vision(
                    path, _PRODUCTS_COORDS, "products")
            elif sheet_type == "expenses":
                sheet_results = _read_sheet_with_vision(
                    path, _EXPENSES_COORDS, "expenses")
            else:
                logger.warning(f"Unknown sheet type for {path.name} — skipping.")
                continue

            # Merge: a confirmed non-zero reading always overwrites any previous
            # value for that key. A blank/zero result only fills in keys that
            # haven't been seen yet — it never erases a confirmed reading.
            for k, v in sheet_results.items():
                if _is_meaningful(v):
                    results[k] = v
                elif k not in results:
                    results[k] = v

        except Exception as e:
            logger.error(f"Error processing {path}: {e}", exc_info=True)
            continue

    return results


# =============================================================
# SHEET TYPE DETECTION
# =============================================================

def _classify_text(text: str) -> str:
    """
    Given lowercase OCR text from a crop, return the sheet type string
    or empty string if no match found.

    Keyword strategy (v6):
    - meter book sheets: require BOTH "meter book" AND product name
    - cash2: require "cash"+"sales" AND any banking/expense sub-keyword
    - cash1: require "cash"+"sales" (catches page 1 of cash sheet)
    - products: "product"+"sales" checked independently (handles partial title
                where "Station" is hidden by book spine), OR "lubricants" alone
                (visible field label picked up when title is partially obscured)
    - expenses: "daily"+"expenses" first; fallback to "expenses" alone only
                after ruling out cash sheets (avoids false match on cash2
                which mentions "expenses this shift")
    """
    # Meter book — most specific, check first.
    # Both PMS and AGO sheets have instruction text "PMS/AGO fuel sales are
    # calculated from the Meter Book", so bare "pms"/"ago" + "meter book"
    # is not enough. Check for "(pms)"/"(ago)" from the title first, then
    # fall back to bare keywords after stripping the instruction phrase.
    if "meter book" in text:
        if "(ago)" in text:
            return "ago"
        if "(pms)" in text:
            return "pms"
        # OCR may drop parentheses — strip cross-contaminating phrases and retry
        stripped = (text.replace("pms/ago", "")
                        .replace("pms / ago", "")
                        .replace("as the pms sheet", "")
                        .replace("pms sheet", ""))
        if "ago" in stripped:
            return "ago"
        if "pms" in stripped:
            return "pms"
        # meter book present but product undetermined — fall through to other checks

    # Cash sheets — check before expenses to avoid false "expenses" match
    if "cash" in text and "sales" in text:
        # Use page number from title first — most reliable signal
        if "2 of 2" in text or "2of2" in text:
            return "cash2"
        if "1 of 2" in text or "1of2" in text:
            return "cash1"
        # Fallback: banking/expense section keywords unique to page 2
        if any(k in text for k in ["banking", "actual cash", "expenses this shift",
                                   "expenses (this shift)"]):
            return "cash2"
        return "cash1"

    # Product Sales Totals — words checked independently
    # Also catches "lubricants" as a field label unique to this sheet
    if ("product" in text and "sales" in text) or "lubricants" in text:
        return "products"

    # Daily Expenses — "daily expenses" exact first, then "expenses" alone
    # (safe here because cash sheets already handled above)
    if "daily" in text and "expenses" in text:
        return "expenses"
    if "expenses" in text:
        return "expenses"

    return ""


def _detect_sheet_type(path: Path) -> str:
    try:
        img  = Image.open(str(path))
        w, h = img.size

        # ── Pass 1: top third (fast, works for well-framed photos) ──
        top_third = img.crop((0, 0, w, h // 3)).convert("L")
        text = pytesseract.image_to_string(top_third, config="--psm 3").lower()
        logger.debug(f"Pass-1 OCR text for {path.name!r}: {text[:200]!r}")

        result = _classify_text(text)
        if result:
            return result

        # ── Pass 2: top 60% (handles angled shots where title sits lower) ──
        top_60 = img.crop((0, 0, w, int(h * 0.60))).convert("L")
        text2 = pytesseract.image_to_string(top_60, config="--psm 3").lower()
        logger.debug(f"Pass-2 OCR text for {path.name!r}: {text2[:200]!r}")

        result = _classify_text(text2)
        if result:
            return result

        # ── Filename hint fallback (last resort) ──
        name = path.name.lower()
        for hint, stype in [("pms", "pms"), ("ago", "ago"), ("cash1", "cash1"),
                             ("cash2", "cash2"), ("product", "products"),
                             ("expense", "expenses")]:
            if hint in name:
                return stype

        return "unknown"

    except Exception as e:
        logger.warning(f"Sheet detection failed for {path}: {e}")
        return "unknown"


# =============================================================
# VISION-BASED DIGIT READING
# =============================================================

def _enhance_for_ocr(crop: Image.Image) -> Image.Image:
    """
    Pre-process a crop so handwritten digit strokes are crisp for GPT vision.
    Steps: grayscale → strong contrast boost → two sharpening passes.
    This makes ink strokes dark and distinct, reducing 5-vs-3, 1-vs-7 errors.
    """
    g = crop.convert("L")                          # remove colour noise
    g = ImageEnhance.Contrast(g).enhance(2.5)      # darken ink, whiten paper
    g = g.filter(ImageFilter.SHARPEN)              # crisp edges (pass 1)
    g = g.filter(ImageFilter.SHARPEN)              # crisp edges (pass 2)
    return g.convert("RGB")


def _read_sheet_with_vision(path: Path, coord_map: dict, sheet_label: str) -> dict:
    results = {}
    try:
        img = Image.open(str(path))
        img_w, img_h = img.size
        sx = img_w / REF_W
        sy = img_h / REF_H

        # UPDF fields are always blank at civilian stations — hardcode 0 and skip
        # them from the tiled image to prevent cross-contamination in GPT responses
        zero_result = {"value": "0", "status": "ok", "note": ""}
        hardcoded_zeros = {}
        active_coord_map = {}
        for fid in coord_map:
            if "updf" in fid:
                hardcoded_zeros[fid] = zero_result
            else:
                active_coord_map[fid] = coord_map[fid]

        crops = {}
        for fid, (x1, y1, x2, y2) in active_coord_map.items():
            cx1 = max(0, min(int(x1 * sx), img_w))
            cy1 = max(0, min(int(y1 * sy), img_h))
            cx2 = max(0, min(int(x2 * sx), img_w))
            cy2 = max(0, min(int(y2 * sy), img_h))
            if cx2 > cx1 and cy2 > cy1:
                crops[fid] = img.crop((cx1, cy1, cx2, cy2))

        if not crops:
            return {fid: _unread("No valid field regions found.") for fid in coord_map}

        TILE_W   = 700   # wider → ~117px per digit box, easier for GPT to read
        CROP_H   = 110  # taller → more vertical resolution per crop
        LABEL_H  = 22
        TILE_GAP = 6

        fid_list = list(crops.keys())
        tile_h   = (CROP_H + LABEL_H + TILE_GAP) * len(fid_list)
        tiled    = Image.new("RGB", (TILE_W, tile_h), (255, 255, 255))
        draw     = ImageDraw.Draw(tiled)

        y_cursor = 0
        for fid in fid_list:
            crop   = crops[fid].convert("RGB")
            crop_e = _enhance_for_ocr(crop)                        # contrast+sharpen
            crop_r = crop_e.resize((TILE_W, CROP_H), Image.LANCZOS)
            draw.text((4, y_cursor + 4), fid, fill=(180, 0, 0))
            y_cursor += LABEL_H
            tiled.paste(crop_r, (0, y_cursor))
            y_cursor += CROP_H + TILE_GAP

        buf = io.BytesIO()
        tiled.save(buf, format="PNG")   # lossless — no JPEG blur on digit edges
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

        field_list_text = "\n".join(f"- {fid}" for fid in fid_list)
        prompt = (
            f"This image shows cropped fields from a fuel station daily entry form "
            f"({sheet_label} sheet). Each crop is labelled in red with its field name.\n\n"
            f"YOUR TASK: For each field, read ONLY the handwritten digits inside the "
            f"input boxes.\n\n"
            f"STRICT RULES:\n"
            f"- Return 0 if the boxes are empty (no handwriting present)\n"
            f"- Return 0 if you see only printed form labels, field names, or "
            f"instructions — these are NOT digit values\n"
            f"- Return 0 if the crop shows a blank margin or background\n"
            f"- Only return a non-zero number when you can CLEARLY see handwritten "
            f"ink digits written inside the boxes\n"
            f"- Empty printed boxes are NOT zeros — a row of blank boxes is 0, not "
            f"a string of zeros. Do NOT pad a number with extra trailing or leading "
            f"zeros to fill empty boxes.\n"
            f"- Digits only — no commas, no spaces, no letters\n\n"
            f"COMPLETENESS RULES — read these carefully:\n"
            f"- Read exactly the digits that are handwritten — no more, no fewer.\n"
            f"- The leftmost handwritten digit may appear at the very left edge of the "
            f"image crop — it is still a real digit, include it.\n"
            f"- Do NOT add digits for empty boxes. If 6 boxes are filled and 2 are "
            f"blank, the answer has 6 digits, not 8.\n"
            f"- Cost prices and selling prices (fields ending in cost_price / selling_price) "
            f"are 4-digit numbers in the range 4000–9000 UGX/L. If you read only 3 digits "
            f"below 1000, the leftmost digit was likely cut at the image edge — look again.\n"
            f"- Digits are written one per box; read every FILLED box from left to right.\n\n"
            f"DIGIT DISAMBIGUATION — easily confused pairs:\n"
            f"- 5 vs 3: a 5 has a FLAT horizontal top stroke then a belly; a 3 has two "
            f"open curves and NO flat top. Flat bar on top = 5.\n"
            f"- 1 vs 7: a 7 has a diagonal top stroke; a 1 is near-vertical.\n"
            f"- 0 vs 6 vs 8: 0 is a plain oval; 6 has one loop at the bottom; 8 has two loops.\n"
            f"- 9 vs 4: 9 has a closed top loop with a tail; 4 has an open top angle.\n"
            f"Take your time — accuracy matters more than speed.\n\n"
            f"Fields:\n{field_list_text}\n\n"
            f"Reply ONLY with a JSON object. No markdown, no backticks, no extra text.\n"
            f'Example: {{"pms1_opening_dip": "315", "pms1_cost_price": "6467"}}'
        )

        raw_response = _call_vision_api(b64, prompt)
        if raw_response is None:
            return {
                fid: _unread("Vision API call failed — please enter manually.")
                for fid in coord_map
            }

        try:
            clean = raw_response.strip().strip("`").strip()
            if clean.lower().startswith("json"):
                clean = clean[4:].strip()
            api_values = json.loads(clean)
        except json.JSONDecodeError as e:
            logger.error(f"JSON parse failed: {e} | raw: {raw_response[:300]}")
            return {
                fid: _unread("Could not parse vision API response.")
                for fid in coord_map
            }

        for fid in active_coord_map:
            digits = "".join(ch for ch in str(api_values.get(fid, "")) if ch.isdigit())
            n_boxes = _box_count(fid)

            if digits == "" or digits == "0":
                results[fid] = _unread("Field appears blank — please enter manually.")
            elif len(digits) > n_boxes:
                # More digits than the field physically has boxes → misread.
                results[fid] = {
                    "value": digits, "status": "warn",
                    "note": (f"Read {len(digits)} digits but this field only has "
                             f"{n_boxes} boxes — please verify against the sheet."),
                }
            else:
                status, note = _validate(fid, digits)
                results[fid] = {"value": digits, "status": status, "note": note}

        results.update(hardcoded_zeros)

    except Exception as e:
        logger.error(f"_read_sheet_with_vision failed for {path}: {e}", exc_info=True)
        return {fid: _unread("Unexpected error — please enter manually.") for fid in coord_map}

    return results


# =============================================================
# VISION API CALL  —  server-routed, with local fallback
# =============================================================

def _call_vision_api(image_b64: str, prompt: str):
    # ── Primary: route through the StationDeck server (key stays server-side) ──
    if _OCR_STATION_NAME and _OCR_MACHINE_ID and _REQUESTS_OK:
        try:
            resp = requests.post(
                OCR_SERVER_URL,
                json={
                    "station_name": _OCR_STATION_NAME,
                    "machine_id":   _OCR_MACHINE_ID,
                    "image_b64":    image_b64,
                    "prompt":       prompt,
                },
                timeout=60,
            )
            data = resp.json()
            if data.get("success"):
                return data.get("text")
            logger.error(f"Server OCR failed ({resp.status_code}): {data.get('message')}")
            return None
        except Exception as e:
            logger.error(f"Server OCR request failed: {e}")
            return None

    # ── Fallback: local OpenAI key (dev / test_ocr.py only) ──
    if _openai_client is None:
        logger.error("No OCR backend: no station credentials and no local OpenAI key.")
        return None
    try:
        response = _openai_client.chat.completions.create(
            model=OPENAI_MODEL,
            max_tokens=800,
            temperature=0,   # deterministic — same photo always reads the same
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{image_b64}",
                            "detail": "high",
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }],
        )
        return response.choices[0].message.content
    except Exception as e:
        logger.error(f"OpenAI vision API call failed: {e}")
        return None


# =============================================================
# VALIDATION
# =============================================================

_TYPICAL_MAX = {
    "cs_lubes_ltrs": 50_000, "cs_lpg_kgs": 5_000,
    "cs_lubes_ugx": 80_000_000, "cs_tba_ugx": 10_000_000,
    "cs_lpg_ugx": 5_000_000, "cs_plus_card_credits": 80_000_000,
    "cs_shop_sales": 5_000_000, "cs_other_payment": 30_000_000,
    "cs_tyre_ugx": 2_000_000, "cs_plus_pms": 30_000_000,
    "cs_plus_ago": 50_000_000, "cs_plus_others": 10_000_000,
    "cs_momo": 5_000_000, "cs_airtel": 5_000_000,
    "cs_visa": 20_000_000, "cs_debtors_credit": 60_000_000,
    "cs_exp_umeme": 3_000_000, "cs_exp_water": 1_000_000,
    "cs_exp_security": 2_000_000, "cs_exp_stationery": 500_000,
    "cs_exp_generator": 2_000_000, "cs_exp_meals": 2_000_000,
    "cs_exp_transport": 1_000_000, "cs_exp_salaries": 10_000_000,
    "cs_exp_sanitary": 500_000, "cs_exp_airtime": 300_000,
    "cs_exp_misc": 1_000_000, "cs_exp_shop_packaging": 500_000,
    "cs_stock_tba": 5_000_000, "cs_stock_lpg_acc": 2_000_000,
    "cs_stock_shop": 5_000_000, "cs_cash_banked": 100_000_000,
    "pt_lubricants": 80_000_000, "pt_tba": 10_000_000,
    "pt_lpg_acc": 3_000_000, "pt_lpg": 30_000_000,
    "pt_car_wash": 2_000_000, "pt_shop": 5_000_000,
    "pt_solar": 3_000_000,
    "dx_meals": 2_000_000, "dx_generator": 2_000_000,
    "dx_electricity": 3_000_000, "dx_water": 1_000_000,
    "dx_salaries": 10_000_000, "dx_stationery": 500_000,
    "dx_security": 2_000_000, "dx_sanitation": 500_000,
    "dx_airtime": 300_000, "dx_transport": 1_000_000,
    "dx_nssf": 500_000, "dx_sundries": 1_000_000,
    "dx_maintenance": 2_000_000, "dx_vat": 5_000_000,
    "dx_photocopy": 200_000, "dx_tax_compliance": 1_000_000,
}

def _validate(field_id: str, value: str) -> tuple:
    try:
        v = float(value)
    except (ValueError, TypeError):
        return "unread", "Could not parse value — please enter manually."
    if v < 0:
        return "warn", "Negative value — please verify."
    if "dip" in field_id or "return_tank" in field_id or "updf" in field_id:
        if v > 25_000:
            return "warn", f"Dip reading {v:,.0f} L seems very high — please verify."
        return "ok", ""
    if "cost_price" in field_id or "selling_price" in field_id:
        if 100 <= v < 1_000:
            return "warn", (
                f"Price read as {v:,.0f} UGX/L — likely missing a leading digit "
                f"(expected 4000–9000). Check the original sheet and type the correct value."
            )
        if v < 3_000 or v > 12_000:
            return "warn", f"Price {v:,.0f} UGX/L outside expected range — please verify."
        return "ok", ""
    if field_id in _TYPICAL_MAX:
        cap = _TYPICAL_MAX[field_id]
        if v > cap * 10:
            return "warn", (
                f"Value {v:,.0f} UGX is far above the typical maximum "
                f"({cap:,.0f}) — likely an extra digit. Please verify."
            )
        if v > cap:
            return "warn", (
                f"Value {v:,.0f} UGX is above the typical range "
                f"({cap:,.0f}) — please double-check against the sheet."
            )
    return "ok", ""


# =============================================================
# HELPERS
# =============================================================

def _unread(note: str) -> dict:
    return {"value": "", "status": "unread", "note": note}

def _all_unread(reason: str) -> dict:
    all_fields = []
    for product in ("pms", "ago"):
        for tank in range(1, 5):
            for f in ("opening_dip", "closing_dip", "return_tank",
                      "updf_receipt", "updf_consumption",
                      "cost_price", "selling_price"):
                all_fields.append(f"{product}{tank}_{f}")
    all_fields += list(_CASH1_COORDS.keys())
    all_fields += list(_CASH2_COORDS.keys())
    all_fields += list(_PRODUCTS_COORDS.keys())
    all_fields += list(_EXPENSES_COORDS.keys())
    return {fid: _unread(reason) for fid in all_fields}