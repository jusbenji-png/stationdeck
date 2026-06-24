# =============================================================
# src/ocr_reader.py
# StationDeck — OCR Capture Sheet Reader  (v6)
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
    from PIL import Image, ImageDraw

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

# ── OpenAI client ─────────────────────────────────────────────
try:
    from openai import OpenAI
    from config.settings import OPENAI_API_KEY, OPENAI_MODEL
    _openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
    VISION_AVAILABLE = _openai_client is not None
except Exception as e:
    _openai_client = None
    VISION_AVAILABLE = False
    logger.warning(f"OpenAI client not available: {e}")


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
    "return_tank":      (1220, 565, 1900, 657),
    # UPDF row — shifted +30px down from v4
    "updf_receipt":     ( 280, 742,  700, 832),
    "updf_consumption": ( 750, 742, 1170, 832),
    # Price row — shifted +28px down, selling_price x extended to full width
    "cost_price":       ( 280, 800,  700, 900),
    "selling_price":    ( 600, 800, 1920, 900),
}

# AGO sheet has the same internal layout as PMS
_AGO_TANK1 = {
    "opening_dip":      ( 280, 565,  700, 657),
    "closing_dip":      ( 750, 565, 1170, 657),
    "return_tank":      (1220, 565, 1900, 657),
    "updf_receipt":     ( 280, 742,  700, 832),
    "updf_consumption": ( 750, 742, 1170, 832),
    "cost_price":       ( 280, 800,  700, 900),
    "selling_price":    ( 600, 800, 1920, 900),
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

_CASH1_COORDS = {
    "cs_lubes_ltrs":        ( 310,  600,  700,  680),
    "cs_lubes_ugx":         ( 955,  600, 1870,  680),
    "cs_tba_ugx":           ( 310,  708,  875,  798),
    "cs_plus_card_credits": ( 955,  708, 1870,  798),
    "cs_lpg_kgs":           ( 310,  820,  700,  910),
    "cs_lpg_ugx":           ( 955,  820, 1870,  910),
    "cs_shop_sales":        ( 310,  935,  875, 1025),
    "cs_other_payment":     ( 955,  935, 1870, 1025),
    "cs_tyre_ugx":          ( 310, 1060,  875, 1150),
    "cs_plus_pms":          ( 310, 1230,  875, 1315),
    "cs_plus_ago":          ( 955, 1230, 1870, 1315),
    "cs_plus_others":       ( 310, 1345,  875, 1435),
    "cs_momo":              ( 955, 1345, 1870, 1435),
    "cs_airtel":            ( 310, 1470,  875, 1555),
    "cs_visa":              ( 955, 1470, 1870, 1555),
    "cs_debtors_credit":    ( 310, 1590,  875, 1680),
}

_CASH2_COORDS = {
    "cs_exp_umeme":          ( 310,  745,  875,  835),
    "cs_exp_water":          ( 955,  772, 1870,  862),
    "cs_exp_security":       ( 310,  872,  875,  962),
    "cs_exp_stationery":     ( 955,  896, 1870,  986),
    "cs_exp_generator":      ( 310,  999,  875, 1089),
    "cs_exp_meals":          ( 955, 1019, 1870, 1109),
    "cs_exp_transport":      ( 310, 1125,  875, 1215),
    "cs_exp_salaries":       ( 955, 1143, 1870, 1233),
    "cs_exp_sanitary":       ( 310, 1250,  875, 1340),
    "cs_exp_airtime":        ( 955, 1266, 1870, 1356),
    "cs_exp_misc":           ( 310, 1374,  875, 1464),
    "cs_exp_shop_packaging": ( 955, 1389, 1870, 1479),
    "cs_stock_tba":          ( 310, 1554,  875, 1644),
    "cs_stock_lpg_acc":      ( 955, 1566, 1870, 1656),
    "cs_stock_shop":         ( 310, 1678,  875, 1768),
    "cs_cash_banked":        ( 290, 1855,  890, 1945),
}

_PRODUCTS_COORDS = {
    "pt_lubricants": ( 310,  630,  875,  720),
    "pt_tba":        ( 955,  630, 1870,  720),
    "pt_lpg_acc":    ( 310,  765,  875,  855),
    "pt_lpg":        ( 955,  765, 1870,  855),
    "pt_car_wash":   ( 310,  895,  875,  985),
    "pt_shop":       ( 955,  895, 1870,  985),
    "pt_solar":      ( 310, 1020,  875, 1110),
}

_EXPENSES_COORDS = {
    "dx_meals":          ( 310,  465,  875,  555),
    "dx_generator":      ( 955,  475, 1870,  565),
    "dx_electricity":    ( 310,  575,  875,  665),
    "dx_water":          ( 955,  585, 1870,  675),
    "dx_salaries":       ( 310,  685,  875,  775),
    "dx_stationery":     ( 955,  695, 1870,  785),
    "dx_security":       ( 310,  795,  875,  885),
    "dx_sanitation":     ( 955,  805, 1870,  895),
    "dx_airtime":        ( 310,  905,  875,  995),
    "dx_transport":      ( 955,  915, 1870, 1005),
    "dx_nssf":           ( 310, 1015,  875, 1105),
    "dx_sundries":       ( 955, 1025, 1870, 1115),
    "dx_maintenance":    ( 310, 1130,  875, 1220),
    "dx_vat":            ( 955, 1135, 1870, 1225),
    "dx_photocopy":      ( 310, 1250,  875, 1340),
    "dx_tax_compliance": ( 955, 1245, 1870, 1335),
}


# =============================================================
# PUBLIC ENTRY POINT
# =============================================================

def read_capture_sheets(photo_paths: list, entry_date: str, shift: str) -> dict:
    if not OCR_AVAILABLE:
        return _all_unread("pytesseract or Pillow is not installed on this machine.")
    if not VISION_AVAILABLE:
        return _all_unread("OpenAI API key not configured — cannot read digits.")

    results = {}
    for path in photo_paths:
        try:
            sheet_type = _detect_sheet_type(path)
            logger.info(f"Detected sheet type '{sheet_type}' for {path.name}")

            if sheet_type == "pms":
                sheet_results = _read_sheet_with_vision(
                    path, _meter_book_coords("pms"), "pms")
            elif sheet_type == "ago":
                sheet_results = _read_sheet_with_vision(
                    path, _meter_book_coords("ago"), "ago")
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

            results.update(sheet_results)

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
    # Meter book — most specific, check first
    if "meter book" in text and "pms" in text:
        return "pms"
    if "meter book" in text and "ago" in text:
        return "ago"

    # Cash sheets — check before expenses to avoid false "expenses" match
    if "cash" in text and "sales" in text:
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

def _read_sheet_with_vision(path: Path, coord_map: dict, sheet_label: str) -> dict:
    results = {}
    try:
        img = Image.open(str(path))
        img_w, img_h = img.size
        sx = img_w / REF_W
        sy = img_h / REF_H

        crops = {}
        for fid, (x1, y1, x2, y2) in coord_map.items():
            cx1 = max(0, min(int(x1 * sx), img_w))
            cy1 = max(0, min(int(y1 * sy), img_h))
            cx2 = max(0, min(int(x2 * sx), img_w))
            cy2 = max(0, min(int(y2 * sy), img_h))
            if cx2 > cx1 and cy2 > cy1:
                crops[fid] = img.crop((cx1, cy1, cx2, cy2))

        if not crops:
            return {fid: _unread("No valid field regions found.") for fid in coord_map}

        TILE_W   = 500
        CROP_H   = 90
        LABEL_H  = 22
        TILE_GAP = 4

        fid_list = list(crops.keys())
        tile_h   = (CROP_H + LABEL_H + TILE_GAP) * len(fid_list)
        tiled    = Image.new("RGB", (TILE_W, tile_h), (255, 255, 255))
        draw     = ImageDraw.Draw(tiled)

        y_cursor = 0
        for fid in fid_list:
            crop   = crops[fid].convert("RGB")
            crop_r = crop.resize((TILE_W, CROP_H), Image.LANCZOS)
            draw.text((4, y_cursor + 4), fid, fill=(180, 0, 0))
            y_cursor += LABEL_H
            tiled.paste(crop_r, (0, y_cursor))
            y_cursor += CROP_H + TILE_GAP

        buf = io.BytesIO()
        tiled.save(buf, format="JPEG", quality=92)
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

        field_list_text = "\n".join(f"- {fid}" for fid in fid_list)
        prompt = (
            f"This image shows cropped handwritten digit fields from a fuel station "
            f"daily entry form ({sheet_label} sheet). Each crop is labelled in red "
            f"with its field name above it.\n\n"
            f"For each field listed below, read the handwritten digits visible in the "
            f"boxes in that crop. If a crop is blank (no digits written), return 0. "
            f"Digits only — no commas, no spaces, no letters.\n\n"
            f"Fields:\n{field_list_text}\n\n"
            f"Reply ONLY with a JSON object. No markdown, no backticks, no extra text.\n"
            f'Example: {{"pms1_opening_dip": "315", "pms1_closing_dip": "84"}}'
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

        for fid in coord_map:
            raw_val = str(api_values.get(fid, "")).strip()
            if raw_val == "" or raw_val == "0":
                results[fid] = _unread("Field appears blank — please enter manually.")
            else:
                status, note = _validate(fid, raw_val)
                results[fid] = {"value": raw_val, "status": status, "note": note}

    except Exception as e:
        logger.error(f"_read_sheet_with_vision failed for {path}: {e}", exc_info=True)
        return {fid: _unread("Unexpected error — please enter manually.") for fid in coord_map}

    return results


# =============================================================
# OPENAI VISION API CALL
# =============================================================

def _call_vision_api(image_b64: str, prompt: str):
    if _openai_client is None:
        logger.error("OpenAI client not initialised.")
        return None
    try:
        response = _openai_client.chat.completions.create(
            model=OPENAI_MODEL,
            max_tokens=800,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{image_b64}",
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
        if v < 3_000 or v > 12_000:
            return "warn", f"Price {v:,.0f} UGX/L outside expected range — please verify."
        return "ok", ""
    if field_id in _TYPICAL_MAX and v > _TYPICAL_MAX[field_id] * 10:
        return "warn", f"Value {v:,.0f} UGX far above typical range — possible extra digit."
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