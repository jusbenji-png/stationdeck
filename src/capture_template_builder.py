# =============================================================
# src/capture_template_builder.py
# StationDeck — Capture Template PDF Builder
#
# Generates a print-ready 5-sheet PDF capture template.
# All 5 sheets in one PDF for single-click printing.
#
# Layout rules (derived from A4 math):
#   Content width = 186mm (210 - 12 - 12 margins)
#   2-column layout = 93mm per column  ← only layout that fits 8 boxes
#   3-column layout = 62mm per column  ← fits 6-digit litre boxes only
#   8-digit UGX box row at 7.5mm/box = 71mm → fits in 93mm col ✓
#   6-digit litre box row at 7.5mm/box = 54mm → fits in 62mm col ✓
# =============================================================

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from io import BytesIO

# ── Colours ───────────────────────────────────────────────────
DARK       = colors.HexColor("#1a1a1a")
HEADER_BG  = colors.HexColor("#1f4d3d")
SECTION_BG = colors.HexColor("#1a1a1a")
MID_GREY   = colors.HexColor("#cccccc")
LIGHT_GREY = colors.HexColor("#f5f5f5")
WHITE      = colors.white

# ── Page layout ───────────────────────────────────────────────
PAGE_W, PAGE_H = A4          # 210 × 297mm
MARGIN_L  = 12 * mm
MARGIN_R  = 12 * mm
MARGIN_T  = 12 * mm
MARGIN_B  = 10 * mm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R   # 186mm

# ── Box dimensions ────────────────────────────────────────────
# These are sized so 8 boxes fit in a 2-column (93mm) layout
DIGIT_W     = 7.5 * mm    # width of one digit box
DIGIT_H     = 10.0 * mm   # height — comfortable for handwriting
GAP         = 1.0 * mm    # gap between adjacent boxes
GROUP_GAP   = 2.5 * mm    # gap between digit groups (e.g. 3|3|2)

# ── Typography ────────────────────────────────────────────────
FS_TITLE    = 11
FS_SUB      = 8
FS_META     = 7.5
FS_SECTION  = 8
FS_LABEL    = 7.5
FS_UNIT     = 6.5
FS_INSTRUCT = 6.5
FS_FOOTER   = 6.5


# =============================================================
# PUBLIC ENTRY POINT
# =============================================================

def build_capture_template_pdf(
    buffer: BytesIO,
    station_name: str,
    location: str,
    station_id: str,
) -> None:
    c = canvas.Canvas(buffer, pagesize=A4)
    sub = f"{station_name} \u00b7 {location}" if location else station_name

    _sheet_meter_book(c, "PMS", sub, station_id,
        footer="Page 1 of 2 \u2014 AGO tanks continue on the next sheet.")
    c.showPage()

    _sheet_meter_book(c, "AGO", sub, station_id,
        footer="Page 2 of 2 \u2014 Meter Book complete. Continue to Cash & Sales Capture.")
    c.showPage()

    _sheet_cash_sales(c, sub, station_id)
    c.showPage()

    _sheet_product_totals(c, sub, station_id)
    c.showPage()

    _sheet_daily_expenses(c, sub, station_id)

    c.save()


# =============================================================
# SHARED DRAWING HELPERS
# =============================================================

def _page_header(c, title, subtitle, station_id, shift_box=True):
    """Draw page header. Returns Y coordinate below the divider line."""
    y = PAGE_H - MARGIN_T

    c.setFont("Helvetica-Bold", FS_TITLE)
    c.setFillColor(DARK)
    c.drawString(MARGIN_L, y, f"StationDeck \u2014 {title}")

    c.setFont("Helvetica", FS_SUB)
    c.setFillColor(colors.HexColor("#555"))
    c.drawString(MARGIN_L, y - 5*mm, subtitle)

    rx = PAGE_W - MARGIN_R
    c.setFont("Helvetica", FS_META)
    c.setFillColor(DARK)
    c.drawRightString(rx, y, "Date: _______________________")
    if shift_box:
        c.drawRightString(rx, y - 4.5*mm, "Shift: \u2610 Day  \u2610 Night")
    else:
        c.drawRightString(rx, y - 4.5*mm, "(One entry per day \u2014 not per shift)")
    c.drawRightString(rx, y - 9*mm, f"Station ID: {station_id}")

    divider_y = y - 13*mm
    c.setStrokeColor(DARK)
    c.setLineWidth(1.5)
    c.line(MARGIN_L, divider_y, PAGE_W - MARGIN_R, divider_y)

    return divider_y - 3*mm


def _instruction_box(c, text, y):
    """Draw grey instruction box. Returns Y below the box."""
    bh = 9*mm
    c.setFillColor(colors.HexColor("#f0f0f0"))
    c.setStrokeColor(MID_GREY)
    c.setLineWidth(0.5)
    c.rect(MARGIN_L, y - bh, CONTENT_W, bh, fill=1, stroke=1)
    c.setFillColor(DARK)
    c.setFont("Helvetica", FS_INSTRUCT)
    c.drawString(MARGIN_L + 2*mm, y - 6*mm, text[:200])
    return y - bh - 2*mm


def _section_bar(c, label, y, bg=None):
    """Dark section header bar. Returns Y below the bar."""
    bh = 6*mm
    c.setFillColor(bg or SECTION_BG)
    c.rect(MARGIN_L, y - bh, CONTENT_W, bh, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", FS_SECTION)
    c.drawString(MARGIN_L + 2*mm, y - 4.2*mm, label)
    return y - bh


def _box_row_width(groups):
    """Calculate total width of a digit box row for given groups e.g. [3,3,2]."""
    n = sum(groups)
    g = len(groups) - 1
    return n * (DIGIT_W + GAP) + g * (GROUP_GAP - GAP)


def _draw_boxes(c, x, y, groups):
    """Draw digit boxes. Returns x after the last box."""
    cx = x
    for gi, gs in enumerate(groups):
        for _ in range(gs):
            c.setStrokeColor(DARK)
            c.setFillColor(WHITE)
            c.setLineWidth(0.8)
            c.rect(cx, y - DIGIT_H, DIGIT_W, DIGIT_H, fill=1, stroke=1)
            cx += DIGIT_W + GAP
        if gi < len(groups) - 1:
            cx += GROUP_GAP - GAP
    return cx


def _field(c, x, y, w, label, unit, groups):
    """
    Draw one field block: outer border, label+unit on top, boxes below.
    Returns the bottom Y of the block.
    """
    bh = DIGIT_H + 8*mm
    # Outer border
    c.setFillColor(WHITE)
    c.setStrokeColor(MID_GREY)
    c.setLineWidth(0.4)
    c.rect(x, y - bh, w, bh, fill=1, stroke=1)
    # Label
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", FS_LABEL)
    c.drawString(x + 2*mm, y - 4.5*mm, label)
    # Unit tag
    if unit:
        lw = c.stringWidth(label, "Helvetica-Bold", FS_LABEL)
        c.setFont("Helvetica", FS_UNIT)
        c.setFillColor(colors.HexColor("#777"))
        c.drawString(x + 2*mm + lw + 1.5*mm, y - 4.5*mm, unit)
    # Digit boxes — centred in the field width
    rw = _box_row_width(groups)
    bx = x + (w - rw) / 2
    _draw_boxes(c, bx, y - 6.5*mm, groups)
    return y - bh


def _two_col_row(c, y, fields):
    """Draw a row of 2 fields side by side. Returns Y below row."""
    col_w = CONTENT_W / 2
    bottom = y
    for i, (lbl, unit, grp) in enumerate(fields):
        bx = MARGIN_L + i * col_w
        b = _field(c, bx, y, col_w, lbl, unit, grp)
        bottom = min(bottom, b)
    return bottom


def _three_col_row(c, y, fields):
    """Draw a row of 3 fields (for litre-only rows). Returns Y below row."""
    col_w = CONTENT_W / 3
    bottom = y
    for i, (lbl, unit, grp) in enumerate(fields):
        bx = MARGIN_L + i * col_w
        b = _field(c, bx, y, col_w, lbl, unit, grp)
        bottom = min(bottom, b)
    return bottom


def _footer_note(c, text):
    """Draw footer note at bottom of page."""
    c.setFont("Helvetica", FS_FOOTER)
    c.setFillColor(colors.HexColor("#555"))
    c.drawString(MARGIN_L, MARGIN_B, text)


# =============================================================
# SHEET 1 & 2: METER BOOK (PMS / AGO)
# =============================================================

def _sheet_meter_book(c, product, subtitle, station_id, footer):
    if product == "PMS":
        title = "Meter Book Capture (PMS)"
        instruct = (
            "Litre boxes are 6 digits (grouped 3-3); UGX boxes are 8 digits (grouped 3-3-2). "
            "Leave leading boxes blank for smaller numbers. Leave UPDF boxes blank if no UPDF activity. "
            "Do NOT write Net Sales, Loss/Gain or Stock Value \u2014 StationDeck calculates these."
        )
    else:
        title = "Meter Book Capture (AGO)"
        instruct = "Same instructions as the PMS sheet. Leave UPDF boxes blank if not applicable."

    y = _page_header(c, title, subtitle, station_id, shift_box=True)
    y = _instruction_box(c, instruct, y)
    y -= 2*mm

    for tank_num in range(1, 5):
        y = _draw_tank_block(c, product, tank_num, y)
        y -= 1.5*mm

    _footer_note(c, footer)


def _draw_tank_block(c, product, tank_num, y):
    """Draw one tank block. Returns Y below the block."""
    y = _section_bar(c, f"{product} {tank_num}", y)

    # Row 1: Opening Dip | Closing Dip | Return to Tank  (3 litre cols — 62mm each)
    y = _three_col_row(c, y, [
        ("Opening Dip", "Ltrs", [3, 3]),
        ("Closing Dip", "Ltrs", [3, 3]),
        ("Return to Tank", "Ltrs", [3, 3]),
    ])

    # Row 2: UPDF Receipt | UPDF Consumption  (2 litre cols)
    #         then Cost Price | Selling Price  (2 UGX cols)
    # Split into two rows of 2 to keep readable
    y = _two_col_row(c, y, [
        ("UPDF Receipt", "Ltrs", [3, 3]),
        ("UPDF Consumption", "Ltrs", [3, 3]),
    ])
    y = _two_col_row(c, y, [
        ("Cost Price", "UGX/L", [3, 3, 2]),
        ("Selling Price", "UGX/L", [3, 3, 2]),
    ])

    return y


# =============================================================
# SHEET 3: CASH & SALES (split across 2 pages)
# =============================================================

def _sheet_cash_sales(c, subtitle, station_id):
    instruct = (
        "Every UGX box is 8 digits (grouped 3-3-2). Leave leading boxes blank for smaller numbers. "
        "PMS/AGO fuel sales are calculated from the Meter Book \u2014 do not re-enter them here."
    )

    # ── PAGE 3A: Other Incomes + Cashless Sales ───────────────
    y = _page_header(c, "Cash & Sales Capture (1 of 2)", subtitle, station_id, shift_box=True)
    y = _instruction_box(c, instruct, y)
    y -= 1*mm

    y = _section_bar(c, "OTHER INCOMES", y, bg=HEADER_BG)
    y = _two_col_row(c, y, [("Lubes Sales", "Ltrs", [3,3]), ("Lubes Sales", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("TBA", "UGX", [3,3,2]), ("Plus Card Credits", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("LPG Sales", "KGS", [3,3]), ("LPG Sales", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Shop Sales", "UGX", [3,3,2]), ("Other Payment", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Tyre", "UGX", [3,3,2]), ("", "", [3,3,2])])
    y -= 2*mm

    y = _section_bar(c, "CASHLESS SALES", y, bg=HEADER_BG)
    y = _two_col_row(c, y, [("Plus Card \u2014 PMS", "UGX", [3,3,2]), ("Plus Card \u2014 AGO", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Plus Card \u2014 Others", "UGX", [3,3,2]), ("MOMO Pay", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Airtel Pay", "UGX", [3,3,2]), ("VISA", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Debtors / Credit", "UGX", [3,3,2]), ("", "", [3,3,2])])

    _footer_note(c, "Page 3 of 6 \u2014 Continue to page 4 for Expenses, Stock Purchases, and Banking.")
    c.showPage()

    # ── PAGE 3B: Expenses + Stock Purchases + Banking ─────────
    y = _page_header(c, "Cash & Sales Capture (2 of 2)", subtitle, station_id, shift_box=True)
    y = _instruction_box(c, instruct, y)
    y -= 1*mm

    y = _section_bar(c, "EXPENSES (THIS SHIFT)", y, bg=HEADER_BG)
    y = _two_col_row(c, y, [("UMEME", "UGX", [3,3,2]), ("Water", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Security", "UGX", [3,3,2]), ("Stationary", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Generator", "UGX", [3,3,2]), ("Meals", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Transport", "UGX", [3,3,2]), ("Salaries", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Sanitary", "UGX", [3,3,2]), ("Airtime", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Misc.", "UGX", [3,3,2]), ("Shop Packaging", "UGX", [3,3,2])])
    y -= 2*mm

    y = _section_bar(c, "OTHER STOCK PURCHASES", y, bg=HEADER_BG)
    y = _two_col_row(c, y, [("TBA", "UGX", [3,3,2]), ("LPG Accessories", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Shop Purchase", "UGX", [3,3,2]), ("", "", [3,3,2])])
    y -= 2*mm

    y = _section_bar(c, "BANKING", y, bg=HEADER_BG)
    _field(c, MARGIN_L, y, CONTENT_W / 2, "Actual Cash Banked", "UGX", [3,3,2])

    _footer_note(c,
        "Cash to Bank and Delta are calculated automatically by StationDeck from the figures above.")


# =============================================================
# SHEET 4: PRODUCT SALES TOTALS
# =============================================================

def _sheet_product_totals(c, subtitle, station_id):
    title = "Product Sales Totals"
    instruct = (
        "Write the TOTAL UGX sold for each product category this shift (8 digits, grouped 3-3-2). "
        "Category totals only \u2014 individual SKU detail is tracked separately in the master spreadsheet."
    )
    y = _page_header(c, title, subtitle, station_id, shift_box=True)
    y = _instruction_box(c, instruct, y)
    y -= 4*mm

    y = _two_col_row(c, y, [("Lubricants", "UGX", [3,3,2]), ("TBA", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("LPG Accessories", "UGX", [3,3,2]), ("LPG", "UGX", [3,3,2])])
    y = _two_col_row(c, y, [("Car Wash", "UGX", [3,3,2]), ("Shop", "UGX", [3,3,2])])
    # Solar Lanterns: half-width field
    _field(c, MARGIN_L, y, CONTENT_W / 2, "Solar Lanterns", "UGX", [3,3,2])

    _footer_note(c,
        "These totals feed the General Sales Summary. "
        "Individual SKU inventory is tracked separately in the master spreadsheet.")


# =============================================================
# SHEET 5: DAILY EXPENSES
# =============================================================

def _sheet_daily_expenses(c, subtitle, station_id):
    title = "Daily Expenses"
    instruct = (
        "Fill in ONCE PER DAY (not per shift). Records the full day's consolidated expenses for the P&L. "
        "Every box is 8 digits (grouped 3-3-2). "
        "Total Expenses is calculated automatically by StationDeck."
    )
    y = _page_header(c, title, subtitle, station_id, shift_box=False)
    y = _instruction_box(c, instruct, y)
    y -= 4*mm

    rows = [
        [("Meals", "UGX", [3,3,2]),        ("Generator", "UGX", [3,3,2])],
        [("Electricity", "UGX", [3,3,2]),   ("Water", "UGX", [3,3,2])],
        [("Salaries", "UGX", [3,3,2]),      ("Stationary", "UGX", [3,3,2])],
        [("Security", "UGX", [3,3,2]),      ("Sanitation", "UGX", [3,3,2])],
        [("Airtime/Data", "UGX", [3,3,2]),  ("Transport", "UGX", [3,3,2])],
        [("NSSF", "UGX", [3,3,2]),          ("Sundries", "UGX", [3,3,2])],
        [("Maintenance", "UGX", [3,3,2]),   ("VAT/Tax", "UGX", [3,3,2])],
        [("Photocopy", "UGX", [3,3,2]),     ("Tax Compliance", "UGX", [3,3,2])],
    ]
    for row in rows:
        y = _two_col_row(c, y, row)
        y -= 0.5*mm

    _footer_note(c,
        "Total Expenses is calculated automatically by StationDeck as the sum of all categories above.")