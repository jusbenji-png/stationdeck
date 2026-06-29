"""
web/app.py — StationDeck Web Dashboard
Phase 9B + Phase 10 (License Kill Switch) + Phase 10 wrap-up (Settings Page)
Phase 12 — Annual Report Route
Phase 13 prep — Dashboard improvements:
  /preview_month  — live monthly summary before report generation
  /fy_status      — FY month-by-month data availability overview
Phase OCR — Daily Entry System:
  /daily_entry            — new daily entry flow
  /daily_entry/import_photos — OCR photo processing
  /daily_entry/import_excel  — Excel fallback import
  /daily_entry/commit        — write confirmed entry to SQLite
Capture Templates:
  /capture_template       — generate and download station capture template PDF
Phase 14B — Authentication System:
  /register       — new station registration wizard
  /recover        — self-service machine recovery
  /login          — updated to verify against auth server with local fallback
Local Network Access:
  /network_info   — returns this machine's local IP for dashboard display
Auto-Update System:
  /check_update   — returns update status from src/updater.py background check
"""

import sys
import io
import socket

# ── stdout fix — only when a real console is attached ────────────────────────
if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file, jsonify
)
from pathlib import Path
from datetime import datetime
from io import BytesIO
import logging
import os
import json

# ── ROOT path — frozen-aware ──────────────────────────────────────────────────
# When running as a PyInstaller .exe, __file__ points inside _internal/,
# so Path(__file__).parent.parent resolves to _internal/ — NOT the install
# folder. We must use sys.executable.parent when frozen.
#
#   Development:  ROOT = stationdeck/           (project root)
#   Frozen .exe:  ROOT = C:\Program Files (x86)\StationDeck\  (install folder)
#
if getattr(sys, 'frozen', False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).parent.parent

# ── Make sure src/ is importable ─────────────────────────────────────────────
sys.path.insert(0, str(ROOT))

# Writable data root — under %LOCALAPPDATA%\StationDeck when frozen so non-admin
# users can write and data survives updates. ROOT stays read-only (install dir).
from config.settings import DATA_DIR as DATA_ROOT

from src.database import (
    init_db,
    import_from_excel,
    get_records_by_month,
    get_record_count,
    get_date_range_stored,
)
from config.station_loader import load_station_config

# ── Phase 10: License system ──────────────────────────────────────────────────
from src.license import (
    check_license,
    activate_key,
    get_machine_id,
)

# ── Phase 14B: Auth client ────────────────────────────────────────────────────
from src.auth_client import (
    register_station,
    verify_station,
    recover_station,
)

# ── Flask app setup ───────────────────────────────────────────────────────────
_WEB_DIR = Path(__file__).resolve().parent
app = Flask(__name__,
            template_folder=str(_WEB_DIR / "templates"),
            static_folder=str(_WEB_DIR / "static"))
app.secret_key = os.environ.get("FLASK_SECRET", "stationdeck-dev-secret-2026")

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Station config ────────────────────────────────────────────────────────────
# IMPORTANT: STATION_ID must remain "te_rwizi" — this is the local license
# file key and SQLite database key. Changing it breaks local license lookup.
# The display name shown in the navbar comes from session["station_name"]
# which is set to the auth server name (_jus.benji_) on cloud login.
STATION_ID       = "te_rwizi"
STATION_PASSWORD = os.environ.get("STATION_PASSWORD", "stationdeck123")

# ── Uganda Financial Year helpers ─────────────────────────────────────────────
def _fy_start_year_for(year: int, month: int) -> int:
    return year if month >= 7 else year - 1

# ── Initialise database on startup ────────────────────────────────────────────
init_db()


# ──────────────────────────────────────────────────────────────────────────────
# LOCAL NETWORK IP HELPER
# ──────────────────────────────────────────────────────────────────────────────

def _get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(1)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        pass
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:
        return "unknown"


# ──────────────────────────────────────────────────────────────────────────────
# PHASE 10 — LICENSE HELPERS
# ──────────────────────────────────────────────────────────────────────────────

PUBLIC_ROUTES = {
    "login", "logout", "activate", "activate_submit",
    "locked", "status", "register", "register_submit",
    "recover", "recover_submit", "network_info", "check_update"
}


def license_guard():
    # Dev stations (is_dev flag set on cloud login) skip all local
    # license checks — they never expire regardless of local key status.
    if session.get("auth_mode") == "cloud" and session.get("is_dev"):
        return None

    info = check_license(STATION_ID)
    if info["status"] == "not_activated":
        return redirect(url_for("activate"))
    if info["status"] in ("expired", "tampered"):
        return redirect(url_for("locked"))
    return None


def _license_context():
    # Dev sessions get a clean active license context so no expiry
    # banner ever appears in any template.
    if session.get("auth_mode") == "cloud" and session.get("is_dev"):
        return {"status": "active", "expiry": None, "plan": "dev"}
    return check_license(STATION_ID)


@app.context_processor
def inject_license():
    return {"license": _license_context()}


# ──────────────────────────────────────────────────────────────────────────────
# AUTH HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def is_logged_in():
    return session.get("logged_in") is True


def require_login():
    if not is_logged_in():
        return redirect(url_for("login"))
    return None


# ──────────────────────────────────────────────────────────────────────────────
# NETWORK INFO — public route, no login required
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/network_info")
def network_info():
    local_ip = _get_local_ip()
    port     = 5000
    if local_ip != "unknown":
        network_url = f"http://{local_ip}:{port}"
    else:
        network_url = None
    return jsonify({
        "local_ip":    local_ip,
        "network_url": network_url,
        "port":        port,
    })


# ──────────────────────────────────────────────────────────────────────────────
# CHECK UPDATE — public route, no login required
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/check_update")
def check_update():
    try:
        from src.updater import get_update_status
        return jsonify(get_update_status())
    except Exception as e:
        logger.warning(f"check_update route error: {e}")
        return jsonify({"status": "error"})


# ──────────────────────────────────────────────────────────────────────────────
# ACTIVATION AND LOCK SCREEN
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/activate", methods=["GET"])
def activate():
    machine_id = get_machine_id(STATION_ID)
    return render_template("activate.html", machine_id=machine_id)


@app.route("/activate", methods=["POST"])
def activate_submit():
    key = request.form.get("license_key", "").strip()
    if not key:
        flash("Please enter your license key.", "error")
        return redirect(url_for("activate"))

    success = activate_key(key, STATION_ID)
    if success:
        flash("License activated successfully. Welcome to StationDeck!", "success")
        return redirect(url_for("login"))
    else:
        flash(
            "That key is not valid for this machine. "
            "Check that you copied it in full, or contact support.",
            "error"
        )
        return redirect(url_for("activate"))


@app.route("/locked")
def locked():
    info = check_license(STATION_ID)
    return render_template("locked.html", license=info)


# ──────────────────────────────────────────────────────────────────────────────
# PHASE 14B — REGISTRATION
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/register", methods=["GET"])
def register():
    return render_template("register.html")


@app.route("/register", methods=["POST"])
def register_submit():
    data = request.get_json()
    if not data:
        return jsonify({"success": False, "error": "No data received."}), 400

    machine_id = get_machine_id(STATION_ID)

    success, message, license_key = register_station(
        station_name    = data.get("station_name", "").strip(),
        password        = data.get("password", ""),
        email           = data.get("email", "").strip(),
        phone           = data.get("phone", "").strip(),
        region          = data.get("region", "").strip(),
        location        = data.get("location", "").strip(),
        machine_id      = machine_id,
        app_station_id  = STATION_ID,
    )

    if success:
        return jsonify({
            "success":     True,
            "license_key": license_key,
            "message":     message,
        }), 201
    else:
        return jsonify({"success": False, "error": message}), 400


# ──────────────────────────────────────────────────────────────────────────────
# PHASE 14B — MACHINE RECOVERY
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/recover", methods=["GET"])
def recover():
    return render_template("recovery.html")


@app.route("/recover", methods=["POST"])
def recover_submit():
    data = request.get_json()
    if not data:
        return jsonify({"success": False, "error": "No data received."}), 400

    new_machine_id = get_machine_id(STATION_ID)

    success, message, license_key = recover_station(
        station_name   = data.get("station_name", "").strip(),
        password       = data.get("password", ""),
        phone          = data.get("phone", "").strip(),
        new_machine_id = new_machine_id,
    )

    if success:
        return jsonify({
            "success":     True,
            "license_key": license_key,
            "message":     message,
        }), 200
    else:
        return jsonify({"success": False, "error": message}), 400


# ──────────────────────────────────────────────────────────────────────────────
# AUTH ROUTES — LOGIN / LOGOUT
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if is_logged_in():
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        station_name = request.form.get("station_name", "").strip()
        password     = request.form.get("password", "")
        machine_id   = get_machine_id(STATION_ID)

        auth_server_error = None

        try:
            success, message, auth_data = verify_station(
                station_name = station_name,
                password     = password,
                machine_id   = machine_id,
            )

            if success:
                # The auth server wraps station info under "data" key.
                # Fall back to top-level dict for backward compatibility.
                station_data = auth_data.get("data", auth_data)

                session["logged_in"]    = True
                session["station_id"]   = STATION_ID
                session["station_name"] = station_data.get("station_name", station_name)
                session["auth_mode"]    = "cloud"
                session["is_dev"]       = station_data.get("plan") == "dev"
                flash("Welcome back!", "success")
                return redirect(url_for("dashboard"))
            else:
                auth_server_error = message

        except Exception as e:
            logger.warning(f"Auth server unreachable: {e}")
            auth_server_error = "offline"

        # ── Local fallback ────────────────────────────────────────────────────
        if password == STATION_PASSWORD:
            session["logged_in"]  = True
            session["station_id"] = STATION_ID
            session["auth_mode"]  = "local"
            session["is_dev"]     = False
            if auth_server_error == "offline":
                flash("Welcome back! (offline mode — auth server unreachable)", "success")
            else:
                flash("Welcome back!", "success")
            return redirect(url_for("dashboard"))

        if auth_server_error and auth_server_error != "offline":
            flash(auth_server_error, "error")
        else:
            flash("Incorrect password. Please try again.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


# ──────────────────────────────────────────────────────────────────────────────
# DASHBOARD
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    guard = license_guard() or require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
        # Prefer the station name from the auth server session (e.g. _jus.benji_)
        # over the local YAML config (TotalEnergies Rwizi).
        # This way the navbar and dashboard show the correct identity.
        station_name = (
            session.get("station_name")
            or station_config.get("station_name", "StationDeck")
        )
        location = station_config.get("location", "")
    except Exception:
        station_name = session.get("station_name", "StationDeck")
        location     = ""

    total_records = get_record_count(STATION_ID)
    coverage      = get_date_range_stored(STATION_ID)
    now           = datetime.now()
    monthly_df    = get_records_by_month(now.month, now.year, STATION_ID)

    quick_stats = {
        "total_records":          total_records,
        "earliest_date":          coverage.get("earliest", "—"),
        "latest_date":            coverage.get("latest", "—"),
        "current_month_days":     len(monthly_df),
        "current_month_fuel_rev": 0,
        "current_month_expenses": 0,
        "current_month_delta":    0,
        "latest_day":             None,
        "latest_day_delta":       None,
        "latest_day_anomaly":     False,
    }

    if not monthly_df.empty:
        quick_stats["current_month_fuel_rev"] = (
            monthly_df["pms_revenue"].sum() + monthly_df["ago_revenue"].sum()
        )
        quick_stats["current_month_expenses"] = monthly_df["total_expenses"].sum()
        quick_stats["current_month_delta"]    = monthly_df["delta"].sum()

        try:
            latest_row   = monthly_df.sort_values("date").iloc[-1]
            latest_delta = float(latest_row.get("delta", 0))
            quick_stats["latest_day"]         = str(latest_row["date"])[:10]
            quick_stats["latest_day_delta"]   = latest_delta
            quick_stats["latest_day_anomaly"] = abs(latest_delta) > 500_000
        except Exception:
            pass

    try:
        report_files = _list_reports(station_config)
    except Exception:
        report_files = []

    fy_options = _available_fy_options(coverage)

    return render_template(
        "dashboard.html",
        station_name=station_name,
        location=location,
        stats=quick_stats,
        report_files=report_files,
        current_month=now.strftime("%B %Y"),
        current_month_num=now.month,
        current_year=now.year,
        fy_options=fy_options,
    )


# ──────────────────────────────────────────────────────────────────────────────
# FY STATUS
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/fy_status")
def fy_status():
    guard = require_login()
    if guard:
        return jsonify({"error": "login required"}), 401

    now      = datetime.now()
    fy_start = _fy_start_year_for(now.year, now.month)
    fy_label = f"FY {fy_start}/{str(fy_start + 1)[-2:]}"

    fy_months = []
    for m in range(7, 13):
        fy_months.append((m, fy_start))
    for m in range(1, 7):
        fy_months.append((m, fy_start + 1))

    months_out = []
    for (month, year) in fy_months:
        try:
            df = get_records_by_month(month, year, STATION_ID)
        except Exception:
            df = None

        label     = datetime(year, month, 1).strftime("%b %Y")
        month_dt  = datetime(year, month, 1)
        is_future = month_dt > now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

        if is_future:
            months_out.append({
                "label": label, "month": month, "year": year,
                "has_data": False, "is_future": True,
                "days": 0, "fuel_revenue": 0, "total_sales": 0,
                "delta": 0, "delta_status": "—",
            })
            continue

        if df is None or df.empty:
            months_out.append({
                "label": label, "month": month, "year": year,
                "has_data": False, "is_future": False,
                "days": 0, "fuel_revenue": 0, "total_sales": 0,
                "delta": 0, "delta_status": "MISSING",
            })
            continue

        fuel_rev    = float(df["pms_revenue"].sum() + df["ago_revenue"].sum())
        total_sales = float(df["cashless_total"].sum()) if "cashless_total" in df.columns else fuel_rev
        delta       = float(df["delta"].sum())

        months_out.append({
            "label": label, "month": month, "year": year,
            "has_data": True, "is_future": False,
            "days": len(df),
            "fuel_revenue": fuel_rev,
            "total_sales":  total_sales,
            "delta":        delta,
            "delta_status": "SURPLUS" if delta > 0 else ("DEFICIT" if delta < 0 else "BALANCED"),
        })

    return jsonify({"fy_label": fy_label, "months": months_out})


# ──────────────────────────────────────────────────────────────────────────────
# MONTHLY PREVIEW
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/preview_month", methods=["GET"])
def preview_month():
    guard = require_login()
    if guard:
        return jsonify({"error": "login required"}), 401

    try:
        month = int(request.args.get("month", 0))
        year  = int(request.args.get("year",  0))
    except (ValueError, TypeError):
        return jsonify({"has_data": False, "error": "Invalid month or year."})

    if not month or not year:
        return jsonify({"has_data": False, "error": "Month and year are required."})

    try:
        df = get_records_by_month(month, year, STATION_ID)
    except Exception as e:
        return jsonify({"has_data": False, "error": str(e)})

    if df is None or df.empty:
        period = datetime(year, month, 1).strftime("%B %Y")
        return jsonify({
            "has_data": False,
            "period":   period,
            "message":  f"No data found for {period}. Upload the Daily Cash Flow file first.",
        })

    period      = datetime(year, month, 1).strftime("%B %Y")
    days        = len(df)
    fuel_rev    = float(df["pms_revenue"].sum() + df["ago_revenue"].sum())
    pms_vol     = float(df["pms_volume"].sum())
    ago_vol     = float(df["ago_volume"].sum())
    total_sales = float(df["cashless_total"].sum()) if "cashless_total" in df.columns else fuel_rev
    cash        = float(df["total_cash"].sum())     if "total_cash"     in df.columns else 0
    cashless    = total_sales - cash
    expenses    = float(df["total_expenses"].sum()) if "total_expenses"  in df.columns else 0
    delta       = float(df["delta"].sum())          if "delta"           in df.columns else 0
    anomalies   = int((df["delta"] != 0).sum())     if "delta"           in df.columns else 0

    cash_pct     = round(cash     / total_sales * 100, 1) if total_sales > 0 else 0
    cashless_pct = round(cashless / total_sales * 100, 1) if total_sales > 0 else 0
    delta_status = "SURPLUS" if delta > 0 else ("DEFICIT" if delta < 0 else "BALANCED")

    return jsonify({
        "has_data": True, "period": period, "days": days,
        "pms_volume": round(pms_vol, 2), "ago_volume": round(ago_vol, 2),
        "fuel_revenue": fuel_rev, "total_sales": total_sales,
        "cash_collected": cash, "cash_pct": cash_pct,
        "cashless_collected": cashless, "cashless_pct": cashless_pct,
        "total_expenses": expenses, "total_delta": delta,
        "delta_status": delta_status, "anomaly_days": anomalies,
    })


# ──────────────────────────────────────────────────────────────────────────────
# DAILY ENTRY — main page
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/daily_entry", methods=["GET"])
def daily_entry():
    guard = license_guard() or require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
        station_name   = (
            session.get("station_name")
            or station_config.get("station_name", "StationDeck")
        )
    except Exception:
        station_name = session.get("station_name", "StationDeck")

    today = datetime.now().strftime("%Y-%m-%d")

    return render_template(
        "daily_entry.html",
        station_name=station_name,
        today=today,
    )


# ──────────────────────────────────────────────────────────────────────────────
# DAILY ENTRY — OCR photo import
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/daily_entry/import_photos", methods=["POST"])
def daily_entry_import_photos():
    guard = license_guard() or require_login()
    if guard:
        return jsonify({"success": False, "message": "Access denied."}), 403

    photos     = request.files.getlist("photos")
    entry_date = request.form.get("entry_date", "")
    shift      = request.form.get("shift", "Day")

    if not photos or all(p.filename == "" for p in photos):
        return jsonify({"success": False, "message": "No photos received."})

    temp_dir = DATA_ROOT / "data" / "ocr_temp"
    temp_dir.mkdir(parents=True, exist_ok=True)

    saved_paths = []
    for photo in photos:
        if photo.filename == "":
            continue
        safe_name = f"{entry_date}_{shift}_{photo.filename}"
        save_path = temp_dir / safe_name
        photo.save(str(save_path))
        saved_paths.append(save_path)

    if not saved_paths:
        return jsonify({"success": False, "message": "No valid photo files received."})

    try:
        from src.ocr_reader import read_capture_sheets
        # Pass station credentials so OCR routes through the StationDeck server
        # (the server holds the OpenAI key — no key needed on the install).
        fields = read_capture_sheets(
            saved_paths, entry_date, shift,
            station_name=session.get("station_name"),
            machine_id=get_machine_id(STATION_ID),
        )
        return jsonify({"success": True, "fields": fields})
    except Exception as e:
        logger.error(f"OCR failed: {e}", exc_info=True)
        return jsonify({
            "success": True,
            "fields": {},
            "message": "OCR could not read the photos. Please fill in the fields manually.",
        })


# ──────────────────────────────────────────────────────────────────────────────
# DAILY ENTRY — Excel fallback import
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/daily_entry/import_excel", methods=["POST"])
def daily_entry_import_excel():
    guard = license_guard() or require_login()
    if guard:
        return jsonify({"success": False, "message": "Access denied."}), 403

    files      = request.files.getlist("excel_files")
    entry_date = request.form.get("entry_date", "")

    if not files or all(f.filename == "" for f in files):
        return jsonify({"success": False, "message": "No files received."})

    input_dir = DATA_ROOT / "data" / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    total_imported = 0
    imported_files = []

    for file in files:
        if file.filename == "" or not file.filename.endswith((".xlsx", ".xls")):
            continue

        fname = file.filename.lower()
        if "cash" in fname or "daily" in fname:
            save_name = "Daily_cash_flow.xlsx"
        elif "stock" in fname or "mvt" in fname or "movement" in fname:
            save_name = "Stock mvt-Template-2025-2026-V1-7-25.xlsx"
        elif "shop" in fname:
            save_name = "SHOP MONTHLY SALES REPORT FOR APRIL 2026 T.E. Rwizi .xlsx"
        elif "manager" in fname or "end of" in fname:
            save_name = "End of May manager's report.xlsx"
        else:
            save_name = file.filename

        save_path = input_dir / save_name
        file.save(str(save_path))

        if "cash" in fname or "daily" in fname:
            try:
                count = import_from_excel(save_path, STATION_ID)
                total_imported += count
                imported_files.append(f"{file.filename} ({count} records)")
            except Exception as e:
                logger.warning(f"Could not import {file.filename}: {e}")
                imported_files.append(f"{file.filename} (saved, not imported)")
        else:
            imported_files.append(f"{file.filename} (saved)")

    if not imported_files:
        return jsonify({"success": False, "message": "No valid Excel files were processed."})

    msg = f"Imported: {', '.join(imported_files)}."
    if total_imported > 0:
        msg = f"{total_imported} records imported. " + msg

    return jsonify({"success": True, "message": msg, "fields": {}})


# ──────────────────────────────────────────────────────────────────────────────
# DAILY ENTRY — commit confirmed data to SQLite
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/daily_entry/commit", methods=["POST"])
def daily_entry_commit():
    guard = license_guard() or require_login()
    if guard:
        return guard

    entry_date     = request.form.get("entry_date", "").strip()
    shift          = request.form.get("shift", "Day").strip()
    entry_mode     = request.form.get("entry_mode", "ocr").strip()
    entry_data_raw = request.form.get("entry_data", "{}")

    if not entry_date:
        flash("Entry date is required.", "error")
        return redirect(url_for("daily_entry"))

    try:
        entry_data = json.loads(entry_data_raw)
    except (json.JSONDecodeError, ValueError):
        flash("Could not read entry data. Please try again.", "error")
        return redirect(url_for("daily_entry"))

    try:
        from src.daily_entry_processor import process_daily_entry
        result = process_daily_entry(
            entry_date=entry_date,
            shift=shift,
            fields=entry_data,
            station_id=STATION_ID,
            mode=entry_mode,
        )

        if result.get("success"):
            audit_dir  = DATA_ROOT / "data" / "ocr_audit"
            audit_dir.mkdir(parents=True, exist_ok=True)
            audit_file = audit_dir / f"{entry_date}_{shift}_{entry_mode}.json"
            with open(str(audit_file), "w", encoding="utf-8") as f:
                json.dump({
                    "entry_date": entry_date,
                    "shift":      shift,
                    "mode":       entry_mode,
                    "committed":  datetime.now().isoformat(),
                    "fields":     entry_data,
                }, f, indent=2)

            flash(
                f"Daily entry for {entry_date} ({shift} Shift) saved successfully.",
                "success"
            )
        else:
            flash(
                "Entry could not be saved: " + result.get("message", "Unknown error."),
                "error"
            )

    except Exception as e:
        logger.error(f"Daily entry commit failed: {e}", exc_info=True)
        flash(f"Failed to save entry: {str(e)}", "error")

    return redirect(url_for("dashboard"))


# ──────────────────────────────────────────────────────────────────────────────
# UPLOAD — kept for backward compatibility
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/upload", methods=["GET", "POST"])
def upload():
    guard = license_guard() or require_login()
    if guard:
        return guard

    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file or file.filename == "":
            flash("Please select an Excel file to upload.", "error")
            return redirect(url_for("upload"))

        if not file.filename.endswith((".xlsx", ".xls")):
            flash("Only Excel files (.xlsx or .xls) are accepted.", "error")
            return redirect(url_for("upload"))

        input_dir = DATA_ROOT / "data" / "input"
        input_dir.mkdir(parents=True, exist_ok=True)
        save_path = input_dir / "Daily_cash_flow.xlsx"
        file.save(str(save_path))

        count = import_from_excel(save_path, STATION_ID)

        if count > 0:
            flash(f"Successfully imported {count} records from {file.filename}.", "success")
        else:
            flash(
                "File was saved but no records could be imported. "
                "Check that it is the correct Daily Cash Flow Excel file.",
                "error"
            )

        return redirect(url_for("dashboard"))

    return render_template("upload.html")


# ──────────────────────────────────────────────────────────────────────────────
# REPORTS LIST
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/reports")
def reports():
    guard = license_guard() or require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
    except Exception:
        station_config = None

    report_files = _list_reports(station_config)
    return render_template("reports.html", report_files=report_files)


# ──────────────────────────────────────────────────────────────────────────────
# MONTHLY REPORT GENERATION
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/generate", methods=["POST"])
def generate():
    guard = license_guard() or require_login()
    if guard:
        return jsonify({"success": False, "message": "Access denied."}), 403

    month = request.form.get("month", type=int)
    year  = request.form.get("year",  type=int)

    if not month or not year:
        return jsonify({"success": False, "message": "Month and year are required."})

    try:
        station_config = load_station_config(STATION_ID)

        df_check = get_records_by_month(month, year, STATION_ID)
        if df_check.empty:
            month_name = datetime(year, month, 1).strftime("%B %Y")
            return jsonify({
                "success": False,
                "message": (
                    f"No data found in the database for {month_name}. "
                    f"Please upload the Daily Cash Flow file for that month first."
                )
            })

        from src.processor import process_monthly_report
        from src.ai_engine import generate_report
        from src.exporter  import ExportEngine

        metrics      = process_monthly_report(month, year, station_config)
        period_label = datetime(year, month, 1).strftime("%B %Y")
        ai_result    = generate_report(
            metrics, period_label,
            station_name=session.get("station_name"),
            machine_id=get_machine_id(STATION_ID),
        )
        sections     = ai_result.get("sections", {})
        daily_df     = metrics.get("daily_df", df_check)

        engine    = ExportEngine(station_config=station_config)
        pdf_path  = engine.generate_pdf(metrics, sections, period_label)
        docx_path = engine.generate_docx(metrics, sections, period_label)
        xlsx_path = engine.generate_xlsx(metrics, daily_df, period_label)

        return jsonify({
            "success": True,
            "message": f"Monthly report for {period_label} generated successfully.",
            "files": {
                "pdf":  Path(pdf_path).name  if pdf_path  else None,
                "docx": Path(docx_path).name if docx_path else None,
                "xlsx": Path(xlsx_path).name if xlsx_path else None,
            }
        })

    except Exception as e:
        logger.error(f"Monthly report generation failed: {e}", exc_info=True)
        return jsonify({"success": False, "message": f"Generation failed: {str(e)}"})


# ──────────────────────────────────────────────────────────────────────────────
# ANNUAL REPORT GENERATION
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/generate_annual", methods=["POST"])
def generate_annual():
    guard = license_guard() or require_login()
    if guard:
        return jsonify({"success": False, "message": "Access denied."}), 403

    report_mode   = request.form.get("report_mode", "fy")   # "fy" or "last12"
    fy_start_year = request.form.get("fy_start_year", type=int)

    if report_mode == "fy" and not fy_start_year:
        return jsonify({"success": False, "message": "Financial year is required."})

    try:
        station_config = load_station_config(STATION_ID)

        from src.ai_engine  import generate_annual_report
        from src.exporter   import ExportEngine

        if report_mode == "last12":
            from src.processor import process_last12_months_report
            logger.info("Generating annual report for last 12 months of data")
            metrics = process_last12_months_report(station_config)
        else:
            from src.processor import process_annual_report
            logger.info(f"Generating annual report for FY {fy_start_year}/{str(fy_start_year+1)[-2:]}")
            metrics = process_annual_report(fy_start_year, station_config)

        fy_label     = metrics.get("fy_label", "Annual Report")
        period_label = metrics.get("period_label", fy_label)

        months_with_data = metrics.get("months_with_data", 0)
        if months_with_data == 0:
            return jsonify({
                "success": False,
                "message": (
                    f"No data found for {period_label}. "
                    f"Please ensure the Daily Cash Flow file has been uploaded."
                )
            })

        ai_result = generate_annual_report(
            metrics, period_label,
            station_name=session.get("station_name"),
            machine_id=get_machine_id(STATION_ID),
        )
        sections  = ai_result.get("sections", {})

        engine    = ExportEngine(station_config=station_config)
        pdf_path  = engine.generate_annual_pdf(metrics, sections, period_label)
        docx_path = engine.generate_annual_docx(metrics, sections, period_label)
        xlsx_path = engine.generate_annual_xlsx(metrics, period_label)

        return jsonify({
            "success": True,
            "message": (
                f"Annual report for {fy_label} generated successfully. "
                f"({months_with_data}/12 months of data included)"
            ),
            "files": {
                "pdf":  Path(pdf_path).name  if pdf_path  else None,
                "docx": Path(docx_path).name if docx_path else None,
                "xlsx": Path(xlsx_path).name if xlsx_path else None,
            }
        })

    except Exception as e:
        logger.error(f"Annual report generation failed: {e}", exc_info=True)
        return jsonify({"success": False, "message": f"Annual report failed: {str(e)}"})


# ──────────────────────────────────────────────────────────────────────────────
# EMAIL REPORT — send generated report files to configured recipients
# Routed through the StationDeck server (server holds SMTP credentials).
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/email_report", methods=["POST"])
def email_report():
    guard = license_guard() or require_login()
    if guard:
        return jsonify({"success": False, "message": "Access denied."}), 403

    import base64
    import requests as _rq

    data         = request.get_json(silent=True) or {}
    files        = data.get("files", {})            # {"pdf": name, "docx": name, "xlsx": name}
    period_label = data.get("period_label", "Report")

    station_config = load_station_config(STATION_ID)
    recipients     = station_config.get("recipients", []) or []
    if not recipients:
        return jsonify({"success": False,
                        "message": "No recipient emails are configured for this station."})

    report_dirs = station_config.get("report_dirs", {})
    attachments = []
    for kind in ("pdf", "docx", "xlsx"):
        name = files.get(kind)
        if not name:
            continue
        fpath = Path(report_dirs.get(kind, "")) / name
        if fpath.exists():
            with open(fpath, "rb") as fh:
                attachments.append({
                    "filename":    name,
                    "content_b64": base64.b64encode(fh.read()).decode("utf-8"),
                })

    if not attachments:
        return jsonify({"success": False,
                        "message": "Report files not found — generate the report first."})

    station_label = session.get("station_name") or station_config.get("station_name", "StationDeck")
    payload = {
        "station_name": session.get("station_name"),
        "machine_id":   get_machine_id(STATION_ID),
        "recipients":   recipients,
        "subject":      f"StationDeck | {station_label} — {period_label} Report",
        "body":         (f"<p>Please find attached the <strong>{period_label}</strong> "
                         f"report for {station_label}.</p>"
                         f"<p>Generated by StationDeck.</p>"),
        "attachments":  attachments,
    }
    try:
        r = _rq.post("https://web-production-46077.up.railway.app/email_report",
                     json=payload, timeout=120)
        d = r.json()
        return jsonify({"success": d.get("success", False),
                        "message": d.get("message", "")})
    except Exception as e:
        logger.error(f"Email report failed: {e}", exc_info=True)
        return jsonify({"success": False, "message": f"Could not reach email server: {e}"})


# ──────────────────────────────────────────────────────────────────────────────
# DOWNLOAD
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/download/<report_type>/<filename>")
def download(report_type, filename):
    guard = require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
        report_dirs    = station_config.get("report_dirs", {})
        file_path      = Path(report_dirs.get(report_type, "")) / filename

        if not file_path.exists():
            flash(f"File not found: {filename}", "error")
            return redirect(url_for("reports"))

        return send_file(str(file_path), as_attachment=True)

    except Exception as e:
        flash(f"Download failed: {e}", "error")
        return redirect(url_for("reports"))


# ──────────────────────────────────────────────────────────────────────────────
# CAPTURE TEMPLATE PDF DOWNLOAD
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/capture_template")
def capture_template():
    guard = license_guard() or require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
        station_name   = (
            session.get("station_name")
            or station_config.get("station_name", "StationDeck")
        )
        location   = station_config.get("location", "")
        station_id = station_config.get("station_id", STATION_ID)
    except Exception:
        station_name = session.get("station_name", "StationDeck")
        location     = ""
        station_id   = STATION_ID

    try:
        from src.capture_template_builder import build_capture_template_pdf

        buf = BytesIO()
        build_capture_template_pdf(
            buffer=buf,
            station_name=station_name,
            location=location,
            station_id=station_id,
        )
        buf.seek(0)

        filename = f"StationDeck_CaptureTemplates_{station_id}.pdf"
        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as e:
        logger.error(f"Capture template generation failed: {e}", exc_info=True)
        flash(f"Could not generate capture template: {str(e)}", "error")
        return redirect(url_for("settings"))


# ──────────────────────────────────────────────────────────────────────────────
# SETTINGS PAGE
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/settings")
def settings():
    guard = license_guard() or require_login()
    if guard:
        return guard

    try:
        station_config = load_station_config(STATION_ID)
        station_name   = (
            session.get("station_name")
            or station_config.get("station_name", "StationDeck")
        )
        location = station_config.get("location", "")
        operator = station_config.get("operator", "")
    except Exception:
        station_name = session.get("station_name", "StationDeck")
        location     = ""
        operator     = ""

    machine_id  = get_machine_id(STATION_ID)
    lic         = _license_context()
    data_path   = str(DATA_ROOT / "data")
    local_ip    = _get_local_ip()
    network_url = f"http://{local_ip}:5000" if local_ip != "unknown" else None

    try:
        from src.updater import get_installed_version
        installed_version = get_installed_version()
    except Exception:
        installed_version = "1.0.0"

    return render_template(
        "settings.html",
        station_name=station_name,
        location=location,
        operator=operator,
        machine_id=machine_id,
        lic=lic,
        data_path=data_path,
        software_version=installed_version,
        local_ip=local_ip,
        network_url=network_url,
    )


# ──────────────────────────────────────────────────────────────────────────────
# STATUS — always public
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/status")
def status():
    count    = get_record_count(STATION_ID)
    coverage = get_date_range_stored(STATION_ID)
    lic      = check_license(STATION_ID)
    local_ip = _get_local_ip()

    try:
        from src.updater import get_installed_version
        version = get_installed_version()
    except Exception:
        version = "1.0.0"

    return jsonify({
        "status":      "ok",
        "station":     STATION_ID,
        "version":     version,
        "records":     count,
        "earliest":    coverage.get("earliest"),
        "latest":      coverage.get("latest"),
        "license":     lic["status"],
        "expiry":      lic["expiry"],
        "local_ip":    local_ip,
        "network_url": f"http://{local_ip}:5000" if local_ip != "unknown" else None,
    })


# ──────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────

def _list_reports(station_config):
    files = []
    if not station_config:
        return files

    report_dirs = station_config.get("report_dirs", {})

    for rtype in ["pdf", "docx", "xlsx"]:
        dir_path = Path(report_dirs.get(rtype, ""))
        if not dir_path.exists():
            continue
        for f in sorted(dir_path.glob("StationDeck_*"), reverse=True):
            is_annual = "Annual" in f.name
            files.append({
                "name":       f.name,
                "type":       rtype.upper(),
                "type_lower": rtype,
                "size_kb":    round(f.stat().st_size / 1024, 1),
                "modified":   datetime.fromtimestamp(
                    f.stat().st_mtime
                ).strftime("%Y-%m-%d %H:%M"),
                "is_annual":  is_annual,
                "label":      "Annual" if is_annual else "Monthly",
            })

    return files


def _available_fy_options(coverage: dict) -> list:
    earliest = coverage.get("earliest")
    latest   = coverage.get("latest")

    if not earliest or not latest:
        return []

    try:
        from datetime import date
        if isinstance(earliest, str):
            earliest = datetime.strptime(earliest[:10], "%Y-%m-%d").date()
        if isinstance(latest, str):
            latest = datetime.strptime(latest[:10], "%Y-%m-%d").date()

        fy_start = _fy_start_year_for(earliest.year, earliest.month)
        fy_end   = _fy_start_year_for(latest.year,   latest.month)

        options = []
        for fy in range(fy_start, fy_end + 1):
            options.append({
                "fy_start_year": fy,
                "label":         f"FY {fy}/{str(fy+1)[-2:]} "
                                 f"(July {fy} – June {fy+1})",
            })
        options.reverse()
        return options

    except Exception:
        return []


# ──────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Start the background update checker so the dashboard banner works
    # when running via `python web\app.py` (not just via launcher.py).
    from src.updater import start_update_check
    start_update_check()

    local_ip = _get_local_ip()
    print("\n" + "=" * 55)
    print("  StationDeck Web Dashboard")
    print(f"  This PC:        http://127.0.0.1:5000")
    if local_ip != "unknown":
        print(f"  Network (WiFi): http://{local_ip}:5000")
    print("  Press CTRL+C to stop")
    print("=" * 55 + "\n")
    app.run(debug=True, host="0.0.0.0", port=5000)