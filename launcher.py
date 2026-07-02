"""
launcher.py
-----------
StationDeck — Windows Desktop Launcher

Entry point for the PyInstaller .exe build.

What it does:
  1. Adds correct paths so all imports resolve
  2. Starts Flask in a background thread (no terminal)
  3. Waits until Flask is ready (polls localhost:5000)
  4. Opens the default browser to http://localhost:5000
  5. Keeps running until process is killed
  6. Binds to 0.0.0.0 so any device on the same WiFi can access the dashboard
  7. Starts the auto-update check in the background after Flask is ready

The manager sees:
  - Browser opens automatically to the login page
  - No terminal window, no Python visible
  - Local network IP shown in dashboard for other devices
  - Update banner on dashboard if a new version is available

Usage in development:
  python launcher.py

Usage as .exe:
  StationDeck.exe
"""

import sys
import os
import time
import threading
import webbrowser
import traceback
import socket
from pathlib import Path

# ── Console safety ────────────────────────────────────────────────────────────
# In a windowed (console=False) frozen build, sys.stdout/stderr are None, and on
# Windows the console codec may be cp1252 — either one makes a stray print() of a
# unicode char (e.g. the ✅/❌ in processor logs) raise and abort whatever route
# is running (e.g. report generation). We make prints safe by swapping a None
# stream for a discarding stream. IMPORTANT: that stream must NOT expose fileno(),
# otherwise Flask's startup banner (via click) probes it as a Windows console and
# crashes with "I/O operation on closed file".
class _NullStream:
    encoding = "utf-8"
    errors   = "replace"
    def write(self, *a, **k):   return 0
    def writelines(self, *a):   pass
    def flush(self):            pass
    def isatty(self):           return False
    def writable(self):         return True
    def readable(self):         return False
    def seekable(self):         return False
    def close(self):            pass
    # deliberately no fileno() — click then treats this as a non-console stream

def _make_streams_safe():
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            setattr(sys, name, _NullStream())
        else:
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

_make_streams_safe()

# ── Path setup ────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BUNDLE_DIR  = Path(sys._MEIPASS)
    INSTALL_DIR = Path(sys.executable).parent
    if str(BUNDLE_DIR) not in sys.path:
        sys.path.insert(0, str(BUNDLE_DIR))
else:
    BUNDLE_DIR  = Path(__file__).parent
    INSTALL_DIR = Path(__file__).parent
    if str(BUNDLE_DIR) not in sys.path:
        sys.path.insert(0, str(BUNDLE_DIR))

# ── Constants ─────────────────────────────────────────────────────────────────
# Bind to 0.0.0.0 so the app is reachable from any device on the same WiFi.
# The browser on this PC still opens to 127.0.0.1 (localhost) as normal.
HOST      = "0.0.0.0"
PORT      = 5000
LOCAL_URL = f"http://127.0.0.1:{PORT}"      # used by THIS PC's browser
TIMEOUT   = 30


# ── Local IP detection ────────────────────────────────────────────────────────

def _get_local_ip() -> str:
    """
    Return the machine's LAN IP address (e.g. 192.168.3.44).

    Technique: open a UDP socket toward an external address.
    No data is actually sent — this just makes the OS choose
    which network interface it would use, and we read that IP.
    """
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


# ── Startup log ───────────────────────────────────────────────────────────────

def _get_log_path() -> Path:
    # Logs are writable → %LOCALAPPDATA%\StationDeck when frozen (not Program Files).
    if getattr(sys, 'frozen', False):
        appdata = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        base = Path(appdata) / "StationDeck" if appdata else Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    log_dir = base / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "startup.log"


def _log(msg: str):
    """Write a timestamped line to the startup log and print it."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        with open(_get_log_path(), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    try:
        print(line)
    except Exception:
        pass


# ── Directory setup ───────────────────────────────────────────────────────────

def _ensure_dirs():
    """Create all essential directories before Flask starts."""
    try:
        # All writable dirs live under DATA_DIR (%LOCALAPPDATA%\StationDeck when
        # frozen). Station YAML is read-only and ships in the install dir.
        from config.settings import DATA_DIR

        # Report folders are per-station — resolve the registered station id
        # (falls back to te_rwizi for pre-registration installs).
        station_id = "te_rwizi"
        try:
            id_file = DATA_DIR / "station_id.txt"
            if id_file.exists():
                val = id_file.read_text(encoding="utf-8").strip()
                if val:
                    station_id = val
        except Exception:
            pass

        dirs = [
            DATA_DIR / "data" / "input",
            DATA_DIR / "data" / "processed",
            DATA_DIR / "data" / "ocr_temp",
            DATA_DIR / "data" / "ocr_audit",
            DATA_DIR / "logs",
            DATA_DIR / "reports" / station_id / "pdf",
            DATA_DIR / "reports" / station_id / "docx",
            DATA_DIR / "reports" / station_id / "xlsx",
            DATA_DIR / "reports" / station_id / "archive",
        ]
        for d in dirs:
            d.mkdir(parents=True, exist_ok=True)
        _log(f"Directories verified under: {DATA_DIR}")
    except Exception as e:
        _log(f"WARNING: _ensure_dirs() failed: {e}")


# ── Flask thread ──────────────────────────────────────────────────────────────

_flask_error: list = []


def _run_flask():
    """Start Flask in this daemon thread."""
    try:
        _log("Flask thread starting...")
        _ensure_dirs()

        from web.app import app
        _log("Flask app imported successfully.")
        _log(f"Binding to {HOST}:{PORT} (accessible on local network)...")

        local_ip = _get_local_ip()
        if local_ip != "unknown":
            _log(f"Network access: http://{local_ip}:{PORT}")
        else:
            _log("Could not detect local IP — network access may be unavailable.")

        app.run(
            host=HOST,
            port=PORT,
            debug=False,
            use_reloader=False,
            threaded=True,
        )

    except Exception:
        err = traceback.format_exc()
        _log(f"FATAL: Flask thread crashed:\n{err}")
        _flask_error.append(err)


# ── Health check ──────────────────────────────────────────────────────────────

def _wait_for_flask(timeout: int = TIMEOUT) -> bool:
    """Poll localhost:5000 until Flask responds or we time out."""
    import urllib.request

    deadline = time.time() + timeout
    while time.time() < deadline:
        if _flask_error:
            return False
        try:
            urllib.request.urlopen(LOCAL_URL, timeout=1)
            return True
        except Exception:
            time.sleep(0.3)
    return False


# ── Main ──────────────────────────────────────────────────────────────────────

def _already_running() -> bool:
    """True if another StationDeck is already serving port 5000.

    Windows lets two processes bind the same port (SO_REUSEADDR), which
    silently splits traffic between two app instances with separate
    databases. Never start a second instance — just reuse the first.
    """
    import urllib.request
    try:
        with urllib.request.urlopen(f"{LOCAL_URL}/status", timeout=2) as r:
            return b"StationDeck" in r.read(2048) or r.status == 200
    except Exception:
        return False


def main():
    if _already_running():
        _log("StationDeck is already running — opening browser to the existing instance.")
        webbrowser.open(LOCAL_URL)
        sys.exit(0)

    _log("=" * 50)
    _log("StationDeck launcher starting")
    _log(f"Python: {sys.version}")
    _log(f"Frozen: {getattr(sys, 'frozen', False)}")
    if getattr(sys, 'frozen', False):
        _log(f"Executable: {sys.executable}")
        _log(f"Bundle dir: {sys._MEIPASS}")

    local_ip = _get_local_ip()
    _log(f"This machine's local IP: {local_ip}")
    if local_ip != "unknown":
        _log(f"Other devices on this WiFi can access: http://{local_ip}:{PORT}")
    _log("=" * 50)

    # Start Flask daemon thread
    flask_thread = threading.Thread(target=_run_flask, daemon=True)
    flask_thread.start()

    _log(f"Waiting for Flask on {LOCAL_URL}...")
    ready = _wait_for_flask()

    if _flask_error:
        _log("Flask failed to start. Check logs/startup.log for details.")
        time.sleep(3)
        sys.exit(1)

    if ready:
        _log("Flask ready — opening browser.")
        webbrowser.open(LOCAL_URL)
    else:
        _log(
            f"Flask did not respond within {TIMEOUT}s. "
            f"Try opening {LOCAL_URL} manually."
        )
        webbrowser.open(LOCAL_URL)

    # ── Start update check in background ─────────────────────────────────────
    # We do this AFTER Flask is confirmed ready so it never delays startup.
    #
    # WHY WE WAIT 5 SECONDS before starting the check:
    # The browser opens immediately when Flask is ready. If we start the
    # update check at exactly the same moment, the dashboard loads and calls
    # /check_update before the background thread has had time to contact
    # Railway (which takes 2-5 seconds on a typical connection). The result
    # is "pending" and the banner never shows.
    #
    # By delaying 5 seconds, the update check completes BEFORE the user
    # has finished logging in, so /check_update returns the real result
    # the first time the dashboard loads.
    #
    # The dashboard JS also retries /check_update once after 8 seconds
    # as a safety net for slow connections.
    def _delayed_update_check():
        time.sleep(5)
        try:
            from src.updater import start_update_check
            start_update_check()
            _log("Update check started in background (after 5s delay).")
        except Exception as e:
            _log(f"WARNING: Could not start update checker: {e}")

    threading.Thread(target=_delayed_update_check, daemon=True).start()

    # ── Automatic weekly data backup (background, never blocks startup) ──
    def _auto_backup():
        time.sleep(15)   # let the app settle first
        try:
            from config.settings import DATA_DIR
            from src.backup import run_backup_if_due
            if run_backup_if_due(DATA_DIR):
                _log("Weekly data backup created in backups folder.")
        except Exception as e:
            _log(f"WARNING: Automatic backup failed: {e}")

    threading.Thread(target=_auto_backup, daemon=True).start()

    # Keep main thread alive so the daemon Flask thread keeps running
    try:
        while flask_thread.is_alive():
            time.sleep(1)
        if _flask_error:
            _log("Flask thread exited with an error. Check logs/startup.log.")
        else:
            _log("Flask thread exited cleanly.")
    except KeyboardInterrupt:
        _log("StationDeck shutting down (KeyboardInterrupt).")
        sys.exit(0)


if __name__ == "__main__":
    main()