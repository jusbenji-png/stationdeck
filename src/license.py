# src/license.py
# StationDeck - Phase 10: License Key System
# Handles machine binding, key generation, activation, and status checks.
# Designed to be reused as-is by Phase 15 (Flutterwave auto-billing).

import os
import sys
import json
import hashlib
import base64
import platform
from datetime import date, datetime, timedelta
from pathlib import Path

# -- Cryptography (AES via Fernet) --------------------------------------------
from cryptography.fernet import Fernet

# -- Project root detection ---------------------------------------------------
# When frozen as .exe, __file__ points inside _internal bundle folder.
# We must use sys.executable's parent (the actual install folder) instead.
#
# Development:   BASE_DIR = stationdeck/          (project root)
# Frozen .exe:   BASE_DIR = dist/StationDeck/     (exe folder)

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
    # Program Files is read-only for normal users; write license.key to APPDATA instead
    _appdata  = Path(os.environ.get("APPDATA", BASE_DIR))
    CONFIG_DIR = _appdata / "StationDeck"
else:
    BASE_DIR   = Path(__file__).resolve().parent.parent
    CONFIG_DIR = BASE_DIR / "config"

KEY_FILE = CONFIG_DIR / "license.key"


# ------------------------------------------------------------------------------
# SECTION 1 - MACHINE IDENTITY
# ------------------------------------------------------------------------------

def _get_windows_machine_guid() -> str:
    """
    Read the Windows Machine GUID from the registry.
    This value is unique per Windows installation and survives reboots.
    If we can't read it (non-Windows or registry error), we fall back
    to the hostname - less robust but still useful for development.
    """
    if platform.system() == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography"
            )
            guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            winreg.CloseKey(key)
            return guid
        except Exception:
            pass  # fall through to hostname fallback

    # Fallback: hostname (used on Linux/Mac during development)
    return platform.node()


def get_machine_id(station_id: str) -> str:
    """
    Returns a stable SHA-256 hash that uniquely identifies this machine
    for the given station installation.

    Formula: SHA256(MachineGuid + ":" + station_id)

    WHY machine + station?  A key issued for te_rwizi on Machine A
    cannot be reused on Machine B, and cannot be reused for a different
    station installed on the same machine.

    This is the value the manager sees in Settings -> copy to clipboard
    -> WhatsApp to Benjamin when they get a new computer.
    """
    raw = f"{_get_windows_machine_guid()}:{station_id}"
    return hashlib.sha256(raw.encode()).hexdigest()


# ------------------------------------------------------------------------------
# SECTION 2 - ENCRYPTION HELPERS
# ------------------------------------------------------------------------------

def _derive_fernet_key(machine_id: str) -> bytes:
    """
    Derive a 32-byte Fernet-compatible encryption key from the machine_id.

    We use the first 32 bytes of SHA-256(machine_id), then base64url-encode
    them - which is exactly what Fernet expects.

    This means the encrypted license.key file is ONLY decryptable on the
    machine it was activated on. If someone copies the folder to another PC,
    the machine_id changes -> decryption fails -> software stays locked.
    """
    raw_key = hashlib.sha256(machine_id.encode()).digest()   # 32 bytes
    return base64.urlsafe_b64encode(raw_key)                  # Fernet key


def _encrypt_payload(payload: dict, machine_id: str) -> str:
    """
    AES-encrypt a dict payload (via Fernet) and return a base64 string.
    Used when generating and saving a license key.
    """
    fernet = Fernet(_derive_fernet_key(machine_id))
    json_bytes = json.dumps(payload).encode()
    return fernet.encrypt(json_bytes).decode()


def _decrypt_payload(encrypted: str, machine_id: str) -> dict:
    """
    Decrypt an AES-encrypted payload string back to a dict.
    Raises cryptography.fernet.InvalidToken if the machine_id is wrong
    or the file has been tampered with.
    """
    fernet = Fernet(_derive_fernet_key(machine_id))
    json_bytes = fernet.decrypt(encrypted.encode())
    return json.loads(json_bytes)


# ------------------------------------------------------------------------------
# SECTION 3 - KEY FORMAT  SD-YYYY-XXXX-XXXX
# ------------------------------------------------------------------------------

def _build_key_string(payload: dict, machine_id: str) -> str:
    """
    Encode the full encrypted payload into the printable key string that
    Benjamin sends (or Flutterwave emails) to the station.

    Format: SD-<YEAR>-<4 chars>-<4 chars>-<base64 blob>

    The first three segments are human-readable:
      SD        = StationDeck product prefix
      YEAR      = year of issue (easy visual check)
      XXXX-XXXX = first 8 chars of machine_id (helps Benjamin spot
                  which machine a key was made for during support calls)

    Everything after the third dash is the AES-encrypted payload,
    base64url-encoded (no padding = no extra chars).  The station never
    needs to type this by hand - they paste it from the email.
    """
    year_str   = str(payload["issued"][:4])   # e.g. "2026"
    mid_short  = machine_id[:4].upper()       # first 4 hex chars
    mid_short2 = machine_id[4:8].upper()      # next 4 hex chars
    encrypted  = _encrypt_payload(payload, machine_id)
    # Remove any base64 padding '=' chars to keep the string clean
    encrypted_clean = encrypted.replace("=", "")
    return f"SD-{year_str}-{mid_short}-{mid_short2}-{encrypted_clean}"


def _parse_key_string(key: str, machine_id: str) -> dict:
    """
    Parse and decrypt a key string produced by _build_key_string().
    Returns the payload dict on success.
    Raises ValueError if the format is wrong, InvalidToken if tampered.
    """
    parts = key.strip().split("-", 4)   # split at first 4 dashes only
    if len(parts) != 5 or parts[0] != "SD":
        raise ValueError("Invalid key format - must start with SD-YYYY-XXXX-XXXX-...")

    encrypted_clean = parts[4]
    # Re-add padding if needed (Fernet requires correct base64 padding)
    padding = 4 - (len(encrypted_clean) % 4)
    if padding != 4:
        encrypted_clean += "=" * padding

    return _decrypt_payload(encrypted_clean, machine_id)


# ------------------------------------------------------------------------------
# SECTION 4 - KEY GENERATION  (Benjamin's tool / Phase 15 webhook)
# ------------------------------------------------------------------------------

def generate_key(
    station_id: str,
    machine_id: str,
    expiry_date: date,
    is_trial: bool = False
) -> str:
    """
    Generate a printable license key for a given station + machine.

    Parameters
    ----------
    station_id  : e.g. "te_rwizi"
    machine_id  : SHA-256 hash from get_machine_id(station_id)
    expiry_date : date object - when the license expires
    is_trial    : True for the first 5 activations (30-day free trial)

    Returns
    -------
    A printable key string: SD-2026-XXXX-XXXX-<encrypted blob>

    This function is called:
    - NOW (Phase 10): manually by Benjamin via generate_key.py
    - LATER (Phase 15): automatically by the Flutterwave payment webhook
    """
    payload = {
        "station_id": station_id,
        "machine_id": machine_id,
        "expiry_date": expiry_date.isoformat(),   # "2026-06-27"
        "is_trial":   is_trial,
        "issued":     date.today().isoformat(),
    }
    return _build_key_string(payload, machine_id)


# ------------------------------------------------------------------------------
# SECTION 5 - ACTIVATION  (station enters key in wizard)
# ------------------------------------------------------------------------------

def activate_key(key: str, station_id: str) -> bool:
    """
    Validate a license key and, if valid, save it to config/license.key.

    Validation checks:
    1. Key can be decrypted with THIS machine's derived Fernet key.
    2. station_id in payload matches the station_id passed in.
    3. machine_id in payload matches THIS machine's machine_id.
    4. expiry_date has not already passed.

    If all checks pass:
    - The encrypted payload is saved to config/license.key.
    - Returns True.

    If any check fails:
    - Returns False (never raises - the UI catches False and shows error).

    WHY save the encrypted blob and not the plain payload?
    Because storing plain JSON on disk would let a technical user extend
    their expiry date with a text editor. The encrypted file cannot be
    tampered with without knowing the machine_id.
    """
    try:
        machine_id = get_machine_id(station_id)
        payload = _parse_key_string(key, machine_id)

        # Check station match
        if payload.get("station_id") != station_id:
            return False

        # Check machine match
        if payload.get("machine_id") != machine_id:
            return False

        # Check expiry
        expiry = date.fromisoformat(payload["expiry_date"])
        if expiry < date.today():
            return False

        # All good - save to disk
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        encrypted_blob = _encrypt_payload(payload, machine_id)
        KEY_FILE.write_text(encrypted_blob, encoding="utf-8")
        return True

    except Exception:
        return False


# ------------------------------------------------------------------------------
# SECTION 6 - STATUS CHECK  (called on every protected route)
# ------------------------------------------------------------------------------

def check_license(station_id: str = "te_rwizi") -> dict:
    """
    Read the saved license and return its current status.

    Returns a dict with these keys every time (never raises):

    {
        "valid":          bool   - True if active and not expired
        "expiry":         str    - "YYYY-MM-DD" or None
        "days_remaining": int    - calendar days until expiry, or 0
        "status":         str    - see STATUS VALUES below
        "is_trial":       bool   - True if this is a trial key
    }

    STATUS VALUES
    -------------
    "not_activated"    - license.key file does not exist
    "tampered"         - file exists but cannot be decrypted (wrong machine)
    "expired"          - expiry_date is in the past
    "trial_expiring"   - is_trial=True AND <=3 days remaining
    "expiring_soon"    - is_trial=False AND <=7 days remaining
    "renew_reminder"   - is_trial=False AND <=30 days remaining
    "trial"            - is_trial=True AND >3 days remaining
    "active"           - is_trial=False AND >30 days remaining
    """
    empty = {
        "valid": False,
        "expiry": None,
        "days_remaining": 0,
        "status": "not_activated",
        "is_trial": False,
    }

    if not KEY_FILE.exists():
        return empty

    try:
        machine_id = get_machine_id(station_id)
        encrypted_blob = KEY_FILE.read_text(encoding="utf-8")
        payload = _decrypt_payload(encrypted_blob, machine_id)
    except Exception:
        return {**empty, "status": "tampered"}

    expiry       = date.fromisoformat(payload["expiry_date"])
    today        = date.today()
    days_left    = (expiry - today).days
    is_trial     = payload.get("is_trial", False)

    if days_left < 0:
        status = "expired"
        valid  = False
    elif is_trial and days_left <= 3:
        status = "trial_expiring"
        valid  = True
    elif not is_trial and days_left <= 7:
        status = "expiring_soon"
        valid  = True
    elif not is_trial and days_left <= 30:
        status = "renew_reminder"
        valid  = True
    elif is_trial:
        status = "trial"
        valid  = True
    else:
        status = "active"
        valid  = True

    return {
        "valid":          valid,
        "expiry":         expiry.isoformat(),
        "days_remaining": max(days_left, 0),
        "status":         status,
        "is_trial":       is_trial,
    }


# ------------------------------------------------------------------------------
# SECTION 7 - CONVENIENCE HELPERS
# ------------------------------------------------------------------------------

def get_expiry_date(station_id: str = "te_rwizi") -> date | None:
    """Return the expiry date as a date object, or None if not activated."""
    info = check_license(station_id)
    if info["expiry"]:
        return date.fromisoformat(info["expiry"])
    return None


def is_expiring_soon(station_id: str = "te_rwizi", days: int = 30) -> bool:
    """
    Return True if the license expires within `days` calendar days.
    Used by base.html to decide whether to show the yellow renewal banner.
    """
    info = check_license(station_id)
    if not info["valid"]:
        return False
    return info["days_remaining"] <= days