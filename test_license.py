# test_license.py
# Run from project root: python test_license.py
# Tests every function in src/license.py in order.
# Paste the full terminal output back to Claude after running.

import sys
import os
from pathlib import Path
from datetime import date, timedelta

# Allow imports from src/
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.license import (
    get_machine_id,
    generate_key,
    activate_key,
    check_license,
    get_expiry_date,
    is_expiring_soon,
    KEY_FILE,
)

STATION_ID = "te_rwizi"
SEP = "-" * 60

def section(title):
    print(f"\n{SEP}\nTEST: {title}\n{SEP}")

# ── 1. Machine ID ─────────────────────────────────────────────────────────────
section("get_machine_id()")
machine_id = get_machine_id(STATION_ID)
print(f"machine_id : {machine_id}")
assert len(machine_id) == 64, "FAIL: expected 64-char SHA-256 hex"
print("PASS: 64-char SHA-256 ✓")

# ── 2. Generate a valid 30-day key ────────────────────────────────────────────
section("generate_key() — valid 30-day trial key")
expiry_30 = date.today() + timedelta(days=30)
key_trial = generate_key(STATION_ID, machine_id, expiry_30, is_trial=True)
print(f"key        : {key_trial[:60]}...")
assert key_trial.startswith("SD-"), "FAIL: key must start with SD-"
print("PASS: key format starts with SD- ✓")

# ── 3. Generate an already-expired key ───────────────────────────────────────
section("generate_key() — already-expired key")
expiry_past = date.today() - timedelta(days=1)
key_expired = generate_key(STATION_ID, machine_id, expiry_past, is_trial=False)
print(f"expired key: {key_expired[:60]}...")

# ── 4. Clean slate — remove any existing license.key ─────────────────────────
section("Clean slate — remove license.key if present")
if KEY_FILE.exists():
    KEY_FILE.unlink()
    print("Removed existing license.key")
else:
    print("No existing license.key — good")

# ── 5. check_license() before activation ─────────────────────────────────────
section("check_license() — before activation")
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "not_activated", f"FAIL: expected not_activated, got {info['status']}"
assert info["valid"] == False
print("PASS: status=not_activated, valid=False ✓")

# ── 6. Activate the valid trial key ──────────────────────────────────────────
section("activate_key() — valid trial key")
result = activate_key(key_trial, STATION_ID)
print(f"activated  : {result}")
assert result == True, "FAIL: expected True"
assert KEY_FILE.exists(), "FAIL: license.key was not created"
print("PASS: activated=True, license.key created ✓")

# ── 7. check_license() after activation (30-day trial) ───────────────────────
section("check_license() — after activation (30-day trial)")
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["valid"] == True
assert info["status"] == "trial", f"FAIL: expected trial, got {info['status']}"
assert info["is_trial"] == True
assert info["days_remaining"] >= 29
print(f"PASS: status={info['status']}, days_remaining={info['days_remaining']} ✓")

# ── 8. get_expiry_date() ─────────────────────────────────────────────────────
section("get_expiry_date()")
expiry = get_expiry_date(STATION_ID)
print(f"expiry     : {expiry}")
assert expiry == expiry_30, f"FAIL: expected {expiry_30}, got {expiry}"
print("PASS: expiry date matches ✓")

# ── 9. is_expiring_soon() ────────────────────────────────────────────────────
section("is_expiring_soon(days=30) — should be True (30-day key)")
result = is_expiring_soon(STATION_ID, days=30)
print(f"result     : {result}")
assert result == True
print("PASS: is_expiring_soon(30)=True ✓")

section("is_expiring_soon(days=3) — should be False (30-day key)")
result = is_expiring_soon(STATION_ID, days=3)
print(f"result     : {result}")
assert result == False
print("PASS: is_expiring_soon(3)=False ✓")

# ── 10. Activate an expired key ───────────────────────────────────────────────
section("activate_key() — expired key (should return False)")
KEY_FILE.unlink()  # reset
result = activate_key(key_expired, STATION_ID)
print(f"activated  : {result}")
assert result == False, "FAIL: expired key should not activate"
assert not KEY_FILE.exists(), "FAIL: license.key should not be written for expired key"
print("PASS: expired key rejected, license.key not written ✓")

# ── 11. check_license() after failed activation ───────────────────────────────
section("check_license() — after failed activation attempt")
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "not_activated"
print("PASS: still not_activated after rejected key ✓")

# ── 12. Tampered license.key ─────────────────────────────────────────────────
section("check_license() — tampered license.key")
# First activate a good key
activate_key(key_trial, STATION_ID)
# Then corrupt the file
KEY_FILE.write_text("this-is-not-a-valid-encrypted-payload", encoding="utf-8")
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "tampered"
assert info["valid"] == False
print("PASS: tampered file detected ✓")

# ── 13. Wrong station_id in activate ─────────────────────────────────────────
section("activate_key() — wrong station_id")
KEY_FILE.unlink(missing_ok=True)
result = activate_key(key_trial, "wrong_station")
print(f"activated  : {result}")
assert result == False
print("PASS: wrong station_id rejected ✓")

# ── 14. 3-day warning key (trial_expiring) ────────────────────────────────────
section("check_license() — trial key with 2 days remaining (trial_expiring)")
KEY_FILE.unlink(missing_ok=True)
expiry_soon = date.today() + timedelta(days=2)
key_soon = generate_key(STATION_ID, machine_id, expiry_soon, is_trial=True)
activate_key(key_soon, STATION_ID)
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "trial_expiring"
assert info["valid"] == True
print("PASS: status=trial_expiring ✓")

# ── 15. 5-day non-trial key (expiring_soon) ───────────────────────────────────
section("check_license() — non-trial key with 5 days remaining (expiring_soon)")
KEY_FILE.unlink(missing_ok=True)
expiry_5 = date.today() + timedelta(days=5)
key_5 = generate_key(STATION_ID, machine_id, expiry_5, is_trial=False)
activate_key(key_5, STATION_ID)
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "expiring_soon"
assert info["valid"] == True
print("PASS: status=expiring_soon ✓")

# ── 16. 20-day non-trial key (renew_reminder) ─────────────────────────────────
section("check_license() — non-trial key with 20 days remaining (renew_reminder)")
KEY_FILE.unlink(missing_ok=True)
expiry_20 = date.today() + timedelta(days=20)
key_20 = generate_key(STATION_ID, machine_id, expiry_20, is_trial=False)
activate_key(key_20, STATION_ID)
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "renew_reminder"
assert info["valid"] == True
print("PASS: status=renew_reminder ✓")

# ── 17. 60-day non-trial key (active) ────────────────────────────────────────
section("check_license() — non-trial key with 60 days remaining (active)")
KEY_FILE.unlink(missing_ok=True)
expiry_60 = date.today() + timedelta(days=60)
key_60 = generate_key(STATION_ID, machine_id, expiry_60, is_trial=False)
activate_key(key_60, STATION_ID)
info = check_license(STATION_ID)
print(f"result     : {info}")
assert info["status"] == "active"
assert info["valid"] == True
assert info["is_trial"] == False
print("PASS: status=active ✓")

# ── Cleanup: restore a clean 30-day trial key ─────────────────────────────────
section("CLEANUP — restore 30-day trial key as final state")
KEY_FILE.unlink(missing_ok=True)
activate_key(key_trial, STATION_ID)
info = check_license(STATION_ID)
print(f"Final state: {info}")

print(f"\n{'='*60}")
print("ALL 17 TESTS PASSED ✅")
print(f"{'='*60}\n")