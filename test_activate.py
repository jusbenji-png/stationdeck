# test_activate.py
# Run from project root: python test_activate.py
# Paste your generated key when prompted.

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.license import activate_key, check_license, KEY_FILE

STATION_ID = "te_rwizi"

print("\n=== StationDeck — Activation Test ===\n")
print("Paste your license key and press Enter:")
key = input("> ").strip()

# Remove existing license.key so we get a clean activation
if KEY_FILE.exists():
    KEY_FILE.unlink()
    print("(Cleared existing license.key for clean test)")

print("\nActivating...")
result = activate_key(key, STATION_ID)

if result:
    info = check_license(STATION_ID)
    print("\n✅ ACTIVATION SUCCESSFUL")
    print(f"   Status         : {info['status']}")
    print(f"   Expiry         : {info['expiry']}")
    print(f"   Days remaining : {info['days_remaining']}")
    print(f"   Is trial       : {info['is_trial']}")
    print(f"   Valid          : {info['valid']}")
else:
    print("\n❌ ACTIVATION FAILED")
    print("   Check that the key was copied in full with no extra spaces.")