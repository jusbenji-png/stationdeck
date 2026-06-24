"""
fix_dev_license.py
------------------
Regenerates the local license key for the te_rwizi dev station
with a 365-day expiry from today.

RUN FROM PROJECT ROOT:
  python fix_dev_license.py
"""

import sys
from pathlib import Path
from datetime import date, timedelta

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from src.license import get_machine_id, generate_key, activate_key

STATION_ID = "te_rwizi"

def main():
    print("=" * 50)
    print("StationDeck — Dev License Regenerator")
    print("=" * 50)

    machine_id  = get_machine_id(STATION_ID)
    print(f"Machine ID : {machine_id[:16]}...")

    expiry_date = date.today() + timedelta(days=365)
    print(f"Expiry     : {expiry_date}")

    new_key = generate_key(
        station_id  = STATION_ID,
        machine_id  = machine_id,
        expiry_date = expiry_date,
        is_trial    = False,
    )
    print(f"Key        : {new_key[:24]}...")

    success = activate_key(new_key, STATION_ID)

    if success:
        print("\n✅ License renewed successfully — 365 days from today.")
        print("   Restart StationDeck. The expiry banner will be gone.")
    else:
        print("\n❌ activate_key() returned False.")
        print("   Check that config/ is writable and the key was generated correctly.")

if __name__ == "__main__":
    main()