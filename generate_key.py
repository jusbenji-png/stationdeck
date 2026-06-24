# generate_key.py
# Benjamin's admin tool for generating StationDeck license keys.
# Run from project root: python generate_key.py
#
# USE CASES:
#   1. Testing on your own machine during development.
#   2. Machine replacement: station WhatsApps their new machine_id,
#      you run this, send back the new key via WhatsApp. Done in 60 seconds.
#
# PHASE 15 NOTE:
#   When Flutterwave is integrated, the payment webhook will call
#   generate_key() from src/license.py directly — this script
#   will only be used for machine replacements from that point on.

import sys
from pathlib import Path
from datetime import date, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.license import generate_key, get_machine_id

SEP = "=" * 60

def ask(prompt, default=None):
    suffix = f" [{default}]" if default else ""
    val = input(f"{prompt}{suffix}: ").strip()
    return val if val else default

def main():
    print(f"\n{SEP}")
    print("  StationDeck — License Key Generator")
    print(f"{SEP}\n")

    # ── Mode selection ────────────────────────────────────────────────────────
    print("Who is this key for?")
    print("  1. This machine (development / self-test)")
    print("  2. Another machine (machine replacement for a station)")
    mode = ask("Enter 1 or 2", "1")

    print()

    # ── Station ID ────────────────────────────────────────────────────────────
    station_id = ask("Station ID (e.g. te_rwizi)", "te_rwizi")

    # ── Machine ID ────────────────────────────────────────────────────────────
    if mode == "1":
        machine_id = get_machine_id(station_id)
        print(f"\nYour machine_id: {machine_id}")
    else:
        print("\nPaste the machine_id from the station's Settings page:")
        machine_id = ask("machine_id").strip()
        if len(machine_id) != 64:
            print(f"\nERROR: machine_id must be 64 characters. Got {len(machine_id)}.")
            sys.exit(1)

    # ── Trial or paid? ────────────────────────────────────────────────────────
    print("\nKey type:")
    print("  1. Trial key    (30 days, is_trial=True  — stations 1–5 only)")
    print("  2. Monthly key  (30 days, is_trial=False)")
    print("  3. Annual key   (365 days, is_trial=False)")
    print("  4. Custom       (you specify exact days)")
    key_type = ask("Enter 1, 2, 3, or 4", "2")

    if key_type == "1":
        days = 30
        is_trial = True
    elif key_type == "2":
        days = 30
        is_trial = False
    elif key_type == "3":
        days = 365
        is_trial = False
    else:
        days = int(ask("Number of days until expiry", "30"))
        trial_input = ask("Is this a trial key? (y/n)", "n").lower()
        is_trial = trial_input == "y"

    expiry_date = date.today() + timedelta(days=days)

    # ── Confirm before generating ─────────────────────────────────────────────
    print(f"\n{SEP}")
    print("  Key summary — please confirm:")
    print(f"{SEP}")
    print(f"  Station ID   : {station_id}")
    print(f"  Machine ID   : {machine_id[:20]}...{machine_id[-8:]}")
    print(f"  Is trial     : {is_trial}")
    print(f"  Expires      : {expiry_date}  ({days} days from today)")
    print(f"{SEP}")

    confirm = ask("\nGenerate this key? (y/n)", "y").lower()
    if confirm != "y":
        print("\nCancelled. No key generated.")
        sys.exit(0)

    # ── Generate ──────────────────────────────────────────────────────────────
    key = generate_key(station_id, machine_id, expiry_date, is_trial)

    print(f"\n{SEP}")
    print("  LICENSE KEY GENERATED")
    print(f"{SEP}\n")
    print(key)
    print(f"\n{SEP}")
    print(f"  Expiry : {expiry_date}")
    print(f"  Trial  : {is_trial}")
    print(f"  Copy the key above and send it to the station.")
    print(f"{SEP}\n")

if __name__ == "__main__":
    main()