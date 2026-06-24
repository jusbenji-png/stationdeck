# patch_test_ocr_v2.py — final coordinate fix
# Drop in C:\Users\LENOVO\stationdeck\ and run: python patch_test_ocr_v2.py

import re

with open("test_ocr.py", "r") as f:
    content = f.read()

# Find and replace the pms1_coords dict block regardless of version
OLD = re.search(
    r'"pms1_opening_dip".*?"pms1_selling_price".*?\),',
    content, re.DOTALL
)

if not OLD:
    print("ERROR: Could not find pms1 coordinate block in test_ocr.py")
    print("Lines containing pms1_:")
    for i, line in enumerate(content.split('\n'), 1):
        if 'pms1_' in line:
            print(f"  L{i}: {line.rstrip()}")
else:
    NEW_BLOCK = '''"pms1_opening_dip":      ( 280, 565,  700, 657),
    "pms1_closing_dip":      ( 750, 565, 1170, 657),
    "pms1_return_tank":      (1220, 565, 1900, 657),
    "pms1_updf_receipt":     ( 280, 742,  700, 832),
    "pms1_updf_consumption": ( 750, 742, 1170, 832),
    "pms1_cost_price":       ( 280, 800,  700, 900),
    "pms1_selling_price":    ( 600, 800, 1920, 900),'''

    content = content[:OLD.start()] + NEW_BLOCK + content[OLD.end():]

    # Also update expected values for this PMS Day shift photo
    content = re.sub(
        r'"pms1_opening_dip"\s*:\s*"[\d]+"',
        '"pms1_opening_dip": "315"',
        content
    )
    content = re.sub(
        r'"pms1_closing_dip"\s*:\s*"[\d]+"',
        '"pms1_closing_dip": "84"',
        content
    )
    content = re.sub(
        r'"pms1_return_tank"\s*:\s*"[\d]+"',
        '"pms1_return_tank": "0"',
        content
    )
    content = re.sub(
        r'"pms1_cost_price"\s*:\s*"[\d]+"',
        '"pms1_cost_price": "6467"',
        content
    )
    content = re.sub(
        r'"pms1_selling_price"\s*:\s*"[\d]+"',
        '"pms1_selling_price": "6550"',
        content
    )

    with open("test_ocr.py", "w") as f:
        f.write(content)
    print("test_ocr.py patched with final coordinates.")
    print()
    print("Final PMS1 coordinates:")
    print("  opening_dip:      (280,  565, 700,  657)  <- CONFIRMED CORRECT")
    print("  closing_dip:      (750,  565, 1170, 657)  <- CONFIRMED CORRECT")
    print("  return_tank:      (1220, 565, 1900, 657)  <- CONFIRMED CORRECT")
    print("  updf_receipt:     (280,  742, 700,  832)  <- +40px from v5")
    print("  updf_consumption: (750,  742, 1170, 832)  <- +40px from v5")
    print("  cost_price:       (280,  800, 700,  900)  <- CONFIRMED CORRECT")
    print("  selling_price:    (600,  800, 1920, 900)  <- wider x range")