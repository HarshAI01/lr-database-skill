#!/usr/bin/env python3
"""Get last serial number from LR Matrix.xlsx"""

import sys
from pathlib import Path
from openpyxl import load_workbook

def main():
    if len(sys.argv) < 2:
        print("0")
        return

    lr_matrix_path = Path(sys.argv[1])
    if not lr_matrix_path.exists():
        print("0")
        return

    try:
        wb = load_workbook(lr_matrix_path, read_only=True)
        ws = wb.active
        max_serial = 0
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
            val = row[0]
            if isinstance(val, int) and val > max_serial:
                max_serial = val
        print(max_serial)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        print("0")

if __name__ == "__main__":
    main()