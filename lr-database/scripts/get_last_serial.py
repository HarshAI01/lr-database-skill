#!/usr/bin/env python3
"""Read the last continuous serial number from Outputs/LR Matrix.xlsx safely.

Import-safe: ``get_last_serial(path)`` returns an int or raises SerialError.
The CLI prints the integer on stdout and exits 1 with a message on stderr.
"""

from __future__ import annotations

import sys
from pathlib import Path


class SerialError(RuntimeError):
    """Raised when the serial column cannot be read safely."""


def get_last_serial(lr_matrix_path: Path | str) -> int:
    """Return the highest S. No. in the LR matrix, or 0 when empty/missing.

    Non-numeric values in the serial column are treated as corruption and raise
    SerialError rather than being silently ignored.
    """
    from openpyxl import load_workbook

    path = Path(lr_matrix_path)
    if not path.exists():
        return 0

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - actionable context
        raise SerialError(
            f"Could not open {path.name}: {exc}. Close it in Excel and retry."
        ) from exc

    try:
        worksheet = workbook.active
        maximum = 0
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True),
            start=2,
        ):
            value = row[0]
            if value is None or (isinstance(value, str) and not value.strip()):
                continue
            if isinstance(value, bool):
                raise SerialError(
                    f"{path.name} row {row_number}: S. No. is a boolean, not a number."
                )
            if isinstance(value, int):
                serial = value
            elif isinstance(value, float) and value.is_integer():
                serial = int(value)
            elif isinstance(value, str) and value.strip().isdigit():
                serial = int(value.strip())
            else:
                raise SerialError(
                    f"{path.name} row {row_number}: S. No. {value!r} is not an integer. "
                    "Repair the serial column before assembling."
                )
            if serial < 0:
                raise SerialError(
                    f"{path.name} row {row_number}: S. No. {serial} is negative."
                )
            maximum = max(maximum, serial)
        return maximum
    finally:
        workbook.close()


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 1:
        print("Usage: python get_last_serial.py <path-to-LR Matrix.xlsx>", file=sys.stderr)
        return 2
    try:
        print(get_last_serial(Path(argv[0])))
    except SerialError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
