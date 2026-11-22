#!/usr/bin/env python3
"""Generate SOP-compliant Excel templates for the LR Database skill.

Creates, if missing:
  - Outputs/LR Matrix.xlsx            (20 columns, cumulative)
  - Outputs/Meta Data Filled.xlsx     (87 columns, frozen baseline)
  - Outputs/metadata batch YYYY-MM-DD.xlsx (87 columns, per-day batch)

The script is import-safe: pipeline.py calls ensure_templates(), while the CLI
prints a JSON status object.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# SOP constants
# ---------------------------------------------------------------------------

HEADER_FILL_HEX = "1F4E78"
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
HEADER_FILL = PatternFill(
    start_color=f"FF{HEADER_FILL_HEX}", end_color=f"FF{HEADER_FILL_HEX}", fill_type="solid"
)
HEADER_ALIGNMENT = Alignment(wrap_text=True, vertical="center", horizontal="center")
DATA_FONT = Font(name="Calibri", size=10)
DATA_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

METADATA_COLUMNS = [
    "File Name", "Item Type", "Publication Year", "Author", "Title",
    "Publication Title", "ISBN", "ISSN", "DOI", "Url", "Abstract Note",
    "Date", "Date Added", "Date Modified", "Access Date", "Pages",
    "Num Pages", "Issue", "Volume", "Number Of Volumes", "Journal Abbreviation",
    "Short Title", "Series", "Series Number", "Series Text", "Series Title",
    "Publisher", "Place", "Language", "Rights", "Type", "Archive",
    "Archive Location", "Library Catalog", "Call Number", "Extra", "Notes",
    "File Attachments", "Link Attachments", "Manual Tags", "Automatic Tags",
    "Editor", "Series Editor", "Translator", "Contributor", "Attorney Agent",
    "Book Author", "Cast Member", "Commenter", "Composer", "Cosponsor",
    "Counsel", "Interviewer", "Producer", "Recipient", "Reviewed Author",
    "Scriptwriter", "Words By", "Guest", "Number", "Edition", "Running Time",
    "Scale", "Medium", "Artwork Size", "Filing Date", "Application Number",
    "Assignee", "Issuing Authority", "Country", "Meeting Name",
    "Conference Name", "Court", "References", "Reporter", "Legal Status",
    "Priority Numbers", "Programming Language", "Version", "System", "Code",
    "Code Number", "Section", "Session", "Committee", "History",
    "Legislative Body",
]

LR_MATRIX_COLUMNS = [
    "S. No.", "File Name", "Title", "Author", "Year", "Journal", "DOI",
    "Objective", "Methodology", "Key Finding", "Discussion", "Limitation",
    "Theoretical Framework", "Tag 1 (Method)", "Tag 2 (Type of Research)",
    "Independent Variable", "Dependent Variable", "Other Variable",
    "Decision (For Me)", "Reason (For Me)",
]

METADATA_WIDTHS = {
    "File Name": 40, "Item Type": 14, "Publication Year": 16, "Author": 35,
    "Title": 50, "Publication Title": 35, "ISBN": 18, "ISSN": 12, "DOI": 30,
    "Url": 40, "Abstract Note": 60, "Date": 14, "Date Added": 22,
    "Date Modified": 22, "Access Date": 14, "Pages": 14, "Num Pages": 10,
    "Issue": 8, "Volume": 10, "Number Of Volumes": 16,
    "Journal Abbreviation": 22, "Short Title": 30, "Series": 14,
    "Series Number": 14, "Series Text": 14, "Series Title": 20,
    "Publisher": 25, "Place": 18, "Language": 10, "Rights": 14, "Type": 12,
    "Archive": 14, "Archive Location": 18, "Library Catalog": 18,
    "Call Number": 14, "Extra": 30, "Notes": 30, "File Attachments": 50,
    "Link Attachments": 30, "Manual Tags": 18, "Automatic Tags": 18,
    "Editor": 20, "Series Editor": 16, "Translator": 16, "Contributor": 16,
    "Attorney Agent": 16, "Book Author": 16, "Cast Member": 14,
    "Commenter": 14, "Composer": 14, "Cosponsor": 14, "Counsel": 12,
    "Interviewer": 14, "Producer": 14, "Recipient": 14, "Reviewed Author": 16,
    "Scriptwriter": 14, "Words By": 12, "Guest": 10, "Number": 10,
    "Edition": 10, "Running Time": 14, "Scale": 10, "Medium": 12,
    "Artwork Size": 14, "Filing Date": 14, "Application Number": 18,
    "Assignee": 16, "Issuing Authority": 20, "Country": 12,
    "Meeting Name": 20, "Conference Name": 20, "Court": 12, "References": 14,
    "Reporter": 12, "Legal Status": 14, "Priority Numbers": 16,
    "Programming Language": 20, "Version": 10, "System": 10, "Code": 10,
    "Code Number": 12, "Section": 10, "Session": 10, "Committee": 14,
    "History": 10, "Legislative Body": 18,
}

LR_WIDTHS = {
    "S. No.": 8, "File Name": 40, "Title": 50, "Author": 35, "Year": 10,
    "Journal": 35, "DOI": 40, "Objective": 55, "Methodology": 60,
    "Key Finding": 60, "Discussion": 55, "Limitation": 45,
    "Theoretical Framework": 30, "Tag 1 (Method)": 20,
    "Tag 2 (Type of Research)": 22, "Independent Variable": 25,
    "Dependent Variable": 25, "Other Variable": 25,
    "Decision (For Me)": 16, "Reason (For Me)": 30,
}


class TemplateError(RuntimeError):
    """Raised when a template cannot be created or validated."""


# ---------------------------------------------------------------------------
# Atomic save
# ---------------------------------------------------------------------------


def save_workbook_atomic(workbook: Workbook, target: Path) -> None:
    """Save *workbook* to *target* atomically (temp file + os.replace)."""
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=".tmp_", suffix=target.suffix or ".xlsx", dir=str(target.parent)
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        workbook.save(tmp_path)
        os.replace(tmp_path, target)
    except Exception as exc:  # noqa: BLE001 - re-raised with context
        raise TemplateError(
            f"Could not write {target.name}: {exc}. Close the file in Excel and retry."
        ) from exc
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Workbook builders
# ---------------------------------------------------------------------------


def _apply_header(ws, columns: list[str]) -> None:
    for index, _column in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=index)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 30


def _apply_blank_data_row(ws, columns: list[str]) -> None:
    for index in range(1, len(columns) + 1):
        cell = ws.cell(row=2, column=index, value="")
        cell.font = DATA_FONT
        cell.alignment = DATA_ALIGNMENT
        cell.border = THIN_BORDER


def _set_widths(ws, columns: list[str], widths: dict[str, int]) -> None:
    for index, column in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(index)].width = min(
            widths.get(column, 14), 60
        )


def create_workbook(
    columns: list[str], widths: dict[str, int], sheet_name: str
) -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name

    for index, column in enumerate(columns, start=1):
        worksheet.cell(row=1, column=index, value=column)

    _apply_header(worksheet, columns)
    _set_widths(worksheet, columns, widths)
    _apply_blank_data_row(worksheet, columns)
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}2"
    return workbook


def _validate_columns(actual: list[str], expected: list[str], label: str) -> None:
    if actual != expected:
        raise TemplateError(
            f"{label} header does not match the SOP. "
            f"Expected {len(expected)} columns, found {len(actual)}. "
            "Restore the header from references/sop-rules.md before continuing."
        )


def ensure_templates(
    outputs_dir: Path, batch_date: date | str | None = None
) -> dict:
    """Create any missing workbook under *outputs_dir*.

    Returns a report with ``created``, ``existing``, and ``paths`` keys.
    """
    outputs_dir = Path(outputs_dir).resolve()
    outputs_dir.mkdir(parents=True, exist_ok=True)

    if batch_date is None:
        batch_day = date.today()
    elif isinstance(batch_date, str):
        try:
            batch_day = datetime.strptime(batch_date, "%Y-%m-%d").date()
        except ValueError as exc:
            raise TemplateError(
                f"Invalid batch date {batch_date!r}; expected YYYY-MM-DD."
            ) from exc
    elif isinstance(batch_date, date):
        batch_day = batch_date
    else:
        raise TemplateError("batch_date must be a date, a YYYY-MM-DD string, or None.")

    specs = [
        (
            outputs_dir / "LR Matrix.xlsx",
            lambda: create_workbook(LR_MATRIX_COLUMNS, LR_WIDTHS, "LR Matrix"),
        ),
        (
            outputs_dir / "Meta Data Filled.xlsx",
            lambda: create_workbook(METADATA_COLUMNS, METADATA_WIDTHS, "Metadata"),
        ),
        (
            outputs_dir / f"metadata batch {batch_day.isoformat()}.xlsx",
            lambda: create_workbook(METADATA_COLUMNS, METADATA_WIDTHS, "Metadata"),
        ),
    ]

    created: list[str] = []
    existing: list[str] = []
    for path, factory in specs:
        if path.exists():
            existing.append(path.name)
            continue
        workbook = factory()
        save_workbook_atomic(workbook, path)
        created.append(path.name)

    return {
        "outputsDir": str(outputs_dir),
        "batchDate": batch_day.isoformat(),
        "created": created,
        "existing": existing,
        "paths": {
            "lrMatrix": str(outputs_dir / "LR Matrix.xlsx"),
            "metaDataFilled": str(outputs_dir / "Meta Data Filled.xlsx"),
            "metadataBatch": str(
                outputs_dir / f"metadata batch {batch_day.isoformat()}.xlsx"
            ),
        },
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Create the LR Database Excel templates if missing."
    )
    parser.add_argument("output_dir", help="Path to the project's Outputs/ folder.")
    parser.add_argument(
        "--date", dest="batch_date", default=None, help="Batch date YYYY-MM-DD."
    )
    args = parser.parse_args(argv)

    try:
        report = ensure_templates(Path(args.output_dir), args.batch_date)
    except TemplateError as exc:
        print(
            json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=True),
            flush=True,
        )
        return 1

    print(json.dumps({"status": "ok", **report}, ensure_ascii=True), flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
