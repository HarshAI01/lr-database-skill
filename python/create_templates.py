#!/usr/bin/env python3
"""
create_templates.py — Generate SOP-compliant template Excel files programmatically
No external Reference & Template folder needed. Creates:
1. LR Matrix.xlsx (20 columns, formatted headers)
2. Meta Data Filled.xlsx (87 columns, formatted headers)
3. metadata batch YYYY-MM-DD.xlsx (87 columns, formatted headers)
"""

import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# SOP formatting constants
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_ALIGNMENT = Alignment(wrap_text=True, vertical="center", horizontal="center")
DATA_FONT = Font(name="Calibri", size=10)
DATA_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Column definitions from SOP §5A and §5B
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
    "Legislative Body"
]

LR_MATRIX_COLUMNS = [
    "S. No.", "File Name", "Title", "Author", "Year", "Journal", "DOI",
    "Objective", "Methodology", "Key Finding", "Discussion", "Limitation",
    "Theoretical Framework", "Tag 1 (Method)", "Tag 2 (Type of Research)",
    "Independent Variable", "Dependent Variable", "Other Variable",
    "Decision (For Me)", "Reason (For Me)"
]

# Suggested column widths for readability
METADATA_WIDTHS = {
    "File Name": 40, "Item Type": 14, "Publication Year": 16, "Author": 35, "Title": 50,
    "Publication Title": 35, "ISBN": 18, "ISSN": 12, "DOI": 30, "Url": 40,
    "Abstract Note": 60, "Date": 14, "Date Added": 22, "Date Modified": 22, "Access Date": 14,
    "Pages": 14, "Num Pages": 10, "Issue": 8, "Volume": 10, "Number Of Volumes": 16,
    "Journal Abbreviation": 22, "Short Title": 30, "Series": 14, "Series Number": 14,
    "Series Text": 14, "Series Title": 20, "Publisher": 25, "Place": 18, "Language": 10,
    "Rights": 14, "Type": 12, "Archive": 14, "Archive Location": 18, "Library Catalog": 18,
    "Call Number": 14, "Extra": 30, "Notes": 30, "File Attachments": 50, "Link Attachments": 30,
    "Manual Tags": 18, "Automatic Tags": 18, "Editor": 20, "Series Editor": 16,
    "Translator": 16, "Contributor": 16, "Attorney Agent": 16, "Book Author": 16,
    "Cast Member": 14, "Commenter": 14, "Composer": 14, "Cosponsor": 14, "Counsel": 12,
    "Interviewer": 14, "Producer": 14, "Recipient": 14, "Reviewed Author": 16,
    "Scriptwriter": 14, "Words By": 12, "Guest": 10, "Number": 10, "Edition": 10,
    "Running Time": 14, "Scale": 10, "Medium": 12, "Artwork Size": 14, "Filing Date": 14,
    "Application Number": 18, "Assignee": 16, "Issuing Authority": 20, "Country": 12,
    "Meeting Name": 20, "Conference Name": 20, "Court": 12, "References": 14,
    "Reporter": 12, "Legal Status": 14, "Priority Numbers": 16, "Programming Language": 20,
    "Version": 10, "System": 10, "Code": 10, "Code Number": 12, "Section": 10,
    "Session": 10, "Committee": 14, "History": 10, "Legislative Body": 18,
}

LR_WIDTHS = {
    "S. No.": 8, "File Name": 40, "Title": 50, "Author": 35, "Year": 10,
    "Journal": 35, "DOI": 40, "Objective": 55, "Methodology": 60,
    "Key Finding": 60, "Discussion": 55, "Limitation": 45,
    "Theoretical Framework": 30, "Tag 1 (Method)": 20, "Tag 2 (Type of Research)": 22,
    "Independent Variable": 25, "Dependent Variable": 25, "Other Variable": 25,
    "Decision (For Me)": 16, "Reason (For Me)": 30,
}

def apply_header_style(ws, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER

def apply_data_style(ws, start_row, end_row, num_cols):
    for row in range(start_row, end_row + 1):
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = DATA_FONT
            cell.alignment = DATA_ALIGNMENT
            cell.border = THIN_BORDER

def set_column_widths(ws, columns, width_map):
    for idx, col_name in enumerate(columns, 1):
        width = width_map.get(col_name, 14)
        ws.column_dimensions[get_column_letter(idx)].width = min(width, 60)

def create_workbook(columns, width_map, sheet_name="Sheet1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    # Write headers
    for col_idx, col_name in enumerate(columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    apply_header_style(ws, len(columns))
    set_column_widths(ws, columns, width_map)
    # Add one empty data row with formatting (so filter works)
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=2, column=col_idx, value="")
        cell.font = DATA_FONT
        cell.alignment = DATA_ALIGNMENT
        cell.border = THIN_BORDER
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}2"
    return wb, ws

def main():
    if len(sys.argv) < 2:
        print("Usage: python create_templates.py <output_dir>")
        sys.exit(1)

    out_dir = Path(sys.argv[1])
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. LR Matrix.xlsx (cumulative)
    lr_path = out_dir / "LR Matrix.xlsx"
    if not lr_path.exists():
        wb, ws = create_workbook(LR_MATRIX_COLUMNS, LR_WIDTHS, "LR Matrix")
        wb.save(lr_path)
        print(f"Created: {lr_path}")
    else:
        print(f"Exists (skipped): {lr_path}")

    # 2. Meta Data Filled.xlsx (baseline, frozen)
    meta0_path = out_dir / "Meta Data Filled.xlsx"
    if not meta0_path.exists():
        wb, ws = create_workbook(METADATA_COLUMNS, METADATA_WIDTHS, "Metadata")
        wb.save(meta0_path)
        print(f"Created: {meta0_path}")
    else:
        print(f"Exists (skipped): {meta0_path}")

    # 3. Example batch file (optional)
    from datetime import date
    batch_path = out_dir / f"metadata batch {date.today().isoformat()}.xlsx"
    if not batch_path.exists():
        wb, ws = create_workbook(METADATA_COLUMNS, METADATA_WIDTHS, "Metadata")
        wb.save(batch_path)
        print(f"Created: {batch_path}")
    else:
        print(f"Exists (skipped): {batch_path}")

    print("\n[OK] Templates ready. No Reference & Template folder needed.")

if __name__ == "__main__":
    main()