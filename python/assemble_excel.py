#!/usr/bin/env python3
"""
assemble_excel.py — Excel assembly with SOP-compliant formatting
Preserves: 10pt font, wrapped text, top-aligned, thin borders, auto-filter extension
"""

import sys
import json
import os
from pathlib import Path
from openpyxl import Workbook, load_workbook
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

def get_data_style():
    return {
        "font": DATA_FONT,
        "alignment": DATA_ALIGNMENT,
        "border": THIN_BORDER,
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

def extend_autofilter(ws, new_last_row):
    """Extend existing auto-filter to new last row"""
    if ws.auto_filter.ref:
        # Parse existing ref like "A1:O300"
        parts = ws.auto_filter.ref.split(":")
        if len(parts) == 2:
            start_cell = parts[0]
            # Keep same start, extend end
            col_letter = get_column_letter(ws.max_column)
            ws.auto_filter.ref = f"{start_cell}:{col_letter}{new_last_row}"
    else:
        # No filter yet, create one
        col_letter = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{col_letter}{new_last_row}"

def load_or_create_workbook(path, columns, sheet_name="Sheet1"):
    """Load existing workbook or create new with headers"""
    path = Path(path)
    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
        # Verify headers match
        existing_headers = [ws.cell(row=1, column=c).value for c in range(1, len(columns) + 1)]
        if existing_headers != columns:
            print(f"WARNING: Header mismatch in {path}", file=sys.stderr)
        return wb, ws
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        # Write headers
        for col_idx, col_name in enumerate(columns, 1):
            ws.cell(row=1, column=col_idx, value=col_name)
        apply_header_style(ws, len(columns))
        return wb, ws

def append_metadata_rows(ws, extractions, metadata_columns, start_row):
    """Append metadata records to worksheet"""
    for i, ex in enumerate(extractions):
        row = start_row + i
        meta = ex["metadata"]
        for col_idx, col_name in enumerate(metadata_columns, 1):
            value = meta.get(col_name, "")
            # Handle special types
            if isinstance(value, bool):
                value = "TRUE" if value else "FALSE"
            elif value is None:
                value = ""
            ws.cell(row=row, column=col_idx, value=value)
    return start_row + len(extractions) - 1

def append_lr_rows(ws, extractions, lr_columns, start_row):
    """Append LR matrix records to worksheet"""
    for i, ex in enumerate(extractions):
        row = start_row + i
        lr = ex["lrMatrix"]
        for col_idx, col_name in enumerate(lr_columns, 1):
            value = lr.get(col_name, "")
            if isinstance(value, bool):
                value = "TRUE" if value else "FALSE"
            elif value is None:
                value = ""
            ws.cell(row=row, column=col_idx, value=value)
    return start_row + len(extractions) - 1

def get_last_serial(lr_matrix_path):
    """Get last serial number from LR Matrix"""
    path = Path(lr_matrix_path)
    if not path.exists():
        return 0
    wb = load_workbook(path, read_only=True)
    ws = wb.active
    max_serial = 0
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        val = row[0]
        if isinstance(val, int) and val > max_serial:
            max_serial = val
    return max_serial

def main():
    if len(sys.argv) < 2:
        print("Usage: python assemble_excel.py <json_input>", file=sys.stderr)
        sys.exit(1)

    # Parse input JSON (passed as single argument)
    input_json = sys.argv[1]
    data = json.loads(input_json)

    extractions = data["extractions"]
    metadata_batch_file = data["metadataBatchFile"]
    lr_matrix_file = data["lrMatrixFile"]
    start_serial = data["startSerial"]
    metadata_columns = data["metadataColumns"]
    lr_matrix_columns = data["lrMatrixColumns"]

    # Ensure output directories exist
    Path(metadata_batch_file).parent.mkdir(parents=True, exist_ok=True)
    Path(lr_matrix_file).parent.mkdir(parents=True, exist_ok=True)

    # ===== METADATA BATCH FILE (new file per batch) =====
    meta_wb, meta_ws = load_or_create_workbook(metadata_batch_file, metadata_columns, "Metadata")
    # Batch file starts fresh each time (per SOP)
    # Write headers if new
    if meta_ws.max_row == 1 and meta_ws.max_column == 1 and meta_ws.cell(1,1).value is None:
        for col_idx, col_name in enumerate(metadata_columns, 1):
            meta_ws.cell(row=1, column=col_idx, value=col_name)
        apply_header_style(meta_ws, len(metadata_columns))
    
    last_meta_row = meta_ws.max_row
    if last_meta_row == 1 and meta_ws.cell(1,1).value is not None:
        last_meta_row = 1
    
    new_last_meta_row = append_metadata_rows(meta_ws, extractions, metadata_columns, last_meta_row + 1)
    apply_data_style(meta_ws, last_meta_row + 1, new_last_meta_row, len(metadata_columns))
    extend_autofilter(meta_ws, new_last_meta_row)
    meta_wb.save(metadata_batch_file)

    # ===== LR MATRIX FILE (cumulative, append) =====
    lr_wb, lr_ws = load_or_create_workbook(lr_matrix_file, lr_matrix_columns, "LR Matrix")
    
    last_lr_row = lr_ws.max_row
    if last_lr_row == 1 and lr_ws.cell(1,1).value is not None:
        last_lr_row = 1
    
    new_last_lr_row = append_lr_rows(lr_ws, extractions, lr_matrix_columns, last_lr_row + 1)
    apply_data_style(lr_ws, last_lr_row + 1, new_last_lr_row, len(lr_matrix_columns))
    extend_autofilter(lr_ws, new_last_lr_row)
    lr_wb.save(lr_matrix_file)

    print(f"OK: metadata batch -> {metadata_batch_file} ({len(extractions)} rows)")
    print(f"OK: LR matrix -> {lr_matrix_file} (rows {start_serial}-{start_serial + len(extractions) - 1})")

if __name__ == "__main__":
    main()