#!/usr/bin/env python3
"""Validate an extraction payload and write it into the SOP workbooks.

Responsibilities
----------------
* Strict Draft-07 validation against ``references/extraction-schema.json``.
* Normalisation of fields the writer owns (exact filename, File Attachments,
  dates, language).
* Atomic, audited writes to ``LR Matrix.xlsx`` and the per-day metadata batch.
* Rollback of workbooks and PDF moves if any step fails.
* ``Meta Data Filled.xlsx`` is never loaded or modified.

The module is import-safe; ``pipeline.py`` calls ``validate_extractions`` and
``assemble``.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

from jsonschema import Draft7Validator
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import create_templates as templates
from get_last_serial import SerialError, get_last_serial

SCHEMA_PATH = Path(__file__).resolve().parent.parent / "references" / "extraction-schema.json"

METADATA_COLUMNS = templates.METADATA_COLUMNS
LR_MATRIX_COLUMNS = templates.LR_MATRIX_COLUMNS

WORD_BANDS = {
    "Objective": (40, 70),
    "Methodology": (60, 100),
    "Key Finding": (60, 100),
    "Discussion": (40, 70),
    "Limitation": (30, 60),
}

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_MAX_CELL = 32_000


class AssemblyError(RuntimeError):
    """Base error for a refused or failed assembly."""


class ValidationError(AssemblyError):
    """Schema or semantic validation failed."""

    def __init__(self, issues: list[str]):
        self.issues = issues
        super().__init__("Extraction payload failed validation.")


# ---------------------------------------------------------------------------
# Cleaning and normalisation
# ---------------------------------------------------------------------------


def clean_text(value: Any) -> str:
    """Make a value safe for an Excel cell without changing its meaning."""
    if value is None:
        return ""
    if not isinstance(value, str):
        return str(value)
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = _ILLEGAL_XML.sub("", value)
    if len(value) > _MAX_CELL:
        value = value[:_MAX_CELL]
    return value


def _normalise_payload(payload: Any) -> list[dict]:
    if isinstance(payload, dict):
        for key in ("extractions", "records", "items"):
            if key in payload:
                payload = payload[key]
                break
        else:
            raise ValidationError(
                [
                    "Top level must be a JSON array of records, or an object "
                    "with an 'extractions' array."
                ]
            )

    if not isinstance(payload, list):
        raise ValidationError(
            [
                "Top level must be a JSON array of records, or an object with "
                "an 'extractions' array."
            ]
        )
    if not payload:
        raise ValidationError(["The extraction payload is empty."])

    now = datetime.now().astimezone().isoformat(timespec="seconds")
    today = date.today().isoformat()
    normalised: list[dict] = []
    issues: list[str] = []

    for index, record in enumerate(payload):
        where = f"$[{index}]"
        if not isinstance(record, dict):
            issues.append(f"{where}: each record must be a JSON object.")
            continue

        file_name = record.get("fileName")
        if not isinstance(file_name, str) or not file_name.strip():
            issues.append(f"{where}.fileName: exact PDF filename is required.")
            continue
        file_name = clean_text(file_name).strip()
        if Path(file_name).name != file_name or "/" in file_name or "\\" in file_name:
            issues.append(
                f"{where}.fileName: must be a basename only (no folders): {file_name!r}."
            )
            continue

        metadata = record.get("metadata")
        lr_matrix = record.get("lrMatrix")
        if not isinstance(metadata, dict):
            issues.append(f"{where}.metadata: object with all 87 keys is required.")
            continue
        if not isinstance(lr_matrix, dict):
            issues.append(f"{where}.lrMatrix: object with all 19 keys is required.")
            continue

        metadata = {
            key: clean_text(value) if isinstance(value, str) else value
            for key, value in metadata.items()
        }
        lr_matrix = {
            key: clean_text(value) if isinstance(value, str) else value
            for key, value in lr_matrix.items()
        }

        metadata["File Name"] = file_name
        lr_matrix["File Name"] = file_name
        metadata["File Attachments"] = f"Database/{file_name}"

        if not metadata.get("Date Added"):
            metadata["Date Added"] = now
        if not metadata.get("Date Modified"):
            metadata["Date Modified"] = now
        if not metadata.get("Access Date"):
            metadata["Access Date"] = today
        if not metadata.get("Language"):
            metadata["Language"] = "en"

        normalised.append(
            {"fileName": file_name, "metadata": metadata, "lrMatrix": lr_matrix}
        )

    if issues:
        raise ValidationError(issues)
    return normalised


def _schema_error_message(error) -> str:
    path = "$"
    for part in error.absolute_path:
        path += f"[{part}]" if isinstance(part, int) else f".{part}"

    message = error.message
    hint = ""
    validator = getattr(error, "validator", "")

    if "Item Type" in path:
        hint = " Allowed: journalArticle, report, preprint, conferencePaper, thesis."
    elif "Tag 2 (Type of Research)" in path:
        hint = (
            " Allowed: Quantitative, Qualitative, Mixed, Review-Theoretical, Reference."
        )
    elif "Decision (For Me)" in path:
        hint = " Allowed: empty string, Include, Exclude."
    elif ".DOI" in path:
        if "lrMatrix" in path:
            hint = " Use https://doi.org/10.xxxx, or an empty string."
        else:
            hint = " Use bare 10.xxxx, or an empty string."
    elif "Num Pages" in path or "Number Of Volumes" in path:
        hint = " Must be a JSON integer, not a quoted string."
    elif validator == "additionalProperties":
        hint = " Remove keys that are not in the SOP schema."
    elif validator == "required":
        hint = " Every SOP key is required; use \"\" for missing data."
    elif validator == "type" and "string" in message:
        hint = ' Every text field must be a JSON string; use "" when unknown.'
    elif validator == "minLength":
        hint = " This field may not be blank."

    return f"{path}: {message}{hint}"


def validate_extractions(payload: Any) -> tuple[list[dict], list[str]]:
    """Validate and normalise a payload. Returns (records, warnings).

    Raises ValidationError with a list of actionable issues on failure.
    """
    records = _normalise_payload(payload)

    try:
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise AssemblyError(
            f"Could not read the extraction schema at {SCHEMA_PATH}: {exc}"
        ) from exc

    validator = Draft7Validator(schema)
    errors = sorted(validator.iter_errors(records), key=lambda e: list(e.absolute_path))
    if errors:
        issues = [_schema_error_message(error) for error in errors[:40]]
        if len(errors) > 40:
            issues.append(f"... and {len(errors) - 40} more schema errors.")
        raise ValidationError(issues)

    # Serial numbers are assigned by the writer, never trusted from input.
    for record in records:
        record["lrMatrix"].pop("S. No.", None)

    blank_titles = [
        record["fileName"]
        for record in records
        if not str(record["metadata"].get("Title", "")).strip()
        or not str(record["lrMatrix"].get("Title", "")).strip()
    ]
    if blank_titles:
        raise ValidationError(
            [f"{name}: Title may not be blank in metadata or lrMatrix." for name in blank_titles]
        )

    file_names = [record["fileName"] for record in records]
    duplicates = sorted({name for name in file_names if file_names.count(name) > 1})
    if duplicates:
        raise ValidationError(
            [f"Duplicate fileName in payload: {name}" for name in duplicates]
        )

    warnings: list[str] = []
    for record in records:
        name = record["fileName"]
        lr = record["lrMatrix"]
        meta = record["metadata"]

        for field, (low, high) in WORD_BANDS.items():
            text = str(lr.get(field, "")).strip()
            count = len(text.split())
            if not text:
                warnings.append(
                    f"{name}: {field} is blank (SOP target {low}-{high} words)."
                )
            elif count < low or count > high:
                warnings.append(
                    f"{name}: {field} has {count} words (SOP target {low}-{high})."
                )

        meta_doi = str(meta.get("DOI", "")).strip()
        lr_doi = str(lr.get("DOI", "")).strip()
        if meta_doi and not lr_doi:
            warnings.append(f"{name}: LR DOI is blank while metadata DOI is set.")
        elif meta_doi and lr_doi and lr_doi != f"https://doi.org/{meta_doi}":
            warnings.append(
                f"{name}: LR DOI does not equal https://doi.org/{meta_doi}."
            )

    return records, warnings


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------


def _header_values(worksheet, count: int) -> list[Any]:
    return [worksheet.cell(row=1, column=column).value for column in range(1, count + 1)]


def _check_header(worksheet, columns: list[str], label: str) -> None:
    actual = _header_values(worksheet, len(columns))
    if actual != columns:
        raise AssemblyError(
            f"{label} header does not match the SOP. Expected {len(columns)} "
            f"columns in the documented order, found {len(actual)}. Restore the "
            "header before assembling."
        )
    extra = [
        worksheet.cell(row=1, column=column).value
        for column in range(len(columns) + 1, worksheet.max_column + 1)
    ]
    if any(value not in (None, "") for value in extra):
        raise AssemblyError(
            f"{label} has extra header columns beyond {len(columns)}. "
            "Remove them before assembling."
        )


def _read_rows(worksheet, columns: list[str]) -> list[tuple[int, dict]]:
    rows: list[tuple[int, dict]] = []
    for row_number in range(2, worksheet.max_row + 1):
        values = [
            worksheet.cell(row=row_number, column=column).value
            for column in range(1, len(columns) + 1)
        ]
        if all(value is None or (isinstance(value, str) and not value.strip()) for value in values):
            continue
        rows.append((row_number, dict(zip(columns, values))))
    return rows


def _last_data_row(worksheet) -> int:
    last = 1
    for row_number in range(2, worksheet.max_row + 1):
        values = [
            worksheet.cell(row=row_number, column=column).value
            for column in range(1, worksheet.max_column + 1)
        ]
        if any(value is not None and (not isinstance(value, str) or value.strip()) for value in values):
            last = row_number
    return last


def _append_rows(worksheet, columns: list[str], rows: list[dict]) -> tuple[int, int]:
    if not rows:
        return 0, 0

    start = max(_last_data_row(worksheet) + 1, 2)
    for offset, values in enumerate(rows):
        row_number = start + offset
        for column_number, column in enumerate(columns, start=1):
            value = values.get(column, "")
            if value is None:
                value = ""
            cell = worksheet.cell(row=row_number, column=column_number, value=value)
            cell.font = templates.DATA_FONT
            cell.alignment = templates.DATA_ALIGNMENT
            cell.border = templates.THIN_BORDER
            if isinstance(value, str) and value.startswith("="):
                # Never let PDF text become an Excel formula.
                cell.data_type = "s"

    last = start + len(rows) - 1
    worksheet.auto_filter.ref = (
        f"A1:{get_column_letter(len(columns))}{last}"
    )
    return start, last


def _serial_int(value: Any, where: str) -> int:
    if isinstance(value, bool) or value is None:
        raise AssemblyError(f"{where}: serial is missing or not an integer.")
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    raise AssemblyError(f"{where}: serial {value!r} is not an integer.")


def _audit_existing_lr(worksheet) -> list[tuple[int, dict]]:
    _check_header(worksheet, LR_MATRIX_COLUMNS, "LR Matrix.xlsx")

    rows = _read_rows(worksheet, LR_MATRIX_COLUMNS)
    serials: list[int] = []
    names: list[str] = []
    for row_number, row in rows:
        serial = _serial_int(row.get("S. No."), f"LR Matrix row {row_number}")
        if serial < 1:
            raise AssemblyError(
                f"LR Matrix row {row_number}: S. No. must be >= 1."
            )
        serials.append(serial)
        name = str(row.get("File Name") or "").strip()
        if not name:
            raise AssemblyError(
                f"LR Matrix row {row_number}: File Name is blank."
            )
        names.append(name)

    if serials != list(range(1, len(serials) + 1)):
        raise AssemblyError(
            "LR Matrix serial numbers are not continuous 1..N in physical row "
            "order. Run verify, repair the sheet, then assemble again."
        )
    if len(names) != len(set(names)):
        raise AssemblyError(
            "LR Matrix contains duplicate File Name values. Repair them before "
            "appending new rows."
        )
    return rows


def _database_pdfs(database_dir: Path) -> list[str]:
    if not database_dir.exists():
        return []
    return sorted(
        path.name
        for path in database_dir.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf"
    )


def _read_batch_meta(root: Path) -> dict | None:
    path = root / ".tmp_batch_meta.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _write_json_atomic(path: Path, payload: dict) -> None:
    fd, tmp_name = tempfile.mkstemp(
        prefix=".tmp_", suffix=".json", dir=str(path.parent)
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        tmp_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        os.replace(tmp_path, path)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def _save_to_temp(workbook, directory: Path, prefix: str) -> Path:
    fd, tmp_name = tempfile.mkstemp(
        prefix=prefix, suffix=".xlsx", dir=str(directory)
    )
    os.close(fd)
    path = Path(tmp_name)
    try:
        workbook.save(path)
    except Exception as exc:  # noqa: BLE001 - actionable context
        if path.exists():
            try:
                path.unlink()
            except OSError:
                pass
        raise AssemblyError(
            f"Could not write a temporary workbook: {exc}. Close Excel and retry."
        ) from exc
    return path


def _audit_temp_lr(
    path: Path, expected_rows: int, expected_names: set[str]
) -> list[str]:
    workbook = load_workbook(path, data_only=True)
    try:
        worksheet = workbook.active
        _check_header(worksheet, LR_MATRIX_COLUMNS, "Temporary LR Matrix")
        rows = _read_rows(worksheet, LR_MATRIX_COLUMNS)
        issues: list[str] = []
        if len(rows) != expected_rows:
            issues.append(
                f"Temporary LR Matrix has {len(rows)} rows; expected {expected_rows}."
            )
        serials = []
        names = []
        for row_number, row in rows:
            try:
                serials.append(_serial_int(row.get("S. No."), f"row {row_number}"))
            except AssemblyError as exc:
                issues.append(str(exc))
            name = str(row.get("File Name") or "").strip()
            names.append(name)
            if not str(row.get("Title") or "").strip():
                issues.append(f"Row {row_number}: Title is blank.")
        if serials and serials != list(range(1, len(serials) + 1)):
            issues.append("Temporary LR Matrix serial numbers are not continuous 1..N.")
        if len(names) != len(set(names)):
            issues.append("Temporary LR Matrix has duplicate File Name values.")
        if set(names) != expected_names:
            missing = sorted(expected_names - set(names))
            extra = sorted(set(names) - expected_names)
            if missing:
                issues.append(f"Temporary LR Matrix is missing: {', '.join(missing[:5])}")
            if extra:
                issues.append(f"Temporary LR Matrix has unexpected: {', '.join(extra[:5])}")
        return issues
    finally:
        workbook.close()


def _audit_temp_batch(path: Path, expected_rows: int) -> list[str]:
    workbook = load_workbook(path, data_only=True)
    try:
        worksheet = workbook.active
        _check_header(worksheet, METADATA_COLUMNS, "Temporary metadata batch")
        rows = _read_rows(worksheet, METADATA_COLUMNS)
        if len(rows) != expected_rows:
            return [
                f"Temporary metadata batch has {len(rows)} rows; expected {expected_rows}."
            ]
        return []
    finally:
        workbook.close()


def _backup(path: Path) -> Path | None:
    if not path.exists():
        return None
    backup = path.with_name(path.name + ".bak")
    if backup.exists():
        backup.unlink()
    try:
        shutil.copy2(path, backup)
    except OSError as exc:
        if backup.exists():
            try:
                backup.unlink()
            except OSError:
                pass
        raise AssemblyError(
            f"Could not create a backup of {path.name}: {exc}. Close Excel and retry."
        ) from exc
    return backup


def _restore(path: Path, backup: Path | None) -> bool:
    """Best-effort rollback. Returns True when the path is back in its old state."""
    try:
        if backup is not None and backup.exists():
            os.replace(backup, path)
            return True
        if path.exists():
            path.unlink()
        return True
    except OSError:
        return False


def _cleanup(*paths: Path | None) -> None:
    for path in paths:
        if path is None:
            continue
        try:
            if path.exists():
                path.unlink()
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Main assembly
# ---------------------------------------------------------------------------


def assemble(root: Path, records: list[dict], warnings: list[str]) -> dict:
    """Write validated *records* into the workbooks and move the PDFs."""
    root = Path(root).resolve()
    incoming = root / "Incoming"
    database = root / "Database"
    outputs = root / "Outputs"

    if not root.is_dir():
        raise AssemblyError(f"Project root does not exist: {root}")
    incoming.mkdir(parents=True, exist_ok=True)
    database.mkdir(parents=True, exist_ok=True)
    outputs.mkdir(parents=True, exist_ok=True)

    # --- Preflight: files, collisions, coverage -----------------------------
    missing = [
        record["fileName"]
        for record in records
        if not (incoming / record["fileName"]).is_file()
    ]
    if missing:
        already_in_db = [
            name for name in missing if (database / name).is_file()
        ]
        detail = ""
        if already_in_db:
            detail = (
                " These files are already in Database/, so this batch may already "
                "be assembled. Run verify instead of assembling again."
            )
        raise AssemblyError(
            "These PDFs are not in Incoming/: "
            + ", ".join(sorted(missing))
            + "."
            + detail
        )

    collisions = [
        record["fileName"]
        for record in records
        if (database / record["fileName"]).exists()
    ]
    if collisions:
        raise AssemblyError(
            "These PDFs already exist in Database/ and would be overwritten: "
            + ", ".join(sorted(collisions))
            + ". Run prepare to de-duplicate them first."
        )

    # --- Templates and current state ----------------------------------------
    try:
        template_report = templates.ensure_templates(outputs)
    except templates.TemplateError as exc:
        raise AssemblyError(str(exc)) from exc

    lr_path = outputs / "LR Matrix.xlsx"
    batch_path = Path(template_report["paths"]["metadataBatch"])

    try:
        workbook_lr = load_workbook(lr_path)
        workbook_batch = load_workbook(batch_path)
    except Exception as exc:  # noqa: BLE001 - actionable context
        raise AssemblyError(
            f"Could not open a workbook: {exc}. Close Excel and retry."
        ) from exc

    worksheet_lr = workbook_lr.active
    worksheet_batch = workbook_batch.active

    existing_lr_rows = _audit_existing_lr(worksheet_lr)
    _check_header(worksheet_batch, METADATA_COLUMNS, batch_path.name)
    existing_batch_rows = _read_rows(worksheet_batch, METADATA_COLUMNS)

    database_names = _database_pdfs(database)
    existing_lr_names = [
        str(row.get("File Name") or "").strip()
        for _row_number, row in existing_lr_rows
    ]
    if sorted(existing_lr_names) != database_names:
        missing_in_db = sorted(set(existing_lr_names) - set(database_names))
        missing_in_lr = sorted(set(database_names) - set(existing_lr_names))
        detail = []
        if missing_in_db:
            detail.append(
                "LR rows without a Database PDF: " + ", ".join(missing_in_db[:5])
            )
        if missing_in_lr:
            detail.append(
                "Database PDFs without an LR row: " + ", ".join(missing_in_lr[:5])
            )
        raise AssemblyError(
            "The existing database is inconsistent before appending. "
            + "; ".join(detail)
            + ". Run verify and repair it before assembling."
        )

    try:
        last_serial = get_last_serial(lr_path)
    except SerialError as exc:
        raise AssemblyError(str(exc)) from exc
    if last_serial != len(existing_lr_rows):
        raise AssemblyError(
            f"LR Matrix reports last serial {last_serial} but has "
            f"{len(existing_lr_rows)} data rows. Run verify before assembling."
        )

    # --- Build new rows ------------------------------------------------------
    start_serial = last_serial + 1
    lr_rows: list[dict] = []
    metadata_rows: list[dict] = []
    for offset, record in enumerate(records):
        metadata = record["metadata"]
        lr = record["lrMatrix"]
        lr_values = {"S. No.": start_serial + offset}
        for column in LR_MATRIX_COLUMNS[1:]:
            lr_values[column] = lr.get(column, "")
        lr_rows.append(lr_values)
        metadata_rows.append(
            {column: metadata.get(column, "") for column in METADATA_COLUMNS}
        )

    expected_names = set(existing_lr_names) | {
        record["fileName"] for record in records
    }
    expected_lr_rows = len(existing_lr_rows) + len(records)
    expected_batch_rows = len(existing_batch_rows) + len(records)

    _append_rows(worksheet_lr, LR_MATRIX_COLUMNS, lr_rows)
    _append_rows(worksheet_batch, METADATA_COLUMNS, metadata_rows)

    lr_temp: Path | None = None
    batch_temp: Path | None = None
    lr_backup: Path | None = None
    batch_backup: Path | None = None
    moved: list[str] = []

    try:
        lr_temp = _save_to_temp(workbook_lr, outputs, ".tmp_lr_")
        batch_temp = _save_to_temp(workbook_batch, outputs, ".tmp_batch_")

        audit_issues = _audit_temp_lr(lr_temp, expected_lr_rows, expected_names)
        audit_issues += _audit_temp_batch(batch_temp, expected_batch_rows)
        if audit_issues:
            raise AssemblyError(
                "Refusing to commit because the temporary workbooks failed "
                "verification: " + " | ".join(audit_issues[:8])
            )

        lr_backup = _backup(lr_path)
        batch_backup = _backup(batch_path)

        try:
            os.replace(lr_temp, lr_path)
            lr_temp = None
            os.replace(batch_temp, batch_path)
            batch_temp = None
        except Exception as exc:  # noqa: BLE001 - rollback + context
            lr_restored = _restore(lr_path, lr_backup)
            batch_restored = _restore(batch_path, batch_backup)
            backup_note = ""
            if not (lr_restored and batch_restored):
                backup_note = (
                    f" Manual restore may be needed from {lr_path.name}.bak and "
                    f"{batch_path.name}.bak."
                )
            raise AssemblyError(
                f"Could not replace the workbooks: {exc}. Close Excel and retry; "
                f"no changes were kept.{backup_note}"
            ) from exc

        # Move PDFs only after both workbooks are committed.
        try:
            for record in records:
                name = record["fileName"]
                shutil.move(str(incoming / name), str(database / name))
                moved.append(name)
        except Exception as exc:  # noqa: BLE001 - rollback + context
            for name in moved:
                source = database / name
                if source.exists():
                    try:
                        shutil.move(str(source), str(incoming / name))
                    except OSError:
                        pass
            lr_restored = _restore(lr_path, lr_backup)
            batch_restored = _restore(batch_path, batch_backup)
            backup_note = ""
            if not (lr_restored and batch_restored):
                backup_note = (
                    f" Manual restore may be needed from {lr_path.name}.bak and "
                    f"{batch_path.name}.bak."
                )
            raise AssemblyError(
                f"Could not move a PDF into Database/: {exc}. The workbook "
                "changes were rolled back and moved files were returned to "
                f"Incoming/.{backup_note}"
            ) from exc

        _cleanup(lr_backup, batch_backup)
        lr_backup = batch_backup = None

        state = {
            "assembledAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "rowsAdded": len(records),
            "startSerial": start_serial,
            "endSerial": start_serial + len(records) - 1,
            "files": [record["fileName"] for record in records],
            "lrMatrix": str(lr_path),
            "metadataBatch": str(batch_path),
            "warnings": warnings,
            "batchMeta": _read_batch_meta(root),
        }
        try:
            _write_json_atomic(root / ".tmp_assemble_state.json", state)
        except Exception as exc:  # noqa: BLE001 - non-critical audit aid
            warnings.append(
                f"Workbooks and PDFs were updated, but the temporary state file "
                f"could not be written: {exc}"
            )

        return {
            "root": str(root),
            "rowsAdded": len(records),
            "startSerial": start_serial,
            "endSerial": start_serial + len(records) - 1,
            "filesMoved": moved,
            "lrMatrix": str(lr_path),
            "metadataBatch": str(batch_path),
            "metaDataFilled": template_report["paths"]["metaDataFilled"],
            "existingLrRows": len(existing_lr_rows),
            "existingBatchRows": len(existing_batch_rows),
            "warnings": warnings,
        }
    finally:
        _cleanup(lr_temp, batch_temp)
        try:
            workbook_lr.close()
            workbook_batch.close()
        except Exception:  # noqa: BLE001 - best effort
            pass
