#!/usr/bin/env python3
"""Unified CLI for the LR Database skill.

Commands
--------
prepare  --root ROOT
    Create folders/log, normalise filenames, de-duplicate, extract PDF text to
    ``.tmp_batch_text.json`` and ``.tmp_text/``.

assemble --root ROOT --input PATH
    Validate ``.tmp_extractions.json`` against the strict schema, append to the
    SOP workbooks atomically, and move keeper PDFs into ``Database/``.

verify   --root ROOT
    Audit the Database-to-LR-Matrix bijection, serial continuity, and titles;
    append a timestamped entry to ``SOP-Log.md``.

template --file NAME
    Print a blank extraction record with all required keys (helper for agents).

Every command prints one JSON object to stdout and exits 0 on success or 1 on
failure.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

MAX_TEXT_CHARS = 50_000
FIRST_PAGE_CHARS = 3_000
VERSION = "1.0.0"

logging.getLogger("pypdf").setLevel(logging.ERROR)

IMPORT_ERROR: Exception | None = None
try:
    from pypdf import PdfReader

    import assemble_excel as assembler
    import create_templates as templates
    from get_last_serial import SerialError, get_last_serial
except ImportError as exc:  # pragma: no cover - exercised by the CLI guard
    IMPORT_ERROR = exc

    class _MissingDependencyError(Exception):
        pass

    class _MissingAssembler:
        ValidationError = _MissingDependencyError
        AssemblyError = _MissingDependencyError

    class _MissingTemplates:
        TemplateError = _MissingDependencyError

    SerialError = _MissingDependencyError
    assembler = _MissingAssembler()
    templates = _MissingTemplates()


# ---------------------------------------------------------------------------
# Errors and output
# ---------------------------------------------------------------------------


class PipelineError(RuntimeError):
    """Actionable, user-facing pipeline failure."""


def emit(payload: dict, code: int = 0) -> None:
    payload.setdefault("skill", "lr-database")
    payload.setdefault("version", VERSION)
    print(json.dumps(payload, ensure_ascii=True, default=str), flush=True)
    raise SystemExit(code)


def _error(message: str, command: str | None = None, issues: list[str] | None = None) -> dict:
    payload = {"status": "error", "error": message}
    if command:
        payload["command"] = command
    if issues:
        payload["issues"] = issues
    return payload


# ---------------------------------------------------------------------------
# Small utilities
# ---------------------------------------------------------------------------


def resolve_root(raw: str | Path, create: bool = False) -> Path:
    root = Path(raw).expanduser()
    if not root.is_absolute():
        root = Path.cwd() / root
    root = root.resolve()
    if create:
        root.mkdir(parents=True, exist_ok=True)
    if not root.exists():
        raise PipelineError(f"Project root does not exist: {root}")
    if not root.is_dir():
        raise PipelineError(f"Project root is not a directory: {root}")
    return root


def ensure_folders(root: Path) -> None:
    for name in ("Incoming", "Database", "Outputs", "Set-Aside"):
        (root / name).mkdir(parents=True, exist_ok=True)


LOG_HEADER = """# SOP Log — LR Database runs

- Initialized: {date}
- House rule: keep-all. Decision/Reason stay blank unless screening is explicitly requested.
- Audit entries are appended by `pipeline.py verify`.
"""


def ensure_log(root: Path) -> Path:
    log_path = root / "SOP-Log.md"
    if not log_path.exists():
        log_path.write_text(
            LOG_HEADER.format(date=date.today().isoformat()), encoding="utf-8"
        )
    return log_path


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
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


def unique_destination(directory: Path, name: str) -> Path:
    candidate = directory / name
    if not candidate.exists():
        return candidate
    stem = Path(name).stem
    suffix = Path(name).suffix
    for counter in range(2, 10_000):
        candidate = directory / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
    raise PipelineError(f"Could not find a free filename for {name} in {directory}.")


def move_to_setaside(source: Path, setaside: Path, reason: str) -> Path:
    destination = unique_destination(setaside, source.name)
    try:
        shutil.move(str(source), str(destination))
    except Exception as exc:  # noqa: BLE001 - actionable context
        raise PipelineError(
            f"Could not move {source.name} to Set-Aside/: {exc}. Close the file and retry."
        ) from exc
    return destination


# ---------------------------------------------------------------------------
# PDF indexing
# ---------------------------------------------------------------------------


def clean_pdf_text(text: str | None) -> str:
    if not text:
        return ""
    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\x0b\x0c]", "\n", text)
    text = re.sub(r"[\x00-\x08\x0e-\x1f]", "", text)
    return text


def normalise_unicode(text: str) -> str:
    return unicodedata.normalize("NFKC", text).casefold()


def title_tokens(text: str) -> list[str]:
    return [
        token
        for token in re.findall(r"[^\W_]+", normalise_unicode(text), flags=re.UNICODE)
        if len(token) > 3
    ]


def stem_key(text: str) -> str:
    return re.sub(r"[^\w]+", "", normalise_unicode(text), flags=re.UNICODE)


def jaccard(left: list[str], right: list[str]) -> float:
    left_set, right_set = set(left), set(right)
    union = left_set | right_set
    if not union:
        return 0.0
    return len(left_set & right_set) / len(union)


DOI_PATTERN = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.IGNORECASE)


def doi_candidates(text: str) -> set[str]:
    found: set[str] = set()
    for match in DOI_PATTERN.findall(text):
        doi = match.rstrip(".,;")
        while doi.endswith(")") and doi.count(")") > doi.count("("):
            doi = doi[:-1]
        if doi:
            found.add(doi)
    return found


def index_pdf(path: Path, full: bool = False) -> dict:
    """Read a PDF into an index record. Raises PipelineError when unreadable."""
    try:
        reader = PdfReader(str(path))
        if getattr(reader, "is_encrypted", False):
            try:
                unlocked = bool(reader.decrypt(""))
            except Exception:  # noqa: BLE001 - treat as locked
                unlocked = False
            if not unlocked:
                raise PipelineError(
                    "PDF is password-protected (an empty password did not unlock it)."
                )
        pages = reader.pages
        num_pages = len(pages)
        if num_pages == 0:
            raise PipelineError("PDF has no pages.")

        first_page = clean_pdf_text(pages[0].extract_text() or "")
        full_text = ""
        truncated = False
        if full:
            parts: list[str] = []
            total = 0
            for page in pages:
                page_text = clean_pdf_text(page.extract_text() or "")
                if not page_text.strip():
                    continue
                parts.append(page_text)
                total += len(page_text)
                if total >= MAX_TEXT_CHARS:
                    truncated = True
                    break
            full_text = "\n\n".join(parts)[:MAX_TEXT_CHARS]
            if not full_text.strip():
                raise PipelineError(
                    "No extractable text (scanned/image-only PDF; OCR is required)."
                )
    except PipelineError:
        raise
    except Exception as exc:  # noqa: BLE001 - actionable context
        raise PipelineError(f"Unreadable PDF: {exc}") from exc

    return {
        "path": path,
        "fileName": path.name,
        "stemKey": stem_key(path.stem),
        "tokens": title_tokens(path.stem),
        "dois": doi_candidates(first_page),
        "firstPage": first_page,
        "firstPageNorm": normalise_unicode(first_page),
        "size": path.stat().st_size,
        "numPages": num_pages,
        "fullText": full_text,
        "truncated": truncated,
        "textChars": len(full_text),
    }


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------


def find_database_duplicate(info: dict, database_index: list[dict]) -> tuple[dict, str] | None:
    for entry in database_index:
        if info["stemKey"] and info["stemKey"] == entry["stemKey"]:
            return entry, "Normalized filename match"

    if info["dois"]:
        for entry in database_index:
            shared = info["dois"] & entry["dois"]
            if shared:
                return entry, f"DOI match ({sorted(shared)[0]})"

    best: tuple[dict, float] | None = None
    for entry in database_index:
        similarity = jaccard(info["tokens"], entry["tokens"])
        if similarity <= 0.70:
            continue
        forward = sum(
            1
            for token in info["tokens"]
            if len(token) > 4 and token in entry["firstPageNorm"]
        )
        reverse = sum(
            1
            for token in entry["tokens"]
            if len(token) > 4 and token in info["firstPageNorm"]
        )
        if forward >= 3 or reverse >= 3:
            if best is None or similarity > best[1]:
                best = (entry, similarity)

    if best is not None:
        return best[0], f"Title/first-page match (Jaccard {best[1]:.2f})"
    return None


def find_batch_duplicate(info: dict, keepers: list[dict]) -> tuple[dict, str] | None:
    if info["dois"]:
        for entry in keepers:
            shared = info["dois"] & entry["dois"]
            if shared:
                return entry, f"DOI match ({sorted(shared)[0]})"

    for entry in keepers:
        if info["stemKey"] and info["stemKey"] == entry["stemKey"]:
            return entry, "Normalized filename match"

    best: tuple[dict, float] | None = None
    for entry in keepers:
        similarity = jaccard(info["tokens"], entry["tokens"])
        if similarity > 0.85 and (best is None or similarity > best[1]):
            best = (entry, similarity)
    if best is not None:
        return best[0], f"Title Jaccard {best[1]:.2f} (> 0.85)"
    return None


# ---------------------------------------------------------------------------
# Filename normalisation
# ---------------------------------------------------------------------------

_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def normalise_pdf_name(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", stem)
    stem = re.sub(r"\s+", " ", stem).strip().rstrip(" .")
    if not stem:
        stem = "document"
    if stem.upper() in _RESERVED_NAMES:
        stem += "_"
    if len(stem) > 150:
        stem = stem[:150].rstrip()
    return f"{stem}.pdf"


def normalise_incoming_filenames(incoming: Path) -> list[dict]:
    renamed: list[dict] = []
    pdfs = sorted(
        path for path in incoming.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"
    )
    for path in pdfs:
        target_name = normalise_pdf_name(path.name)
        target = path.with_name(target_name)
        if target.name == path.name:
            continue
        target = unique_destination(incoming, target_name)
        try:
            path.rename(target)
        except OSError as exc:
            raise PipelineError(
                f"Could not normalise {path.name}: {exc}. Close the file and retry."
            ) from exc
        renamed.append({"from": path.name, "to": target.name})
    return renamed


# ---------------------------------------------------------------------------
# prepare
# ---------------------------------------------------------------------------


def cmd_prepare(raw_root: str | Path) -> dict:
    root = resolve_root(raw_root, create=True)
    incoming = root / "Incoming"
    database = root / "Database"
    setaside = root / "Set-Aside"
    outputs = root / "Outputs"
    ensure_folders(root)
    ensure_log(root)

    text_dir = root / ".tmp_text"
    if text_dir.exists():
        shutil.rmtree(text_dir)
    text_dir.mkdir(parents=True, exist_ok=True)

    rejects: list[dict] = []
    duplicates: list[dict] = []
    warnings: list[str] = []

    incoming_entries = sorted(
        path for path in incoming.iterdir() if path.is_file()
    )
    for path in incoming_entries:
        if path.suffix.lower() != ".pdf":
            destination = move_to_setaside(path, setaside, "Not a PDF")
            rejects.append({"file": path.name, "movedTo": destination.name, "reason": "Not a PDF"})

    renamed = normalise_incoming_filenames(incoming)
    incoming_pdfs = sorted(
        path for path in incoming.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"
    )

    database_index: list[dict] = []
    for path in sorted(database.iterdir() if database.exists() else []):
        if not path.is_file() or path.suffix.lower() != ".pdf":
            continue
        try:
            database_index.append(index_pdf(path, full=False))
        except PipelineError as exc:
            warnings.append(
                f"Database PDF {path.name} could not be indexed for de-duplication: {exc}"
            )
            # Still index the filename so a same-name upload is caught.
            database_index.append(
                {
                    "path": path,
                    "fileName": path.name,
                    "stemKey": stem_key(path.stem),
                    "tokens": title_tokens(path.stem),
                    "dois": set(),
                    "firstPage": "",
                    "firstPageNorm": "",
                    "size": path.stat().st_size,
                }
            )

    keepers: list[dict] = []
    for path in incoming_pdfs:
        try:
            info = index_pdf(path, full=False)
        except PipelineError as exc:
            destination = move_to_setaside(path, setaside, str(exc))
            rejects.append({"file": path.name, "movedTo": destination.name, "reason": str(exc)})
            continue

        match = find_database_duplicate(info, database_index)
        if match is not None:
            entry, reason = match
            destination = move_to_setaside(
                path, setaside, f"{reason} — matches Database/{entry['fileName']}"
            )
            duplicates.append(
                {
                    "file": path.name,
                    "movedTo": destination.name,
                    "reason": reason,
                    "matchedWith": f"Database/{entry['fileName']}",
                }
            )
            continue

        match = find_batch_duplicate(info, keepers)
        if match is not None:
            entry, reason = match
            if info["size"] > entry["size"]:
                keepers.remove(entry)
                destination = move_to_setaside(
                    entry["path"],
                    setaside,
                    f"{reason}; larger file kept: {info['fileName']}",
                )
                duplicates.append(
                    {
                        "file": entry["fileName"],
                        "movedTo": destination.name,
                        "reason": f"{reason}; larger file kept",
                        "matchedWith": f"Incoming/{info['fileName']}",
                    }
                )
                keepers.append(info)
            else:
                destination = move_to_setaside(
                    path,
                    setaside,
                    f"{reason}; larger file kept: {entry['fileName']}",
                )
                duplicates.append(
                    {
                        "file": path.name,
                        "movedTo": destination.name,
                        "reason": f"{reason}; larger file kept",
                        "matchedWith": f"Incoming/{entry['fileName']}",
                    }
                )
            continue

        keepers.append(info)

    # Full text extraction happens only for accepted keepers.
    batch_entries: list[dict] = []
    accepted_keepers: list[dict] = []
    for info in keepers:
        try:
            full_info = index_pdf(info["path"], full=True)
        except PipelineError as exc:
            destination = move_to_setaside(info["path"], setaside, str(exc))
            rejects.append(
                {"file": info["fileName"], "movedTo": destination.name, "reason": str(exc)}
            )
            continue
        accepted_keepers.append(full_info)
        batch_entries.append(
            {
                "fileName": full_info["fileName"],
                "numPages": full_info["numPages"],
                "textChars": full_info["textChars"],
                "truncated": full_info["truncated"],
                "firstPage": full_info["firstPage"][:FIRST_PAGE_CHARS],
                "fullText": full_info["fullText"],
                "doiCandidates": sorted(full_info["dois"]),
                "filenameTitleTokens": full_info["tokens"],
            }
        )

    # One readable text file per paper makes agent extraction easier.
    seen_names: dict[str, int] = {}
    text_files: list[str] = []
    for entry in batch_entries:
        safe_stem = re.sub(r"[^\w.-]+", "_", Path(entry["fileName"]).stem, flags=re.UNICODE)
        safe_stem = safe_stem[:80].strip("._") or "paper"
        count = seen_names.get(safe_stem.casefold(), 0) + 1
        seen_names[safe_stem.casefold()] = count
        if count > 1:
            safe_stem = f"{safe_stem}_{count}"
        text_path = text_dir / f"{safe_stem}.txt"
        body = (
            f"FILE: {entry['fileName']}\n"
            f"PAGES: {entry['numPages']}\n"
            f"CHARS: {entry['textChars']} (truncated: {'yes' if entry['truncated'] else 'no'})\n\n"
            "=== FIRST PAGE ===\n"
            f"{entry['firstPage']}\n\n"
            "=== FULL TEXT ===\n"
            f"{entry['fullText']}\n"
        )
        text_path.write_text(body, encoding="utf-8")
        text_files.append(str(text_path))

    created_at = datetime.now().astimezone().isoformat(timespec="seconds")
    batch_text_path = root / ".tmp_batch_text.json"
    meta_path = root / ".tmp_batch_meta.json"

    write_json_atomic(batch_text_path, batch_entries)
    write_json_atomic(
        meta_path,
        {
            "createdAt": created_at,
            "root": str(root),
            "incomingCount": len(incoming_pdfs),
            "acceptedCount": len(accepted_keepers),
            "keepers": [entry["fileName"] for entry in batch_entries],
            "duplicates": duplicates,
            "rejected": rejects,
            "renamed": renamed,
            "warnings": warnings,
        },
    )

    accepted_names = [entry["fileName"] for entry in batch_entries]
    return {
        "status": "ok",
        "command": "prepare",
        "root": str(root),
        "incomingPdfs": len(incoming_pdfs),
        "renamed": renamed,
        "accepted": len(accepted_names),
        "acceptedFiles": accepted_names,
        "duplicates": duplicates,
        "rejected": rejects,
        "warnings": warnings,
        "batchTextFile": str(batch_text_path),
        "textDir": str(text_dir),
        "nextStep": (
            "Extract all 87 metadata + 19 LR fields for each accepted PDF into "
            ".tmp_extractions.json, then run assemble."
            if accepted_names
            else "No new PDFs were accepted. Stop here; Incoming/ is empty of processable papers."
        ),
    }


# ---------------------------------------------------------------------------
# assemble
# ---------------------------------------------------------------------------


def _resolve_input(root: Path, raw_input: str) -> Path:
    path = Path(raw_input).expanduser()
    if path.is_absolute():
        return path
    root_candidate = root / path
    if root_candidate.exists():
        return root_candidate
    return (Path.cwd() / path).resolve()


def cmd_assemble(raw_root: str | Path, raw_input: str) -> dict:
    root = resolve_root(raw_root)
    input_path = _resolve_input(root, raw_input)
    if not input_path.is_file():
        raise PipelineError(
            f"Extraction file not found: {input_path}. Save the payload as "
            f"{root / '.tmp_extractions.json'} and retry."
        )

    try:
        payload = json.loads(input_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise PipelineError(
            f"{input_path.name} is not valid JSON: line {exc.lineno}, column "
            f"{exc.colno}: {exc.msg}"
        ) from exc
    except OSError as exc:
        raise PipelineError(f"Could not read {input_path}: {exc}") from exc

    records, warnings = assembler.validate_extractions(payload)

    batch_text_path = root / ".tmp_batch_text.json"
    if batch_text_path.exists():
        try:
            batch_entries = json.loads(batch_text_path.read_text(encoding="utf-8"))
            expected = {entry["fileName"] for entry in batch_entries}
        except (OSError, json.JSONDecodeError, KeyError, TypeError) as exc:
            raise PipelineError(
                f"Could not read {batch_text_path.name}: {exc}. Run prepare again."
            ) from exc
        provided = {record["fileName"] for record in records}
        missing = sorted(expected - provided)
        extra = sorted(provided - expected)
        if missing or extra:
            issues = []
            if missing:
                issues.append("Missing extraction record(s): " + ", ".join(missing))
            if extra:
                issues.append(
                    "Payload contains file(s) that are not keepers in this batch: "
                    + ", ".join(extra)
                )
            raise PipelineError(
                "The extraction payload does not match the prepared batch. "
                + " | ".join(issues)
            )
    else:
        warnings.append(
            "No .tmp_batch_text.json found; batch coverage could not be cross-checked."
        )

    result = assembler.assemble(root, records, warnings)
    result.update(
        {
            "status": "ok",
            "command": "assemble",
            "tempFilesToRemove": [
                str(root / ".tmp_batch_text.json"),
                str(root / ".tmp_extractions.json"),
                str(root / ".tmp_batch_meta.json"),
                str(root / ".tmp_assemble_state.json"),
                str(root / ".tmp_text"),
            ],
        }
    )
    return result


# ---------------------------------------------------------------------------
# verify
# ---------------------------------------------------------------------------


def _database_pdfs(database: Path) -> list[str]:
    if not database.exists():
        return []
    return sorted(
        path.name
        for path in database.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf"
    )


def _read_state(root: Path) -> dict | None:
    path = root / ".tmp_assemble_state.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _read_meta(root: Path) -> dict | None:
    path = root / ".tmp_batch_meta.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _audit_lr_matrix(lr_path: Path, database_names: list[str]) -> dict:
    from openpyxl import load_workbook

    issues: list[str] = []
    warnings: list[str] = []
    checks = {
        "headersValid": True,
        "countsMatch": True,
        "filenameBijection": True,
        "serialsContinuous": True,
        "titlesPresent": True,
    }

    if not lr_path.exists():
        if database_names:
            issues.append(
                "Outputs/LR Matrix.xlsx is missing but Database/ contains "
                f"{len(database_names)} PDF(s). Run assemble to create it."
            )
            checks["countsMatch"] = False
            checks["filenameBijection"] = False
        else:
            warnings.append(
                "Outputs/LR Matrix.xlsx does not exist yet; the database is empty."
            )
        return {
            "issues": issues,
            "warnings": warnings,
            "checks": checks,
            "lrRows": 0,
            "serials": [],
        }

    try:
        workbook = load_workbook(lr_path, data_only=True)
    except Exception as exc:  # noqa: BLE001 - actionable context
        raise PipelineError(
            f"Could not open Outputs/LR Matrix.xlsx: {exc}. Close Excel and retry."
        ) from exc

    try:
        worksheet = workbook.active
        header = [
            worksheet.cell(row=1, column=column).value
            for column in range(1, len(templates.LR_MATRIX_COLUMNS) + 1)
        ]
        if header != templates.LR_MATRIX_COLUMNS:
            checks["headersValid"] = False
            checks["countsMatch"] = False
            checks["filenameBijection"] = False
            checks["serialsContinuous"] = False
            checks["titlesPresent"] = False
            issues.append(
                "Outputs/LR Matrix.xlsx header does not match the 20 SOP columns. "
                f"Found {len(header)} columns; restore the documented header."
            )
            return {
                "issues": issues,
                "warnings": warnings,
                "checks": checks,
                "lrRows": 0,
                "serials": [],
            }

        rows: list[tuple[int, dict]] = []
        for row_number in range(2, worksheet.max_row + 1):
            values = [
                worksheet.cell(row=row_number, column=column).value
                for column in range(1, len(templates.LR_MATRIX_COLUMNS) + 1)
            ]
            if all(
                value is None or (isinstance(value, str) and not value.strip())
                for value in values
            ):
                continue
            rows.append((row_number, dict(zip(templates.LR_MATRIX_COLUMNS, values))))
    finally:
        workbook.close()

    serials: list[int] = []
    serial_failures = False
    for row_number, row in rows:
        raw = row.get("S. No.")
        try:
            serial = assembler._serial_int(raw, f"LR Matrix row {row_number}")
        except assembler.AssemblyError as exc:
            issues.append(str(exc))
            serial_failures = True
            continue
        if serial < 1:
            issues.append(f"LR Matrix row {row_number}: S. No. must be >= 1.")
            serial_failures = True
            continue
        serials.append(serial)

    if serial_failures:
        checks["serialsContinuous"] = False

    if not serial_failures:
        expected_serials = list(range(1, len(rows) + 1))
        if serials != expected_serials:
            checks["serialsContinuous"] = False
            duplicate_serials = sorted(
                {serial for serial in serials if serials.count(serial) > 1}
            )
            detail = []
            if duplicate_serials:
                detail.append("duplicates: " + ", ".join(map(str, duplicate_serials[:10])))
            missing = sorted(set(expected_serials) - set(serials))
            if missing:
                detail.append("gaps: " + ", ".join(map(str, missing[:10])))
            if not detail:
                detail.append("values are not in physical row order 1..N")
            issues.append("S. No. is not continuous: " + "; ".join(detail) + ".")
        elif serials != sorted(serials):
            checks["serialsContinuous"] = False
            issues.append("S. No. values are not in ascending physical row order.")

    blank_titles = [
        row_number
        for row_number, row in rows
        if not str(row.get("Title") or "").strip()
    ]
    if blank_titles:
        checks["titlesPresent"] = False
        issues.append(
            "Blank Title in LR row(s): " + ", ".join(map(str, blank_titles[:10]))
            + (" ..." if len(blank_titles) > 10 else "")
        )

    lr_names = [str(row.get("File Name") or "").strip() for _row_number, row in rows]
    if any(not name for name in lr_names):
        checks["filenameBijection"] = False
        issues.append("One or more LR rows have a blank File Name.")
    duplicate_names = sorted(
        {name for name in lr_names if name and lr_names.count(name) > 1}
    )
    if duplicate_names:
        checks["filenameBijection"] = False
        issues.append(
            "Duplicate File Name in LR matrix: " + ", ".join(duplicate_names[:10])
        )

    if len(rows) != len(database_names):
        checks["countsMatch"] = False
        issues.append(
            f"Row count mismatch: Database/ has {len(database_names)} PDF(s) but "
            f"LR Matrix has {len(rows)} row(s)."
        )

    missing_in_lr = sorted(set(database_names) - set(lr_names))
    missing_in_db = sorted(set(lr_names) - set(database_names))
    if missing_in_lr or missing_in_db:
        checks["filenameBijection"] = False
        if missing_in_lr:
            issues.append(
                "Database PDFs without an LR row: " + ", ".join(missing_in_lr[:10])
                + (" ..." if len(missing_in_lr) > 10 else "")
            )
        if missing_in_db:
            issues.append(
                "LR rows without a Database PDF: " + ", ".join(missing_in_db[:10])
                + (" ..." if len(missing_in_db) > 10 else "")
            )

    return {
        "issues": issues,
        "warnings": warnings,
        "checks": checks,
        "lrRows": len(rows),
        "serials": serials,
    }


def _append_audit(
    root: Path,
    passed: bool,
    checks: dict,
    database_count: int,
    lr_count: int,
    serials: list[int],
    issues: list[str],
    warnings: list[str],
    state: dict | None,
    meta: dict | None,
    database_names: list[str],
) -> Path:
    log_path = ensure_log(root)
    now = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
    lines = ["", f"## {now} — Verification {'PASSED' if passed else 'FAILED'}", f"- Project: `{root}`"]

    if state and isinstance(state.get("rowsAdded"), int):
        state_files = set(state.get("files") or [])
        if state_files and state_files <= set(database_names):
            lines.append(
                f"- Batch: {state['rowsAdded']} row(s), S. No. "
                f"{state.get('startSerial')}–{state.get('endSerial')}"
            )
            if meta:
                duplicates = meta.get("duplicates") or []
                if duplicates:
                    lines.append(f"- Duplicates set aside: {len(duplicates)}")
                    for duplicate in duplicates[:10]:
                        lines.append(
                            f"  - {duplicate.get('file')}: {duplicate.get('reason')} "
                            f"(matched {duplicate.get('matchedWith')})"
                        )

    lines.append(f"- Database PDFs: {database_count}")
    lines.append(f"- LR Matrix rows: {lr_count}")
    if serials:
        lines.append(f"- Serial range: {min(serials)}–{max(serials)}")

    check_text = ", ".join(
        f"{name}={'PASS' if ok else 'FAIL'}" for name, ok in checks.items()
    )
    lines.append(f"- Checks: {check_text}")

    if issues:
        lines.append("- Issues:")
        for issue in issues[:20]:
            lines.append(f"  - {issue}")
        if len(issues) > 20:
            lines.append(f"  - ... and {len(issues) - 20} more")
    if warnings:
        lines.append("- Warnings:")
        for warning in warnings[:10]:
            lines.append(f"  - {warning}")

    with log_path.open("a", encoding="utf-8") as handle:
        handle.write("\n".join(lines) + "\n")
    return log_path


def cmd_verify(raw_root: str | Path) -> dict:
    root = resolve_root(raw_root)
    database = root / "Database"
    lr_path = root / "Outputs" / "LR Matrix.xlsx"
    database_names = _database_pdfs(database)

    audit = _audit_lr_matrix(lr_path, database_names)
    issues = audit["issues"]
    warnings = audit["warnings"]
    checks = audit["checks"]
    serials = audit["serials"]
    lr_rows = audit["lrRows"]

    meta_filled = root / "Outputs" / "Meta Data Filled.xlsx"
    if not meta_filled.exists():
        warnings.append(
            "Outputs/Meta Data Filled.xlsx is missing; assemble recreates it as a "
            "frozen baseline."
        )

    passed = not issues and all(checks.values())
    state = _read_state(root)
    meta = _read_meta(root)
    log_path = _append_audit(
        root,
        passed,
        checks,
        len(database_names),
        lr_rows,
        serials,
        issues,
        warnings,
        state,
        meta,
        database_names,
    )

    payload = {
        "status": "ok" if passed else "error",
        "command": "verify",
        "passed": passed,
        "root": str(root),
        "databasePdfs": len(database_names),
        "lrRows": lr_rows,
        "serialRange": [min(serials), max(serials)] if serials else [],
        "checks": checks,
        "issues": issues,
        "warnings": warnings,
        "log": str(log_path),
    }
    if passed:
        payload["tempFilesToRemove"] = [
            str(root / ".tmp_batch_text.json"),
            str(root / ".tmp_extractions.json"),
            str(root / ".tmp_batch_meta.json"),
            str(root / ".tmp_assemble_state.json"),
            str(root / ".tmp_text"),
        ]
    return payload


# ---------------------------------------------------------------------------
# template
# ---------------------------------------------------------------------------


def cmd_template(file_name: str) -> dict:
    if Path(file_name).name != file_name or "/" in file_name or "\\" in file_name:
        raise PipelineError("--file must be a PDF basename with no folders.")
    if not file_name.lower().endswith(".pdf"):
        raise PipelineError("--file must end with .pdf.")

    now = datetime.now().astimezone().isoformat(timespec="seconds")
    today = date.today().isoformat()

    metadata = {column: "" for column in templates.METADATA_COLUMNS}
    metadata.update(
        {
            "File Name": file_name,
            "Item Type": "journalArticle",
            "Num Pages": 0,
            "Number Of Volumes": 1,
            "Language": "en",
            "Date Added": now,
            "Date Modified": now,
            "Access Date": today,
            "File Attachments": f"Database/{file_name}",
        }
    )

    lr_matrix = {column: "" for column in templates.LR_MATRIX_COLUMNS[1:]}
    lr_matrix.update(
        {
            "File Name": file_name,
            "Tag 2 (Type of Research)": "Quantitative",
            "Decision (For Me)": "",
            "Reason (For Me)": "",
        }
    )

    return {
        "status": "ok",
        "command": "template",
        "fileName": file_name,
        "metadata": metadata,
        "lrMatrix": lr_matrix,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:  # type: ignore[override]
        raise PipelineError(f"Invalid arguments: {message}")


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        prog="pipeline.py",
        description="LR Database pipeline: prepare, assemble, verify.",
    )
    parser.add_argument("--debug", action="store_true", help="Include tracebacks in errors.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    prepare = subparsers.add_parser("prepare", help="Prepare and de-duplicate a batch.")
    prepare.add_argument("--root", required=True, help="Project root folder.")

    assemble = subparsers.add_parser("assemble", help="Validate and write the batch.")
    assemble.add_argument("--root", required=True, help="Project root folder.")
    assemble.add_argument(
        "--input",
        required=True,
        help="Path to .tmp_extractions.json (relative paths resolve against --root).",
    )

    verify = subparsers.add_parser("verify", help="Audit the database and log the result.")
    verify.add_argument("--root", required=True, help="Project root folder.")

    template = subparsers.add_parser("template", help="Print a blank extraction record.")
    template.add_argument("--file", required=True, help="Exact PDF basename.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    command = None
    debug = False
    try:
        args = parser.parse_args(argv)
        command = args.command
        debug = bool(getattr(args, "debug", False))

        if IMPORT_ERROR is not None:
            raise PipelineError(
                f"Missing Python dependency: {IMPORT_ERROR}. Install them with: "
                f'python -m pip install -r "{SKILL_ROOT / "scripts" / "requirements.txt"}"'
            )

        if command == "prepare":
            payload = cmd_prepare(args.root)
        elif command == "assemble":
            payload = cmd_assemble(args.root, args.input)
        elif command == "verify":
            payload = cmd_verify(args.root)
        elif command == "template":
            payload = cmd_template(args.file)
        else:  # pragma: no cover - argparse guarantees a command
            raise PipelineError(f"Unknown command: {command}")

        exit_code = 0
        if command == "verify" and not payload.get("passed", False):
            exit_code = 1
        emit(payload, exit_code)
        return exit_code
    except assembler.ValidationError as exc:
        emit(
            _error(
                "The extraction payload failed validation. Fix the listed issues "
                "and rerun assemble. No workbooks were changed.",
                command,
                exc.issues,
            ),
            1,
        )
    except (PipelineError, assembler.AssemblyError, templates.TemplateError, SerialError) as exc:
        emit(_error(str(exc), command), 1)
    except KeyboardInterrupt:
        emit(_error("Cancelled by user.", command), 1)
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 - last-resort JSON error
        if debug:
            import traceback

            emit(
                _error(
                    f"Unexpected error: {exc}\n{traceback.format_exc()}",
                    command,
                ),
                1,
            )
        emit(_error(f"Unexpected error: {type(exc).__name__}: {exc}", command), 1)
    return 1


if __name__ == "__main__":
    sys.exit(main())
