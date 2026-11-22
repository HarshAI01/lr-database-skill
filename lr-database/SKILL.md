---
name: lr-database
description: >
  Build, extend, or audit a Literature Review Database from research PDFs. Use
  when the user asks to process a batch of papers into a literature review
  matrix, de-duplicate PDFs, extract 87 metadata fields and 20 LR summary
  fields, create SOP-formatted Excel workbooks (LR Matrix.xlsx, metadata batch
  files), verify a literature database, or log a literature-review run. Trigger
  phrases include "literature review database", "LR matrix", "extract papers",
  "dedupe PDFs", "metadata batch", "SOP log", and "literature review Excel".
license: MIT
version: 1.0.0
---

# LR Database

Turn a folder of PDFs into a verified, SOP-formatted Literature Review
Database. The skill is pure Python: it uses the host agent's existing model for
reading and summarising, and never calls an LLM API or asks for a key.

## Outputs

```
<PROJECT_ROOT>/
├── Incoming/                  # drop new PDFs here
├── Database/                  # processed PDFs (source of truth)
├── Outputs/
│   ├── LR Matrix.xlsx         # 20 columns, cumulative, continuous S. No.
│   ├── Meta Data Filled.xlsx  # 87 columns, frozen baseline, never edited
│   └── metadata batch YYYY-MM-DD.xlsx
├── Set-Aside/                 # duplicates, rejects, non-PDFs
└── SOP-Log.md                 # timestamped audit trail
```

## Non-negotiables

- **Keep-all.** Leave `Decision (For Me)` and `Reason (For Me)` blank unless the
  user explicitly requests screening against a written scope.
- **Never renumber or reorder** existing `LR Matrix.xlsx` rows. Append only.
- **Never edit `Meta Data Filled.xlsx`** after it exists. It is a frozen baseline.
- **Never delete PDFs.** Duplicates and rejects move to `Set-Aside/`.
- **Never invent** authors, years, DOIs, statistics, or findings. Use `""` when
  the PDF gives nothing.
- **One batch at a time.** Finish `prepare` → `assemble` → `verify` before
  starting another.
- Read `references/sop-rules.md` before extracting. It is the field-level
  authority; this file is the workflow.

## Setup

Resolve `SKILL_DIR` as the folder containing this `SKILL.md`, then confirm the
dependencies once per machine:

```bash
python -c "import openpyxl, pypdf, jsonschema; print('ok')"
```

If that fails, install them (ask the user first if the host requires approval):

```bash
python -m pip install -r "<SKILL_DIR>/scripts/requirements.txt"
```

Use `python3` instead of `python` on macOS/Linux when that is the available
executable. Quote every path that contains spaces.

## Workflow

### Step 1 — Prepare

```bash
python "<SKILL_DIR>/scripts/pipeline.py" prepare --root "<PROJECT_ROOT>"
```

This creates the folders and `SOP-Log.md`, normalises filenames, de-duplicates
against `Database/` and within the batch, moves losers to `Set-Aside/`, and
writes:

- `.tmp_batch_text.json` — full text for every accepted PDF.
- `.tmp_text/<paper>.txt` — one readable text file per accepted PDF.
- `.tmp_batch_meta.json` — counts, renames, duplicates, rejects.

The command prints one JSON object. Read it:

- `status` must be `"ok"`.
- If `accepted` is `0`, report the duplicates/rejects to the user and **stop**.
  Do not create workbooks, do not run `assemble`.
- If `status` is `"error"`, report `error` and stop.

### Step 2 — Extract

Read `references/sop-rules.md` and `references/extraction-schema.json` once.
Then, for **each** file in `acceptedFiles`:

1. Read that paper's `.tmp_text/<paper>.txt` (or its entry in
   `.tmp_batch_text.json`). Use the first page for metadata and the full text for
   the LR summary.
2. Extract exactly one metadata record (87 keys) and one LR record (19 keys).
   Do not include `S. No.`; `assemble` assigns it.
3. Obey the author, DOI, enum, and word-count rules in `references/sop-rules.md`.

To get a correctly-shaped blank record for a filename:

```bash
python "<SKILL_DIR>/scripts/pipeline.py" template --file "exact-name.pdf"
```

Write all records to `<PROJECT_ROOT>/.tmp_extractions.json` as a JSON array:

```json
[
  {
    "fileName": "exact-name.pdf",
    "metadata": { "...all 87 keys...": "" },
    "lrMatrix": { "...all 19 non-serial keys...": "" }
  }
]
```

The writer owns these values and will enforce them: `File Name`,
`File Attachments`, `Date Added`, `Date Modified`, `Access Date`, `Language`
(when blank), and `S. No.`.

Hard requirements:

- Every key from both schemas must be present; extra keys are rejected.
- `metadata.Item Type` ∈ `journalArticle`, `report`, `preprint`,
  `conferencePaper`, `thesis`.
- `lrMatrix.Tag 2 (Type of Research)` ∈ `Quantitative`, `Qualitative`, `Mixed`,
  `Review-Theoretical`, `Reference`.
- Metadata DOI is bare `10.xxxx`; LR DOI is `https://doi.org/10.xxxx`; both may
  be `""`.
- `Num Pages` is an integer ≥ 0; `Number Of Volumes` is an integer ≥ 1.
- `Title` is never blank.
- Word bands: Objective 40–70, Methodology 60–100, Key Finding 60–100,
  Discussion 40–70, Limitation 30–60.
- One record per accepted file, using the exact filename. No missing or extra
  records.

### Step 3 — Assemble

```bash
python "<SKILL_DIR>/scripts/pipeline.py" assemble \
  --root "<PROJECT_ROOT>" \
  --input "<PROJECT_ROOT>/.tmp_extractions.json"
```

The command validates the whole payload **before touching any workbook**. On
success it appends to `LR Matrix.xlsx` and the per-day metadata batch, moves the
keeper PDFs into `Database/`, and prints JSON with `rowsAdded`, `startSerial`,
`endSerial`, and `warnings`.

- If `status` is `"error"`, fix the listed `issues` (schema paths are exact) and
  rerun. No Excel rows are written on validation failure.
- If it says the files are already in `Database/`, the batch was already
  assembled: run `verify` instead of assembling again.
- Treat `warnings` as real SOP deviations. Fix out-of-band word counts before
  finalising when the fix is substantive; never pad with filler.

### Step 4 — Verify

```bash
python "<SKILL_DIR>/scripts/pipeline.py" verify --root "<PROJECT_ROOT>"
```

`verify` checks the LR header, Database↔LR bijection, row counts, serial
continuity, duplicate filenames, and blank titles, then appends a timestamped
entry to `SOP-Log.md`.

- Exit code `0` and `"passed": true` → success.
- Exit code `1` or `"passed": false` → read `issues`, fix the database, and run
  `verify` again. Do **not** delete temp files yet.
- If the problem is in the Excel data, repair `LR Matrix.xlsx` (or restore the
  `.bak` left by a failed assemble), then re-run `verify`.

### Step 5 — Clean up and report

Only after `verify` exits `0`, delete the batch temp artefacts:

```
.tmp_batch_text.json
.tmp_extractions.json
.tmp_batch_meta.json
.tmp_assemble_state.json
.tmp_text/
```

Then report to the user:

- PDFs seen, accepted, duplicates, rejects.
- Rows added and the S. No. range.
- Workbook paths.
- Verification result.
- Any warnings or anomalies (year-vs-filename mismatch, missing DOI, etc.).

## Command reference

| Command | Purpose | Success |
|---|---|---|
| `prepare --root ROOT` | Normalise, de-duplicate, extract text | `accepted` ≥ 0 |
| `template --file NAME` | Print a blank record with all keys | `status: ok` |
| `assemble --root ROOT --input FILE` | Validate, append rows, move PDFs | `rowsAdded` and serial range |
| `verify --root ROOT` | Audit and log | `passed: true`, exit 0 |

All commands print exactly one JSON object to stdout. Parse it; do not scrape
human-readable text.

## Failure rules

- **Zero accepted PDFs:** stop. Report duplicates/rejects. Do not create empty
  workbooks or invent rows.
- **Scanned or password-protected PDF:** `prepare` moves it to `Set-Aside/` with
  a reason. Report it; OCR is out of scope.
- **Schema rejection:** fix the exact JSON path named in `issues`. Common causes
  are a missing key, a quoted integer, a bad enum, a full DOI in the metadata
  record, or a blank title.
- **Excel file lock:** close the workbook in Excel and rerun the same command.
  Writes are atomic, so the previous state is preserved on failure.
- **Existing serial gap or row/file mismatch:** `assemble` refuses to append.
  Run `verify`, repair the existing rows, then resume.
- **Current batch already assembled:** do not rerun `assemble`. Run `verify`;
  the payload and Incoming files are consumed.

## What this skill deliberately does not do

- No screening or exclusion unless the user asks for it.
- No OCR; scanned PDFs are set aside.
- No LLM API calls, model selection, API keys, or network services.
- No changes to `Meta Data Filled.xlsx` or to existing LR rows.
