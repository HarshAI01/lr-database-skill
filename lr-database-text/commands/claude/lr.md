---
description: Run the text-only LR Database workflow in any folder
argument-hint: [folder]
---

Use the `lr-database-text` skill for this run.

TARGET FOLDER
- If `$ARGUMENTS` is non-empty, treat it as the folder path (resolve relative
  paths against the current working directory).
- Otherwise use the current working directory.
- If the folder does not exist, stop and report the path. Do not create a
  different folder.

SETUP
Create `Incoming/`, `Database/`, `Outputs/`, `Set-Aside/`, and an empty
`SOP-Log.md` in the target if missing. Never overwrite an existing file.

BATCH
Look for PDFs in `Incoming/` first, then the target folder root. Do not recurse
into subfolders unless the user asks.

STATE
Ask once for `LAST SERIAL` and `EXISTING RECORDS`, or `FIRST RUN`. Do not
extract before you have the answer. If `Outputs/LR Matrix.xlsx` exists and you
can read it, use it instead of asking.

RUN
Follow the skill exactly: de-duplicate, extract, then output the batch report,
LR Matrix TSV, metadata TSV, verification table, and SOP log entry. When file
tools exist, save the TSVs under `Outputs/` with today's date and append the log
entry to `SOP-Log.md`.

If there are no PDFs, say so and stop. Never invent papers, rows, serials, or
metadata.
