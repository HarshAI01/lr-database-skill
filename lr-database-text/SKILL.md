---
name: lr-database-text
description: >
  Dependency-free, text-only Literature Review Database skill. Use whenever the
  user wants to turn research papers into a literature review matrix, extract
  paper metadata, build an evidence table, de-duplicate a PDF batch, or audit an
  existing literature database, especially without Python, scripts, or installs,
  or inside a chat window. Produces paste-ready TSV rows for Excel, a batch
  report, a verification checklist, and SOP-Log entries. Trigger phrases:
  "literature review database", "LR database", "LR matrix", "literature matrix",
  "research matrix", "evidence table", "paper summary table", "systematic review
  table", "extract papers", "extract metadata", "metadata batch", "Zotero
  metadata", "87 metadata fields", "20 LR columns", "dedupe PDFs", "duplicate
  DOI", "SOP log", "verify literature database", "text-only literature review",
  "chat-only literature review", "no Python", "no scripts", "literature review
  Excel".
license: MIT
version: 1.0.0
---

# LR Database — Text-Only

Run the Literature Review Database SOP using only reading, reasoning, and text.
No Python, packages, scripts, network calls, or LLM APIs. If you have file
tools, you may save UTF-8 `.tsv` and `.md` files; otherwise print everything in
chat. **Never claim to have created or edited an `.xlsx` file.** You produce
rows the user pastes into their existing SOP workbooks.

## Using this skill

- **In chat:** paste this file as the first message, attach the PDFs, then answer
  the project-state question. Claude, ChatGPT, and Gemini all work.
- **As an installed skill:** invoke `/lr-database-text` or ask naturally ("add
  these papers to my LR database"). The skill is one folder with `SKILL.md`; no
  dependencies are installed.
- **In a file-capable agent:** save the TSV blocks as UTF-8 `.tsv` files beside
  the project if the user prefers files over copy-paste. Never generate `.xlsx`.

## Hard rules

1. **Never invent.** No authors, years, DOIs, journals, statistics, page
   numbers, or findings. If the PDF does not give it, use `""`.
2. **Keep-all.** `Decision (For Me)` and `Reason (For Me)` stay empty unless the
   user explicitly asks for screening against a written scope.
3. **One row per accepted PDF.** No merged rows, no extra rows.
4. **Never renumber existing rows.** New `S. No.` starts at `LAST SERIAL + 1`.
5. **Never modify `Meta Data Filled.xlsx`.** It is a frozen baseline.
6. **Never silently drop a PDF.** Every uploaded file appears in the batch
   report as `ACCEPTED`, `DUPLICATE`, `REVIEW`, or `SET ASIDE`, with a reason.
7. **One batch at a time.** Finish it before accepting another.
8. **Unreadable, scanned, or locked PDFs are set aside.** Do not guess contents.
   OCR is out of scope.
9. **The PDF is the evidence.** Do not fill gaps from general knowledge.
10. **Tabs are data.** Inside a TSV block, separate fields with a literal Tab
    (U+0009). Never spaces or pipes. Never wrap values in quotes.

## What not to do (read this twice)

These are the failures that break real databases. Treat each one as a hard
error, not a preference.

**State and numbering**

- Do not extract before you have `LAST SERIAL` and `EXISTING RECORDS`, unless
  the user said `FIRST RUN`.
- Do not start at `1` when `LAST SERIAL` is greater than `0`.
- Do not leave a gap, duplicate a serial, or sort serials by title or author.
- Do not renumber, reorder, delete, or rewrite rows that already exist.
- Do not append when the existing sheet has a gap or a filename mismatch. Stop
  and report it.

**Rows and duplicates**

- Do not create a row for a duplicate, an unreadable PDF, or a non-PDF file.
- Do not silently drop a file. Every file gets `ACCEPTED`, `DUPLICATE`,
  `REVIEW`, or `SET ASIDE`.
- Do not merge two papers into one row, even when they are similar.
- Do not keep both copies of a duplicate "just in case".

**Field formats**

- Do not put `https://doi.org/...` in metadata `DOI`; metadata is bare `10.…`.
- Do not put a bare `10.…` in the LR `DOI`; LR is the full link.
- Do not use the LR author format in metadata or the metadata format in LR.
- Do not write `et al.` when the paper lists the authors.
- Do not invent initials, affiliations, or corresponding-author details.
- Do not put a full date in `Year` or a year in `Date`.
- Do not write `NA`, `N/A`, `-`, or `NULL`; an unknown field is `""`.
- Do not fill Zotero role fields (42–87) unless the PDF explicitly assigns the
  role.

**Content**

- Do not paraphrase statistics loosely. Copy the reported coefficient, sample,
  period, and significance exactly.
- Do not add citations, page references, or "according to the authors" inside
  summary fields.
- Do not pad a summary with filler to reach a word band. Re-read the paper and
  replace filler with substance.
- Do not use outside knowledge to complete a missing year, DOI, or finding.
- Do not screen or exclude papers by default. Keep-all.
- Do not guess a page count. If the interface hides it, use `0` and warn.

**Output mechanics**

- Do not omit columns, merge columns, or change the header order.
- Do not print a 19-column LR table; `S. No.` is always required.
- Do not put tabs, newlines, carriage returns, pipes, or quotes inside a value.
- Do not output more than 8 papers in one block without repeating the header.
- Do not mark a failed verification check as PASS.
- Do not claim you saved or edited an `.xlsx` file.
- Do not run OCR, browse for the paper, or call an external service.

## Step 0 — Get project state

If the user has not provided state, ask once, in one short message:

> Before I start, paste your project state (or say **FIRST RUN**):
> 1. **LAST SERIAL** — the last `S. No.` in LR Matrix.xlsx (`0` if none)
> 2. **EXISTING RECORDS** — copy the File Name, DOI, Title and Year columns
>    from LR Matrix.xlsx, one paper per line:
>    `filename | DOI | Title | Year`
> 3. Attach the PDFs for this batch.

Do not extract until the user answers or says FIRST RUN. If they upload an
LR Matrix workbook or paste the columns, read the values instead of asking.

Keep in working memory:

- `lastSerial` — integer, `0` on a first run.
- `existing` — filename, DOI, title, year for every current row.

State template the user can fill:

```text
=== PROJECT STATE ===
LAST SERIAL: 0
EXISTING RECORDS:
paper-one.pdf | 10.1234/abc | Title One | 2021
paper-two.pdf |  | Title Two | 2020
=== END PROJECT STATE ===
```

## Step 1 — Ingest and de-duplicate

Process PDFs in alphabetical filename order. For each file:

1. Confirm it is readable and note its page count.
2. Read page 1 for metadata and DOI candidates.
3. Read the full text for the LR summaries.
4. Index: normalized filename, DOI, title, first author, year, page count.

Normalize a filename for comparison: lowercase, drop `.pdf`, delete every
character that is not a letter or digit.

A paper is a **duplicate** when any test is true. Run tests in this order; the
first match wins.

**Against `existing` (the current Database):**

1. Normalized filename equals an existing filename.
2. A DOI on page 1 equals an existing DOI. Compare bare forms, ignoring
   `https://doi.org/` and letter case.
3. The title shares more than ~70% of its significant words (length > 3) with an
   existing title, **and** at least 3 title words appear on that paper's first
   page, **and** the first author or year matches. If the first-page text is
   unavailable for an old record, require a near-identical title plus the same
   first author or year.

**Within this batch:**

1. DOI match.
2. Normalized filename match.
3. Title token overlap above ~0.85.

For a within-batch duplicate, keep the **version of record**: the file whose
first page shows a journal, volume, issue, or page range; otherwise the file
with more pages; otherwise the first alphabetically. The other goes to
`SET ASIDE`.

Ambiguous near-duplicates go to `REVIEW`, not `SET ASIDE`. Ask the user one
question and wait.

Non-PDF uploads are `SET ASIDE` with reason `Not a PDF`.

**If zero papers are accepted:** print the batch report and stop. Do not
extract, do not invent rows, do not continue to output tables.

### Batch report (always print)

| # | File | Status | Reason | Matched with |
|---:|---|---|---|---|
| 1 | paper.pdf | ACCEPTED | Unique | — |

## Step 2 — Extract

For every accepted paper, produce one metadata record (87 fields) and one LR
record (19 fields; `S. No.` is assigned by you, not extracted). Work in chunks
of at most 8 papers to avoid truncation; keep serials continuous across chunks.

### Field rules

- **Metadata authors** — printed order, `Last, First`, semicolon-separated:
  `Sharma, Pankaj; Patnaik, Tirthankar`. Single-token names stay as-is.
  Corporate author: the organisation goes in `Author`.
- **LR authors** — `First Last`, comma-separated: `Pankaj Sharma, Tirthankar Patnaik`.
- **Year/date** — as printed. If the PDF contradicts the filename, use the PDF
  and note the anomaly in the report.
- **DOI metadata** — bare `10.…`, strip `https://doi.org/` and
  `http://dx.doi.org/`.
- **DOI LR** — full `https://doi.org/…`, or `""`.
- **URL / ISSN / ISBN** — only if printed; bare numbers for ISSN/ISBN.
- **Abstract Note** — verbatim from the labelled abstract, max ~1200 characters.
- **Objective** — end of the introduction: questions, hypotheses, motivation.
  **40–70 words.**
- **Methodology** — sample, period, variables, techniques, reported statistics.
  **60–100 words.**
- **Key Finding** — results and abstract; copy reported numbers exactly.
  **60–100 words.**
- **Discussion** — discussion/conclusion. **40–70 words.**
- **Limitation** — stated limits; if none, the scope boundary. **30–60 words.**
- **Theoretical Framework** — only frameworks the paper names or clearly applies.
- **Variables** — only variables the paper names; independent, dependent, other
  (controls, moderators, mediators).
- **Item Type** — `journalArticle`, `report`, `preprint`, `conferencePaper`, or
  `thesis`. ArXiv-only or IEEE format without a printed venue = `preprint`.
- **Num Pages** — integer page count; use `0` only if the interface hides it and
  list it as a warning.
- **Number Of Volumes** — integer ≥ 1; default `1`.
- **Language** — language code; default `en` for English.
- **File Attachments** — `Database/<exact filename>`.
- **File Name** — exact basename, no path.
- **Date Added / Date Modified** — ISO timestamp; **Access Date** — ISO date.
  Use today's date if known, otherwise leave blank and warn.
- **Decision / Reason** — `""` by default.
- Metadata fields 42–87 are role fields. Leave them blank unless the PDF
  explicitly assigns that role. Blank is correct; invention is not.

### Exact metadata header — 87 columns, this order

Copy this line and keep the tabs. One row per accepted paper.

```text
File Name	Item Type	Publication Year	Author	Title	Publication Title	ISBN	ISSN	DOI	Url	Abstract Note	Date	Date Added	Date Modified	Access Date	Pages	Num Pages	Issue	Volume	Number Of Volumes	Journal Abbreviation	Short Title	Series	Series Number	Series Text	Series Title	Publisher	Place	Language	Rights	Type	Archive	Archive Location	Library Catalog	Call Number	Extra	Notes	File Attachments	Link Attachments	Manual Tags	Automatic Tags	Editor	Series Editor	Translator	Contributor	Attorney Agent	Book Author	Cast Member	Commenter	Composer	Cosponsor	Counsel	Interviewer	Producer	Recipient	Reviewed Author	Scriptwriter	Words By	Guest	Number	Edition	Running Time	Scale	Medium	Artwork Size	Filing Date	Application Number	Assignee	Issuing Authority	Country	Meeting Name	Conference Name	Court	References	Reporter	Legal Status	Priority Numbers	Programming Language	Version	System	Code	Code Number	Section	Session	Committee	History	Legislative Body
```

### Exact LR header — 20 columns, this order

`S. No.` is yours: start at `lastSerial + 1`, increment by 1 for each accepted
paper. Never leave a gap. Use `lastSerial + 1 … lastSerial + N`.

```text
S. No.	File Name	Title	Author	Year	Journal	DOI	Objective	Methodology	Key Finding	Discussion	Limitation	Theoretical Framework	Tag 1 (Method)	Tag 2 (Type of Research)	Independent Variable	Dependent Variable	Other Variable	Decision (For Me)	Reason (For Me)
```

`Tag 2 (Type of Research)` must be `Quantitative`, `Qualitative`, `Mixed`,
`Review-Theoretical`, or `Reference`. `Tag 1 (Method)` is free text, e.g.
`Panel regression`, `Difference-in-differences`, `Semi-structured interviews`.

## Step 3 — Output contract

Print these sections in this order. Do not stop early. Do not claim files were
saved unless you actually wrote them with a file tool.

### A. Batch report

The table from Step 1, plus counts: accepted, duplicates, review, set aside.

### B. LR Matrix rows — TSV

One fenced `text` block with the 20-column header and one line per accepted
paper, in ascending `S. No.` order. If there are more than 8 papers, use one
block per 8 with the header repeated and label them `chunk 1`, `chunk 2`, …

### C. Metadata rows — TSV

One fenced `text` block with the 87-column header and one line per accepted
paper in the same order. Chunk the same way if needed.

### D. Verification report

A table with one row per check:

| Check | Result | Detail |
|---|---|---|
| Headers exact | PASS | 20 and 87 columns |
| Serials continuous | PASS | 1–2, no gaps |
| One row per accepted PDF | PASS | 2 files, 2 rows |
| Filenames unique | PASS | — |
| Titles non-blank | PASS | — |
| Enums valid | PASS | — |
| DOI formats | PASS | bare + full link |
| Author formats | PASS | — |
| Word bands | PASS | Objective 55, Method 82, Finding 74, Discussion 51, Limitation 38 |
| Decision/Reason blank | PASS | keep-all |

Include the word counts for all five summary fields for every paper. If any
check fails, fix the rows and re-run the checks before printing the final
output. Never mark a failing check as PASS.

### E. SOP log entry

Print a block the user can append to `SOP-Log.md`:

```markdown
## YYYY-MM-DD HH:MM — Text-only batch — PASSED
- PDFs seen: 3
- Accepted: 2 | Duplicates: 1 | Set aside: 0 | Review: 0
- Rows added: 2 (S. No. 1–2)
- Warnings: none
```

### F. Handoff

Then give these steps, numbered and short:

1. Open `Outputs/LR Matrix.xlsx`, click the first empty cell in column A after
   the last row, and paste the LR TSV block. Excel splits on tabs.
2. Open `Outputs/metadata batch YYYY-MM-DD.xlsx`; paste the metadata block at
   `A1` in a new file or the first empty row in an existing same-day file.
3. Save both workbooks. Do not change existing rows.
4. Append the log entry to `SOP-Log.md`.
5. Optionally run the Python `verify` command if that skill is installed.

## Worked example (shape only)

Synthetic paper `sharma-2024-climate.pdf`, `FIRST RUN`.

Metadata — show only non-empty fields; every other field is `""`:

| Field | Value |
|---|---|
| File Name | `sharma-2024-climate.pdf` |
| Item Type | `journalArticle` |
| Publication Year | `2024` |
| Author | `Sharma, Pankaj; Patnaik, Tirthankar` |
| Title | `Climate Risk Disclosure and Firm Value` |
| Publication Title | `Journal of Climate Finance` |
| DOI | `10.1234/jcf.2024.001` |
| Abstract Note | verbatim abstract, max ~1200 chars |
| Pages | `101–118` |
| Num Pages | `18` |
| Issue | `2` |
| Volume | `7` |
| Number Of Volumes | `1` |
| Language | `en` |
| File Attachments | `Database/sharma-2024-climate.pdf` |
| Date Added / Modified | `2026-11-22T10:00:00+05:30` |
| Access Date | `2026-11-22` |

LR row (serial 1): `Pankaj Sharma, Tirthankar Patnaik`; DOI
`https://doi.org/10.1234/jcf.2024.001`; Objective 55 words; Methodology 82;
Key Finding 74; Discussion 51; Limitation 38; Tag 2 `Quantitative`.

Common format mistakes:

| Field | Wrong | Right |
|---|---|---|
| Metadata DOI | `https://doi.org/10.1234/x` | `10.1234/x` |
| LR DOI | `10.1234/x` | `https://doi.org/10.1234/x` |
| Metadata Author | `Pankaj Sharma` | `Sharma, Pankaj` |
| LR Author | `Sharma, Pankaj` | `Pankaj Sharma` |
| Serial | omitted | `1`, then `2`, then `3` |
| Unknown field | `NA` | empty |

## TSV rules

- Literal Tab between fields; one paper per line.
- No tabs, newlines, or carriage returns inside a value. Replace them with one
  space.
- Empty field = nothing between two tabs. Do not write `""`, `NA`, `-`, or
  `NULL`.
- Integers are plain: `1`, `2024`, `12`. No thousands separators.
- Do not quote fields and do not add padding spaces.
- Keep the exact column order from the headers above; never omit a column.
- If a value contains a formula-like prefix such as `=`, keep the text as-is;
  the user's Excel will treat it as text after paste only if they format the
  column as text. Warn them if this happens.

## Audit mode

Use this when the user asks to check an existing database instead of adding
papers. Accept an uploaded workbook or pasted columns.

Check and report:

1. LR header equals the 20 columns in exact order.
2. `S. No.` is `1..N` in physical row order — no gaps, duplicates, or disorder.
3. `File Name` is present and unique on every row.
4. `Title` is never blank.
5. DOI formats: metadata bare, LR full link.
6. Author formats: metadata `Last, First; …`, LR `First Last, …`.
7. Word counts within bands.
8. `Decision`/`Reason` blank unless screening was requested.
9. If the user supplies the `Database/` filename list, check the 1:1 bijection
   between PDFs and LR rows, and that counts match.

Output `PASSED` or `FAILED`, a check table, a list of exact fixes, and a log
entry. Do not rewrite or renumber the user's rows; propose fixes only.

## Self-test (after editing this skill)

If `references/test-cases.md` is available, run TC1–TC5 against synthetic text.
They are QA fixtures: never write them into a real database. At minimum, check:

- **TC1 clean batch** → 2 accepted, serials `1–2`, every check PASS.
- **TC2 duplicate DOI** → 0 accepted, batch report only, no TSV tables.
- **TC3 serial gap** → stop before extraction, no append, gap named.
- **TC4 scanned PDF** → set aside, remaining papers still processed.
- **TC5 TSV integrity** → 20 and 87 fields per row, no embedded tabs.

## Edge cases

| Case | Action |
|---|---|
| Scanned or image-only PDF | `SET ASIDE` — no extractable text, OCR required |
| Password-protected PDF | `SET ASIDE` — locked |
| DOI appears only in the reference list | Ignore it; use the article DOI on page 1 |
| Corporate author | Organisation in `Author`; invent no person |
| Signatory on a circular | Not an author |
| PDF year differs from filename | Use the PDF; note the anomaly |
| Generic title words overlap | Require first-page or author/year confirmation |
| Same-day second batch | Append to the same `metadata batch YYYY-MM-DD.xlsx` |
| Existing serials have gaps | Stop; report the gap; do not append |
| More than ~15 papers | Sub-batch: ingest → extract → output, then continue serials |
| User asks to exclude papers | Only then fill Decision/Reason |
| One field is unknown | `""` and a warning; never guess |

## Failure rules

- **No accepted PDFs:** report and stop.
- **Unreadable PDF:** set it aside, keep processing the rest.
- **Missing project state:** ask once, then wait.
- **Ambiguous duplicate:** mark `REVIEW` and ask one question.
- **Output would be truncated:** chunk at 8 papers and repeat the header.
- **User asks for an `.xlsx` file:** explain that this skill produces TSV rows
  for their existing workbooks; use the Python `lr-database` skill if they can
  install it.

## What this skill does not do

- No Python, Node, packages, scripts, or `.xlsx` generation.
- No OCR, screening, or exclusion unless explicitly requested.
- No edits to `Meta Data Filled.xlsx` or existing LR rows.
- No claims that a file was saved when only chat text was produced.
