# LR Database — SOP rules and schema reference

This file is the authority for every extraction, write, and audit. `SKILL.md`
controls the workflow; this file controls the data. When they disagree about a
field, this file wins.

---

## 1. Project layout

```
<PROJECT_ROOT>/
├── Incoming/                  # new PDFs only
├── Database/                  # processed PDFs, source of truth
├── Outputs/
│   ├── LR Matrix.xlsx         # cumulative, 20 columns, continuous S. No.
│   ├── Meta Data Filled.xlsx  # frozen baseline, 87 columns, never appended
│   └── metadata batch YYYY-MM-DD.xlsx
├── Set-Aside/                 # duplicates, rejects, non-PDFs
└── SOP-Log.md                 # human-readable audit trail
```

House rule: **keep-all**. Leave `Decision (For Me)` and `Reason (For Me)` blank
unless the user explicitly asks for screening against a written scope.

---

## 2. Deduplication rules (run before reading any paper)

Filenames are normalised first: remove `<>:"/\|?*` and control characters,
collapse whitespace, strip trailing dots/spaces, force the `.pdf` extension.
Reserved Windows names get a trailing `_`.

Run these tests in order. The first match wins.

### 2.1 Against `Database/`

| Order | Test | Threshold | Result |
|---|---|---|---|
| 1 | Normalised filename match | Casefolded, alphanumeric-only stem equality | Duplicate — reason `Normalized filename match` |
| 2 | DOI match | Any DOI in the incoming first page equals a DOI in a Database first page | Duplicate — reason `DOI match` |
| 3 | Title/first-page Jaccard | Filename-title token Jaccard **> 0.70** and at least **3** incoming title tokens (length > 4) appear in the Database first page | Duplicate — reason `Title/first-page match` |

A Database copy always wins. The incoming copy moves to `Set-Aside/`.

### 2.2 Within the incoming batch

| Order | Test | Threshold | Result |
|---|---|---|---|
| 1 | DOI match | Any DOI intersection | Keep the larger file; loser to `Set-Aside/` |
| 2 | Normalised filename match | Alphanumeric-only stem equality | Keep the larger file |
| 3 | Filename-title Jaccard | **> 0.85** | Keep the larger file |

Tie on size: keep the file encountered first (alphabetical order).

### 2.3 Tokenisation used for Jaccard

Unicode NFKC, casefold, extract word runs with no underscores, drop tokens of
length ≤ 3. Jaccard = |A ∩ B| / |A ∪ B|.

### 2.4 Never de-duplicate against

- DOI strings that appear only in a reference list (first page only is read).
- Title-only similarity without the first-page confirmation (Database test).
- Metadata text; the PDF is the evidence.

Unreadable, encrypted, or image-only PDFs (no extractable text) move to
`Set-Aside/` with an actionable reason. They are not silently dropped.

---

## 3. Extraction conventions (binding)

1. **PDF content only.** Official publisher/regulator pages are allowed only for
   institutional reports (e.g. SEBI circulars, NBER papers) when the PDF itself
   points there.
2. **Metadata authors** — printed order, `Last, First`, semicolon-separated:
   `Sharma, Pankaj; Patnaik, Tirthankar`.
   - Single-token names stay as-is.
   - Corporate report with no personal author: organisation goes in `Author`.
   - A signatory at the bottom of a circular is **not** an author.
3. **LR authors** — display form, `First Last`, comma-separated:
   `Pankaj Sharma, Tirthankar Patnaik`.
4. **Year/date** exactly as printed. If the PDF contradicts the filename, use the
   PDF and note the anomaly in the run report.
5. **DOI in metadata** — bare form beginning `10.`. Strip `https://doi.org/` and
   `http://dx.doi.org/`.
6. **DOI in LR matrix** — full `https://doi.org/...` link, or blank.
7. **URL** — only if printed in the PDF. **ISSN/ISBN** — bare numbers.
8. **Abstract** — verbatim from the labelled Abstract, max ~1200 characters.
9. **Objective** — end of the Introduction: research questions, hypotheses,
   motivation. **40–70 words.**
10. **Methodology** — data and methods: sample, period, variables, techniques,
    reported statistics. **60–100 words.**
11. **Key Finding** — results tables plus abstract. Copy reported numbers; never
    invent. **60–100 words.**
12. **Discussion** — discussion/conclusion. **40–70 words.**
13. **Limitation** — discussion/conclusion; if the paper states none, summarise
    the scope boundary honestly. **30–60 words.**
14. **Theory** — only frameworks the paper names or clearly applies; else blank.
15. **Variables** — only variables the paper names, split into independent,
    dependent, and other (control/moderator/mediator).
16. **Decision/Reason** — blank by default.
17. `Date Added`, `Date Modified`, `Access Date`, `Language`, `File Attachments`
    are repair defaults: the writer fills them when blank and always forces
    `File Attachments` to `Database/<exact filename>`.

---

## 4. Metadata schema — 87 columns, exact order

`Num Pages` and `Number Of Volumes` are JSON integers; every other field is a
JSON string. `Item Type` is an enum. All 87 keys are required; use `""` when the
PDF gives nothing.

| # | Column | Rule / source |
|---:|---|---|
| 1 | File Name | Exact PDF basename, no path. Writer enforces it. |
| 2 | Item Type | `journalArticle`, `report`, `preprint`, `conferencePaper`, or `thesis`. ArXiv-only = `preprint`; IEEE format without a printed venue = `preprint`. |
| 3 | Publication Year | Year as printed; may include month/day if printed that way. |
| 4 | Author | Metadata author format from §3.2. |
| 5 | Title | Exact printed title; never blank. |
| 6 | Publication Title | Journal or venue name. |
| 7 | ISBN | Bare ISBN, else `""`. |
| 8 | ISSN | Bare ISSN, else `""`. |
| 9 | DOI | Bare `10.…`, else `""`. |
| 10 | Url | Only if printed in the PDF. |
| 11 | Abstract Note | Verbatim abstract, max ~1200 characters. |
| 12 | Date | Printed date (year, year-month, or full date). |
| 13 | Date Added | ISO timestamp; writer fills if blank. |
| 14 | Date Modified | ISO timestamp; writer fills if blank. |
| 15 | Access Date | ISO date; writer fills if blank. |
| 16 | Pages | Article page range as printed, e.g. `101–118`. |
| 17 | Num Pages | Integer page count of the PDF. |
| 18 | Issue | Issue number. |
| 19 | Volume | Volume number. |
| 20 | Number Of Volumes | Integer ≥ 1; default `1`. |
| 21 | Journal Abbreviation | Only if printed or standard for the venue. |
| 22 | Short Title | Only if printed as a running title. |
| 23 | Series | Series name. |
| 24 | Series Number | Series number. |
| 25 | Series Text | Series text. |
| 26 | Series Title | Series title. |
| 27 | Publisher | Publisher or issuing body. |
| 28 | Place | Place of publication. |
| 29 | Language | Language code; writer defaults to `en` if blank. |
| 30 | Rights | Copyright/licence statement. |
| 31 | Type | Document subtype printed by the venue. |
| 32 | Archive | Archive name. |
| 33 | Archive Location | Archive location. |
| 34 | Library Catalog | Catalog name. |
| 35 | Call Number | Call number. |
| 36 | Extra | Free notes; do not use for data that belongs elsewhere. |
| 37 | Notes | Free notes. |
| 38 | File Attachments | `Database/<exact filename>`; writer enforces it. |
| 39 | Link Attachments | Attached URL, else `""`. |
| 40 | Manual Tags | Semicolon-separated manual tags. |
| 41 | Automatic Tags | Machine tags, else `""`. |
| 42 | Editor | Editors, metadata author format. |
| 43 | Series Editor | Series editors. |
| 44 | Translator | Translators. |
| 45 | Contributor | Named contributors. |
| 46 | Attorney Agent | Legal role, else `""`. |
| 47 | Book Author | Book author. |
| 48 | Cast Member | Cast member. |
| 49 | Commenter | Commenter. |
| 50 | Composer | Composer. |
| 51 | Cosponsor | Cosponsor. |
| 52 | Counsel | Counsel. |
| 53 | Interviewer | Interviewer. |
| 54 | Producer | Producer. |
| 55 | Recipient | Recipient. |
| 56 | Reviewed Author | Reviewed author. |
| 57 | Scriptwriter | Scriptwriter. |
| 58 | Words By | Lyricist. |
| 59 | Guest | Guest. |
| 60 | Number | Report/standard number. |
| 61 | Edition | Edition. |
| 62 | Running Time | Media running time. |
| 63 | Scale | Map scale. |
| 64 | Medium | Medium. |
| 65 | Artwork Size | Artwork size. |
| 66 | Filing Date | Legal filing date. |
| 67 | Application Number | Patent application number. |
| 68 | Assignee | Patent assignee. |
| 69 | Issuing Authority | Issuing authority. |
| 70 | Country | Country. |
| 71 | Meeting Name | Meeting name. |
| 72 | Conference Name | Conference name. |
| 73 | Court | Court. |
| 74 | References | Reference count or note, else `""`. |
| 75 | Reporter | Legal reporter. |
| 76 | Legal Status | Legal status. |
| 77 | Priority Numbers | Patent priority numbers. |
| 78 | Programming Language | Programming language. |
| 79 | Version | Software version. |
| 80 | System | Software system. |
| 81 | Code | Code identifier. |
| 82 | Code Number | Code number. |
| 83 | Section | Section. |
| 84 | Session | Session. |
| 85 | Committee | Committee. |
| 86 | History | Legislative history. |
| 87 | Legislative Body | Legislative body. |

Fields 42–87 are Zotero-style role fields. Leave them blank unless the PDF
explicitly assigns that role. Blank is correct; invention is not.

---

## 5. LR Matrix schema — 20 columns, exact order

| # | Column | Rule |
|---:|---|---|
| 1 | S. No. | Assigned by `assemble`; continuous `1..N`; never supplied by the extractor. |
| 2 | File Name | Exact PDF basename, same as metadata. |
| 3 | Title | Exact title, never blank. |
| 4 | Author | LR author format from §3.3. |
| 5 | Year | Year as printed. |
| 6 | Journal | Journal/venue. |
| 7 | DOI | Full `https://doi.org/...`, or `""`. |
| 8 | Objective | 40–70 words. |
| 9 | Methodology | 60–100 words with concrete specifics and reported statistics. |
| 10 | Key Finding | 60–100 words with concrete specifics and reported statistics. |
| 11 | Discussion | 40–70 words. |
| 12 | Limitation | 30–60 words. |
| 13 | Theoretical Framework | Named/applied frameworks only. |
| 14 | Tag 1 (Method) | Free-text method label, e.g. `Panel regression`, `Difference-in-differences`, `Semi-structured interviews`, `Systematic review`. |
| 15 | Tag 2 (Type of Research) | `Quantitative`, `Qualitative`, `Mixed`, `Review-Theoretical`, or `Reference`. |
| 16 | Independent Variable | As named in the paper. |
| 17 | Dependent Variable | As named in the paper. |
| 18 | Other Variable | Controls, moderators, mediators. |
| 19 | Decision (For Me) | `""` by default; `Include`/`Exclude` only for explicit screening. |
| 20 | Reason (For Me) | `""` by default. |

Word bands are binding. The writer reports out-of-band fields as warnings; fix
them rather than padding with filler.

---

## 6. Excel formatting (exact)

- Header row: Calibri **10 pt**, bold, white text, fill **`#1F4E78`**, wrapped,
  horizontally and vertically centred, thin borders.
- Data rows: Calibri **10 pt**, wrapped, vertically top-aligned, thin borders on
  every side. Do not set row heights; let Excel auto-size.
- Auto-filter: `A1:<last column letter><last data row>` on both workbooks.
- Column widths: use the widths in `scripts/create_templates.py`; never change
  them after creation.
- `Outputs/Meta Data Filled.xlsx` is a **frozen baseline**. Create it if missing,
  then never load, append, or overwrite it.
- `Outputs/LR Matrix.xlsx` is cumulative. Never renumber, delete, reorder, or
  rewrite existing rows. Append only.
- `Outputs/metadata batch YYYY-MM-DD.xlsx` is per-day. A second batch on the
  same day appends to the same file.
- All workbook writes are atomic: write a temp file in `Outputs/`, reopen and
  audit it, replace the target with `os.replace()`, and keep a `.bak` until the
  batch succeeds.
- Text is cleaned of XML-illegal control characters and capped at 32,000
  characters. Strings beginning with `=` are forced to text, never formulas.

---

## 7. Extraction JSON contract

`.tmp_extractions.json` is a JSON array (or an object with an `extractions`
array). Each item:

```json
{
  "fileName": "exact-name.pdf",
  "metadata": { "all 87 keys": "..." },
  "lrMatrix": { "all 19 non-serial keys": "..." }
}
```

Rules enforced by `scripts/assemble_excel.py`:

- All 87 metadata keys and all 19 LR keys are required.
- `additionalProperties` is false in both objects.
- `Item Type` and `Tag 2 (Type of Research)` enums are enforced.
- Metadata DOI must match `^(|10\.\S+)$`; LR DOI must match
  `^(|https://doi\.org/10\.\S+)$`.
- `Num Pages` is an integer ≥ 0; `Number Of Volumes` is an integer ≥ 1.
- `Title` and `File Name` are non-empty.
- `S. No.` is optional and ignored; `assemble` overwrites it.
- Every `fileName` must exist in `Incoming/` and must match a keeper from
  `.tmp_batch_text.json` exactly. Missing or extra records abort the write.
- Duplicate `fileName` values abort the write.

`Decision (For Me)` and `Reason (For Me)` must be present; use `""`.

---

## 8. Verification checklist

`verify` exits `0` only when all of these pass:

1. `Outputs/LR Matrix.xlsx` header equals the 20 columns in exact order.
2. Every Database PDF has exactly one LR row, and every LR row points to an
   existing Database PDF.
3. LR row count equals Database PDF count.
4. Serial numbers are `1..N` in physical row order: no gaps, duplicates, or
   scrambled values.
5. No duplicate `File Name` values in the LR matrix.
6. No blank `Title` values.
7. The audit entry is appended to `SOP-Log.md` with a timestamp and result.

An empty database (0 PDFs, no LR matrix) is a valid pass. A missing LR matrix
with PDFs in Database is a failure.

---

## 9. Edge cases

| Case | Action |
|---|---|
| Scanned/image-only PDF | Move to `Set-Aside/`; reason: no extractable text, OCR required. |
| Password-protected PDF | Try empty password; if still locked, move to `Set-Aside/`. |
| DOI only in the reference list | Ignore it; the article DOI must appear on the first page/header. |
| Corporate author | Organisation in `Author`; no personal name invented. |
| Circular signed by an official | Signatory is not an author. |
| PDF year differs from filename | Use the PDF; note the anomaly. |
| Accented/dashed filenames | NFKC normalise, casefold for comparison, keep a readable filename. |
| Two papers share generic title words | Jaccard alone is not enough; require first-page confirmation for the 0.70 Database rule. |
| Same-day second batch | Append to the same `metadata batch YYYY-MM-DD.xlsx`. |
| Excel has the workbook open | `os.replace` fails; close Excel and rerun `assemble`. |
| Existing LR serials have gaps | Refuse to append; run `verify`, repair the sheet, then rerun. |
| Metadata batch name collision | Append rows; never overwrite earlier same-day rows. |
| User asks to exclude papers | Only then fill Decision/Reason; default remains keep-all. |
| More than ~50 papers | Extract in smaller sub-batches; run prepare → assemble → verify per sub-batch. |
