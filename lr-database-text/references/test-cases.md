# Test cases — lr-database-text

QA fixtures for this skill. Run them after any edit to `SKILL.md`. They are
synthetic: never write them into a real project database or reuse their DOIs.

## TC1 — Clean first batch

**State:** `FIRST RUN`.

**Input:** two readable papers.

| File | Title | First author | Year | DOI |
|---|---|---|---|---|
| `alpha-2024.pdf` | Alpha Study of Sustainable Finance | Alice Alpha | 2024 | `10.1234/alpha.2024` |
| `beta-2023.pdf` | Beta Evidence on Green Bonds | Carol Gamma | 2023 | `10.1234/beta.2023` |

**Expected**

- Batch report: both `ACCEPTED`, 0 duplicates, 0 set aside.
- LR serials `1, 2`, no gap.
- Metadata `DOI` is bare; LR `DOI` is `https://doi.org/...`.
- Metadata author `Alpha, Alice` and `Gamma, Carol`; LR author `Alice Alpha`
  and `Carol Gamma`.
- Five word counts inside their bands.
- Every verification check PASS. No `.xlsx` claimed.

## TC2 — Duplicate by DOI within the batch

**State:** `FIRST RUN`.

**Input:** `alpha-2024.pdf` and `alpha-working-paper.pdf` both show
`DOI: 10.1234/alpha.2024`; the first is the journal version with volume and
pages, the second is a working paper.

**Expected**

- One `ACCEPTED` (the journal version), one `DUPLICATE`.
- Exactly one LR row and one metadata row, serial `1`.
- The duplicate has no row and is listed as `SET ASIDE`.
- Reason names the shared DOI and the kept file.

## TC3 — Existing serial gap

**State:**

```text
LAST SERIAL: 5
EXISTING RECORDS:
one.pdf | 10.1234/one | One | 2020
two.pdf | 10.1234/two | Two | 2021
three.pdf | 10.1234/three | Three | 2022
four.pdf | 10.1234/four | Four | 2023
```

**Expected**

- Stop before extraction. `LAST SERIAL` says 5 but only 4 records exist.
- Report: gap at serial 5; the new batch would create a false `6`.
- No TSV rows, no append, no renumbering.

## TC4 — Scanned and non-PDF files

**State:** `FIRST RUN`.

**Input:** one readable PDF, one image-only PDF, one `.docx`.

**Expected**

- Readable PDF: `ACCEPTED`, one row, serial `1`.
- Image-only PDF: `SET ASIDE` — no extractable text, OCR required.
- `.docx`: `SET ASIDE` — not a PDF.
- The readable paper is still fully extracted; no guessing on the others.

## TC5 — TSV integrity and chunking

**State:** `FIRST RUN`, 9 accepted papers.

**Expected**

- LR output split into two blocks of at most 8 papers, each with the 20-column
  header repeated.
- Metadata output split the same way, each with the 87-column header repeated.
- Every LR row has exactly 20 tab-separated fields; every metadata row has 87.
- No value contains a tab, newline, pipe, or surrounding quotes.
- Serials run `1..9` across the chunks with no restart.

## Regression checklist

After TC1–TC5, confirm:

1. Description still names the trigger phrases and is under 1024 characters.
2. No tool or package is required.
3. Keep-all default is intact.
4. The seven output sections appear in order and none are skipped.
5. Failures are reported, never papered over.
