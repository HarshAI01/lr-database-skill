---
description: Audit an existing LR Matrix without adding papers
argument-hint: [folder]
---

Use the `lr-database-text` skill in audit mode.

TARGET: `$ARGUMENTS` if provided, otherwise the current working directory.

Find `Outputs/LR Matrix.xlsx` (or an `.xlsx`, `.csv`, or `.tsv` the user names).
If you cannot read the workbook without libraries, ask the user to upload it or
paste the 20 columns. Also list `Database/` PDF filenames when available.

Check:

1. LR header equals the 20 columns in exact order.
2. Serials are `1..N` with no gaps, duplicates, or disorder.
3. Every `File Name` is present and unique.
4. No blank `Title`.
5. Metadata DOI is bare; LR DOI is the full link.
6. Author formats are correct for each table.
7. Word counts are inside their bands.
8. `Decision` and `Reason` are blank unless screening was requested.
9. Database PDFs and LR rows are a 1:1 match when the file list is available.

Do not modify, renumber, or reorder the user's rows. Output `PASSED` or
`FAILED`, a check table, exact fixes, and an SOP log entry.
