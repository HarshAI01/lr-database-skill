# LR Database Skill — Literature Review Database Pipeline

A **free, open-source** OpenCode skill that fully automates the **Literature Review Database SOP** (Standard Operating Procedure). Processes PDF batches through a 6-step pipeline: Ingest → Deduplicate → Extract (LLM) → Assemble (Excel) → Verify → Log.

## Features

- **Complete SOP Automation**: All 6 steps from `SOP.md` implemented faithfully
- **Model-Agnostic**: Works with any OpenCode model (OpenAI, Anthropic, Gemini, Ollama, local, etc.) — no restrictions, no defaults
- **Excel Formatting Fidelity**: Preserves 10pt Calibri, wrapped text, top-aligned, thin borders, auto-filter via Python/openpyxl
- **Smart Deduplication**: Title token Jaccard + first-page content comparison (DOI, title/author overlap)
- **No External Templates**: Creates SOP-compliant Excel files automatically
- **OpenCode Native**: Skill.json, tools, workflows, system prompt — install and run
- **Free & Open Source**: MIT licensed, no limits, no tiers

## Quick Start

### 1. Install the Skill

```bash
# Copy the skill folder to your OpenCode skills directory
cp -r lr-database-skill ~/.opencode/skill/lr-database

# Install dependencies
cd ~/.opencode/skill/lr-database
npm install
```

### 2. Configure Your Model in OpenCode

OpenCode uses whatever model you have configured. Set up your preferred model:

- **OpenAI**: `opencode config set model openai/gpt-4o`
- **Anthropic**: `opencode config set model anthropic/claude-3-5-sonnet`
- **Gemini**: `opencode config set model gemini/gemini-1.5-pro`
- **Ollama (local)**: `opencode config set model ollama/llama3.1`
- **Any other**: Configure in OpenCode settings

### 3. Set Up Your Project

```bash
mkdir my-literature-review
cd my-literature-review

# Folders auto-created on first run
```

### 4. Add PDFs

```bash
mkdir -p Incoming
cp ~/papers/*.pdf Incoming/
```

### 5. Run

```bash
# Via OpenCode (recommended)
opencode run lr-batch
```

---

## How It Works

The skill provides **tools** for each pipeline step. The **agent (you/OpenCode)** uses its configured model to do the extraction:

| Step | Tool | What Happens |
|------|------|--------------|
| 1. Ingest | `lr-ingest` | Normalizes filenames in `Incoming/` |
| 2. Dedup | `lr-dedup` | Compares against `Database/` + within batch; moves duplicates to `Set-Aside/` |
| 3. Extract | `lr-extract` | Reads PDF text, returns content for **agent to extract** using its model |
| 4. Assemble | `lr-assemble` | Writes Excel with SOP formatting from agent-provided data |
| 5. Verify | `lr-verify` | Checks counts, bijection, serials, spot-checks |
| 6. Log | `lr-log` | Appends entry to `SOP-Log.md` |

**No LLM config in the skill** — the agent uses its own model for extraction.

---

## SOP Conventions Enforced

- **Authors**: Metadata = "Last, First; Last, First" | LR = "First Last, First Last"
- **DOI**: Metadata = bare "10.xxxx" | LR = full "https://doi.org/xxxx"
- **Item Types**: journalArticle, report, preprint, conferencePaper, thesis
- **Research Types**: Quantitative, Qualitative, Mixed, Review-Theoretical, Reference
- **Word Counts**: Objective 40-70, Methodology 60-100, Key Finding 60-100, Discussion 40-70, Limitation 30-60
- **Keep-All Default**: `Decision`/`Reason` blank unless screening requested
- **Serial Numbers**: Continuous from 1, no gaps, no renumbering
- **Formatting**: 10pt Calibri, wrapped, top-aligned, thin borders, header = bold white on dark blue

---

## Folder Structure (Auto-Created)

```
<ROOT>/
├── Incoming/              # Drop new PDFs here
├── Database/              # Processed PDFs (source of truth, auto-managed)
├── Outputs/               # Excel workbooks (auto-created with SOP formatting)
│   ├── LR Matrix.xlsx                # Cumulative 20-column matrix
│   ├── Meta Data Filled.xlsx         # Baseline metadata (frozen, never appended)
│   └── metadata batch YYYY-MM-DD.xlsx  # Per-batch metadata (87 cols)
├── Set-Aside/             # Duplicates & rejects (never tabled, auto-managed)
└── SOP-Log.md             # Run log (skill appends to it)
```

**No `Reference & Template/` folder required** — templates created programmatically.

---

## Configuration

```json
{
  "projectRoot": ".",
  "scope": "Your research topic (optional)",
  "pythonPath": "python"
}
```

| Option | Description | Default |
|--------|-------------|---------|
| `projectRoot` | Project folder | `.` |
| `scope` | Screening scope (optional) | `""` |
| `pythonPath` | Python executable | `python` |

**No LLM configuration** — uses your OpenCode model automatically.

---

## Requirements

- Node.js 20+
- Python 3.10+ with `openpyxl` (installed via `npm install`)
- OpenCode with any model configured

---

## License

MIT — Free to use, modify, distribute.

---

**Version**: 0.1.0 | **SOP Version**: 2026-10