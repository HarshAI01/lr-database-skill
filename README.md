# LR Database Skill — Literature Review Database Pipeline

A **free, open-source** OpenCode skill that fully automates the **Literature Review Database SOP** (Standard Operating Procedure). Processes PDF batches through a 6-step pipeline: Ingest → Deduplicate → Extract (LLM) → Assemble (Excel) → Verify → Log.

## Features

- **Complete SOP Automation**: All 6 steps from `SOP.md` implemented faithfully
- **Zero-Config Default**: Works instantly with local **Ollama** (no API keys, no cloud)
- **Multi-provider LLM**: Ollama (default), OpenAI, Anthropic, Gemini
- **Excel Formatting Fidelity**: Preserves 10pt Calibri, wrapped text, top-aligned, thin borders, auto-filter via Python/openpyxl
- **Smart Deduplication**: Title token Jaccard + first-page content comparison (DOI, title/author overlap)
- **No External Templates**: Creates SOP-compliant Excel files automatically
- **OpenCode Native**: Skill.json, tools, workflows, system prompt — install and run
- **Free & Open Source**: MIT licensed, no limits, no tiers

## Quick Start (Zero Config)

### 1. Install Ollama (one-time, 2 minutes)

```bash
# macOS
brew install ollama

# Linux
curl -fsSL https://ollama.com/install.sh | sh

# Windows
# Download from https://ollama.com/download
```

```bash
# Start Ollama service
ollama serve

# Pull a model (in another terminal)
ollama pull llama3.1
```

### 2. Install the Skill

```bash
# Copy the skill folder to your OpenCode skills directory
cp -r lr-database-skill ~/.opencode/skill/lr-database

# Install dependencies
cd ~/.opencode/skill/lr-database
npm install
```

### 3. Set Up Your Project

```bash
mkdir my-literature-review
cd my-literature-review

# That's it! No config needed - uses Ollama defaults
# Optional: create config.json to customize
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

# Or directly via CLI (uses Ollama localhost:11434 + llama3.1 by default)
npx tsx ~/.opencode/skill/lr-database/opencode/tools/lr-batch.ts -r .
```

---

## Alternative: Cloud LLMs (Optional)

If you prefer cloud providers, create `config.json`:

```json
{
  "projectRoot": ".",
  "scope": "Your research topic",
  "llm": {
    "provider": "openai",
    "model": "gpt-4o",
    "apiKey": "sk-..."
  },
  "pythonPath": "python"
}
```

Or via CLI flags:
```bash
# OpenAI
npx tsx lr-batch.ts -r . -p openai -m gpt-4o -k sk-...

# Anthropic
npx tsx lr-batch.ts -r . -p anthropic -m claude-3-5-sonnet -k sk-ant-...

# Gemini
npx tsx lr-batch.ts -r . -p gemini -m gemini-1.5-pro -k ...

# Ollama (custom model/url)
npx tsx lr-batch.ts -r . -p ollama -m mistral --base-url http://localhost:11434
```

## Pipeline Steps (SOP Compliant)

| Step | Tool | Description |
|------|------|-------------|
| 1. Ingest | `lr-ingest` | Scan `Incoming/`, normalize filenames (strip illegal chars) |
| 2. Dedup | `lr-dedup` | Compare against `Database/` + within batch; keep-best rule |
| 3. Extract | `lr-extract` | LLM extracts 87 metadata + 20 LR fields per PDF |
| 4. Assemble | `lr-assemble` | Append to Excel workbooks with SOP formatting |
| 5. Verify | `lr-verify` | Counts, bijection, serial continuity, spot-check 2 PDFs |
| 6. Log | `lr-log` | Append structured entry to `SOP-Log.md` |

**Full pipeline**: `lr-batch` (runs all 6 steps in order)

## SOP Conventions Enforced

- **Authors**: Metadata = "Last, First; Last, First" | LR = "First Last, First Last"
- **DOI**: Metadata = bare "10.xxxx" | LR = full "https://doi.org/xxxx"
- **Item Types**: journalArticle, report, preprint, conferencePaper, thesis
- **Research Types**: Quantitative, Qualitative, Mixed, Review-Theoretical, Reference
- **Word Counts**: Objective 40-70, Methodology 60-100, Key Finding 60-100, Discussion 40-70, Limitation 30-60
- **Keep-All Default**: `Decision`/`Reason` blank unless screening requested
- **Serial Numbers**: Continuous from 1, no gaps, no renumbering
- **Formatting**: 10pt Calibri, wrapped, top-aligned, thin borders, header = bold white on dark blue

## Configuration Options

| Option | Description | Default |
|--------|-------------|---------|
| `projectRoot` | Project folder containing Incoming/, Database/, etc. | `.` |
| `scope` | Screening scope paragraph (optional) | `""` |
| `llm.provider` | `ollama` \| `openai` \| `anthropic` \| `gemini` | `ollama` |
| `llm.model` | Model name (ollama: `llama3.1`, `mistral`; cloud: `gpt-4o`, etc.) | `llama3.1` |
| `llm.apiKey` | API key for cloud providers (not needed for ollama) | `""` |
| `llm.baseUrl` | Ollama base URL | `http://localhost:11434` |
| `pythonPath` | Python executable for Excel | `python` |

## Environment Variables (Cloud Providers Only)

```bash
export OPENAI_API_KEY="sk-..."
export ANTHROPIC_API_KEY="sk-ant-..."
export GEMINI_API_KEY="..."
```

**Ollama needs no API key or environment variables.**

## Output Files

- `Outputs/LR Matrix.xlsx` — Cumulative 20-column matrix (S.No. 1…N)
- `Outputs/metadata batch YYYY-MM-DD.xlsx` — Per-batch 87-column metadata
- `Outputs/Meta Data Filled.xlsx` — Baseline metadata (frozen, never appended)
- `SOP-Log.md` — Human-readable run log with duplicate details, verification status

## How It Works (Internally)

### 1. Ingest
- Scans `Incoming/` for PDFs
- Normalizes filenames (removes `<>:"|?*`, collapses whitespace)
- No files placed directly in `Database/`

### 2. Deduplicate
- Compares each incoming PDF against all files in `Database/`
- Uses title token Jaccard similarity (>0.7 threshold)
- For high-similarity pairs: compares first-page text (title, authors, journal, DOI)
- Keep-best rule: newer, cleaner text, richer metadata wins
- Within-batch duplicates also detected
- Duplicates moved to `Set-Aside/` with reason logged

### 3. Extract (LLM)
- Reads full PDF text (first 50k chars) + first page (3k chars)
- Sends to LLM with detailed SOP extraction prompt
- Returns structured JSON with 87 metadata fields + 20 LR matrix fields
- Validates against Zod schemas

### 4. Assemble (Excel via Python)
- Creates/updates `LR Matrix.xlsx` (cumulative, 20 cols)
- Creates per-batch `metadata batch YYYY-MM-DD.xlsx` (87 cols)
- Preserves SOP formatting: 10pt Calibri, wrapped, top-aligned, thin borders, dark blue headers
- Extends auto-filter to new rows
- Continuous serial numbers (S.No.)

### 5. Verify
- Record counts match keeper counts
- Every `Database/` file has a row and vice versa
- Serial numbers continuous from 1, no gaps
- No blank titles
- Spot-checks 2 PDFs against extracted data

### 6. Log
- Appends structured entry to `SOP-Log.md`
- Date, files in/out, duplicates with reasons, row ranges, anomalies

## Development

```bash
# Build TypeScript
npm run build

# Run in watch mode
npm run dev

# Run tests
npm test

# Lint
npm run lint
```

## Project Structure

```
lr-database-skill/
├── src/                    # TypeScript core pipeline
│   ├── schemas.ts          # Zod schemas for 87 + 20 fields
│   ├── pipeline/           # Orchestrator (ingest→dedup→extract→assemble→verify→log)
│   └── index.ts            # Main exports
├── python/                 # Excel formatting (openpyxl)
│   ├── create_templates.py # Creates SOP templates on first run
│   ├── assemble_excel.py   # Append with SOP formatting
│   └── get_last_serial.py  # Helper
├── opencode/               # OpenCode skill package
│   ├── skill.json          # Skill manifest
│   ├── tools/              # CLI tools (lr-batch, lr-ingest, etc.)
│   ├── workflows/          # YAML workflows
│   └── prompts/            # System prompt
├── schemas/                # JSON schemas (reference)
└── tests/                  # Test fixtures
```

## Requirements

- Node.js ≥ 20
- Python ≥ 3.10 with `openpyxl` (installed via `npm install`)
- LLM API key (OpenAI, Anthropic, or Gemini)

## License

MIT — Free to use, modify, distribute.

---

**Version**: 0.1.0 | **SOP Version**: 2026-10