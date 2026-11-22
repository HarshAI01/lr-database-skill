# LR Database Skill — Complete Usage & Installation Guide

## What This Does

**Literature Review Database Pipeline** — Automates the full SOP (Standard Operating Procedure) for managing a literature review database.

### 6-Step Pipeline

| Step | Action | What Happens |
|------|--------|--------------|
| **1. Ingest** | Scan `Incoming/` | Finds PDFs, normalizes filenames (removes illegal chars) |
| **2. Dedup** | Compare vs `Database/` | Title similarity + first-page content (DOI, authors); moves duplicates to `Set-Aside/` |
| **3. Extract** | Read PDFs | Returns PDF text + extraction prompt → **YOUR LLM extracts 87 metadata + 20 LR fields** |
| **4. Assemble** | Write Excel | Creates/updates `LR Matrix.xlsx` + batch metadata with exact SOP formatting |
| **5. Verify** | Quality check | Counts, file↔row bijection, serial continuity, blank titles, spot-check 2 PDFs |
| **6. Log** | Write `SOP-Log.md` | Appends structured entry with date, counts, duplicates, verification status |

### SOP Fields Produced

- **Metadata (87 columns)**: File Name, Item Type, Publication Year, Author, Title, Publication Title, ISBN, ISSN, DOI, URL, Abstract, Date, Dates, Pages, Num Pages, Issue, Volume, Journal Abbreviation, Short Title, Series, Publisher, Place, Language, Rights, Type, Archive, Library Catalog, Call Number, Extra, Notes, File Attachments, Link Attachments, Manual Tags, Automatic Tags, Editor, Series Editor, Translator, Contributor, Attorney Agent, Book Author, Cast Member, Commenter, Composer, Cosponsor, Counsel, Interviewer, Producer, Recipient, Reviewed Author, Scriptwriter, Words By, Guest, Number, Edition, Running Time, Scale, Medium, Artwork Size, Filing Date, Application Number, Assignee, Issuing Authority, Country, Meeting Name, Conference Name, Court, References, Reporter, Legal Status, Priority Numbers, Programming Language, Version, System, Code, Code Number, Section, Session, Committee, History, Legislative Body

- **LR Matrix (20 columns)**: S. No., File Name, Title, Author, Year, Journal, DOI, Objective, Methodology, Key Finding, Discussion, Limitation, Theoretical Framework, Tag 1 (Method), Tag 2 (Type of Research), Independent Variable, Dependent Variable, Other Variable, Decision (For Me), Reason (For Me)

---

## Installation

### Option 1: Universal Installer (Recommended)

```bash
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
```

Detects your platform and installs the right adapter.

### Option 2: Manual Install by Platform

#### OpenCode (Native Skill)
```bash
git clone https://github.com/HarshAI01/lr-database-skill ~/.opencode/skill/lr-database
cd ~/.opencode/skill/lr-database && npm install
# Use: opencode run lr-batch
```

#### Claude Desktop (MCP Server)
```bash
git clone https://github.com/HarshAI01/lr-database-skill ~/.config/claude/mcp/lr-database
cd ~/.config/claude/mcp/lr-database/adapters/claude-mcp && npm install
```
Add to `~/.config/claude/claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "lr-database": {
      "command": "node",
      "args": ["/home/user/.config/claude/mcp/lr-database/adapters/claude-mcp/server.js"],
      "env": { "OPENAI_API_KEY": "sk-..." }
    }
  }
}
```

#### Cursor / VS Code
```bash
git clone https://github.com/HarshAI01/lr-database-skill ~/.opencode/skill/lr-database
cd ~/.opencode/skill/lr-database && npm install
```
Adds `.cursor/rules/lr-database.mdc` with commands.

#### Standalone CLI (Any Terminal)
```bash
git clone https://github.com/HarshAI01/lr-database-skill
cd lr-database-skill/adapters/cli && npm install && npm link
# Use: lr-db batch -r /path/to/project
```

#### Docker
```bash
git clone https://github.com/HarshAI01/lr-database-skill
cd lr-database-skill/adapters/docker && docker build -t lr-database .
# Use: docker run -v /path/to/project:/project lr-database batch -r /project
```

---

## Quick Start

### 1. Set Up Project
```bash
mkdir my-literature-review
cd my-literature-review
```

### 2. Add PDFs
```bash
mkdir -p Incoming
cp ~/your-papers/*.pdf Incoming/
```

### 3. Run Pipeline

**OpenCode:**
```bash
opencode run lr-batch
```

**CLI:**
```bash
lr-db batch -r .
```

**Claude Desktop (MCP):**
In Claude chat: "Run lr_batch on my project"

**Cursor Terminal:**
```bash
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-batch.ts -r .
```

**Docker:**
```bash
docker run -v $(pwd):/project lr-database batch -r /project
```

---

## How Extraction Works (Key Concept)

**The toolkit handles file operations & Excel formatting. YOUR LLM does the extraction.**

### Step-by-Step Flow

1. **Tool reads PDFs** → Returns text content + extraction prompt
2. **You/your LLM processes** → Uses the prompt to extract 87+20 fields as JSON
3. **Tool assembles** → Writes to Excel with exact SOP formatting

### Example Extraction Prompt (provided by tool)

```
You are an expert academic metadata extractor...
Return JSON with "metadata" (87 fields) and "lrMatrix" (20 fields).
SOP conventions: Authors = "Last, First; Last, First" | LR = "First Last, First Last"
DOI metadata = bare "10.xxxx" | LR = full "https://doi.org/xxxx"
Word counts: Objective 40-70, Methodology 60-100, Key Finding 60-100...
```

### For CLI/OpenCode: Interactive Mode
```bash
# Tool shows PDF content, you paste JSON result
📄 Extracting: paper.pdf
FIRST PAGE: [content]
FULL TEXT: [content]
📋 EXTRACTION PROMPT: [prompt]
✅ Paste JSON result (or 'skip'):
```

### For Claude/Cursor: Chat-Based
1. Tool returns PDF content + prompt
2. You paste into chat with your model
3. Model returns JSON
4. You save as `extractions.json`
5. Run assemble step

---

## Folder Structure (Auto-Created)

```
your-project/
├── Incoming/              # ← Drop PDFs here
├── Database/              # Processed PDFs (auto-managed)
├── Outputs/               # Excel files (auto-created)
│   ├── LR Matrix.xlsx                # Cumulative 20-col matrix
│   ├── Meta Data Filled.xlsx         # Baseline (frozen)
│   └── metadata batch YYYY-MM-DD.xlsx  # Per-batch 87-col
├── Set-Aside/             # Duplicates (auto-managed)
└── SOP-Log.md             # Run log (auto-updated)
```

**No `Reference & Template/` folder needed** — templates created programmatically.

---

## Configuration

Create `config.json` in project root (optional):

```json
{
  "projectRoot": ".",
  "scope": "ESG and firm value in India; AI and ESG; sentiment to Nifty",
  "pythonPath": "python"
}
```

| Option | Description | Default |
|--------|-------------|---------|
| `projectRoot` | Project folder | `.` |
| `scope` | Screening scope (optional) | `""` |
| `pythonPath` | Python executable | `python` |

**No LLM configuration** — uses your platform's configured model.

---

## Requirements

- **Node.js 20+** — `node --version`
- **Python 3.10+** with `openpyxl` — `python --version` (installed via `npm install`)
- **Any LLM** configured in your platform (OpenAI, Anthropic, Gemini, Ollama, local, etc.)

---

## Commands Reference

### OpenCode Skill
```bash
opencode run lr-batch          # Full pipeline
opencode run lr-ingest         # Step 1 only
opencode run lr-dedup          # Step 2 only
opencode run lr-extract-pdfs   # Step 3 only
opencode run lr-assemble       # Step 4 only
opencode run lr-verify         # Step 5 only
opencode run lr-log            # Step 6 only
```

### CLI (lr-db)
```bash
lr-db batch -r .                    # Full pipeline
lr-db ingest -r .                   # Step 1
lr-db dedup -r .                    # Step 2
lr-db extract -r .                  # Step 3 (prints prompt + PDF content)
lr-db assemble -r . -i extractions.json  # Step 4
lr-db verify -r .                   # Step 5
lr-db log -r . -k 5 -d 0 --serial-start 1 --serial-end 5  # Step 6
lr-db prompt                        # Print extraction prompt
```

### Claude Desktop MCP Tools
- `lr_batch` — Full pipeline
- `lr_ingest` — Step 1
- `lr_dedup` — Step 2
- `lr_extract_pdfs` — Step 3 (returns content + prompt)
- `lr_assemble` — Step 4 (needs extractions array)
- `lr_verify` — Step 5
- `lr_log` — Step 6
- `lr_get_extraction_prompt` — Get prompt

---

## SOP Conventions Enforced

| Field | Convention |
|-------|------------|
| **Authors (metadata)** | "Last, First; Last, First" (semicolon, printed order) |
| **Authors (LR)** | "First Last, First Last" (comma, display form) |
| **DOI (metadata)** | Bare "10.xxxx" |
| **DOI (LR)** | Full "https://doi.org/xxxx" |
| **Item Types** | journalArticle, report, preprint, conferencePaper, thesis |
| **Research Types** | Quantitative, Qualitative, Mixed, Review-Theoretical, Reference |
| **Word Counts** | Objective 40-70, Methodology 60-100, Key Finding 60-100, Discussion 40-70, Limitation 30-60 |
| **Decision/Reason** | Blank (keep-all default) |
| **Serial Numbers** | Continuous from 1, no gaps |
| **Excel Format** | 10pt Calibri, wrapped, top-aligned, thin borders, dark blue headers (bold white), auto-filter |

---

## Troubleshooting

### "No PDFs in Incoming/"
- Ensure PDFs are in `Incoming/` folder (not subfolders)
- Check file extension is `.pdf`

### "Python not found"
- Install Python 3.10+: `python --version`
- Or set `pythonPath` in config to full path (e.g., `/usr/bin/python3`)

### "Ollama connection refused"
- Start Ollama: `ollama serve`
- Pull model: `ollama pull llama3.1`
- Or configure different model in your platform

### "Excel formatting wrong"
- Run `python python/create_templates.py Outputs` to regenerate templates
- Check `openpyxl` installed: `pip list | grep openpyxl`

### "Serial numbers not continuous"
- Run verify step: `lr-db verify -r .`
- Check `SOP-Log.md` for details

### "Duplicate not detected"
- Adjust thresholds in code: title Jaccard > 0.7, within-batch > 0.85
- First-page content comparison checks DOI + 3+ title words

---

## License

MIT — Free to use, modify, distribute.

---

## Support

- **Repo**: https://github.com/HarshAI01/lr-database-skill
- **Issues**: GitHub Issues
- **SOP Reference**: `SOP.md` in repo

---

**Version**: 0.1.0 | **SOP Version**: 2026-10