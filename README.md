# LR Database — Literature Review Database Pipeline

A **free, open-source** universal toolkit that fully automates the **Literature Review Database SOP** (Standard Operating Procedure). Processes PDF batches through a 6-step pipeline: Ingest → Deduplicate → Extract (LLM) → Assemble (Excel) → Verify → Log.

**Works on ANY platform**: OpenCode, Claude Desktop (MCP), Cursor, VS Code, Claude Code, Standalone CLI, Docker.

## Features

- **Complete SOP Automation**: All 6 steps from `SOP.md` implemented faithfully
- **Model-Agnostic**: Works with any LLM (OpenAI, Anthropic, Gemini, Ollama, local, etc.) — no restrictions, no defaults
- **Excel Formatting Fidelity**: Preserves 10pt Calibri, wrapped text, top-aligned, thin borders, auto-filter via Python/openpyxl
- **Smart Deduplication**: Title token Jaccard + first-page content comparison (DOI, title/author overlap)
- **No External Templates**: Creates SOP-compliant Excel files automatically
- **Multi-Platform**: Native adapters for OpenCode, Claude MCP, Cursor, CLI, Docker
- **Free & Open Source**: MIT licensed, no limits, no tiers

---

## Quick Start (Universal)

### One-Line Install (All Platforms)

```bash
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
```

This detects your platform and installs the appropriate adapter.

---

## Platform-Specific Installation

### 1. OpenCode (Native Skill)

```bash
# Auto-install
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
# Select "OpenCode"

# Or manual
git clone https://github.com/HarshAI01/lr-database-skill ~/.opencode/skill/lr-database
cd ~/.opencode/skill/lr-database && npm install

# Use
opencode run lr-batch
```

### 2. Claude Desktop (MCP Server)

```bash
# Auto-install
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
# Select "Claude Desktop"

# Or manual
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

**Tools available in Claude**: `lr_batch`, `lr_ingest`, `lr_dedup`, `lr_extract_pdfs`, `lr_assemble`, `lr_verify`, `lr_log`, `lr_get_extraction_prompt`

### 3. Cursor / VS Code

```bash
# Auto-install
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
# Select "Cursor"
```

Or manual - adds `.cursor/rules/lr-database.mdc` with commands.

**Commands in Cursor Terminal**:
```bash
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-batch.ts -r .
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-extract-pdfs.ts -r .
# ... etc
```

### 4. Standalone CLI (Any Terminal)

```bash
# Auto-install
curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash
# Select "CLI"

# Or manual
git clone https://github.com/HarshAI01/lr-database-skill
cd lr-database-skill/adapters/cli && npm install && npm link

# Use anywhere
lr-db batch -r /path/to/project
lr-db extract -r /path/to/project
lr-db assemble -r /path/to/project -i extractions.json
lr-db prompt  # Print extraction prompt for LLM
```

### 5. Docker

```bash
# Build
git clone https://github.com/HarshAI01/lr-database-skill
cd lr-database-skill/adapters/docker && docker build -t lr-database .

# Run
docker run -v /path/to/project:/project lr-database batch -r /project

# With Docker Compose
cd lr-database-skill/adapters/docker && docker-compose up
```

---

## How It Works (Universal Pipeline)

| Step | Tool | Description |
|------|------|-------------|
| 1. **Ingest** | `lr_ingest` / `lr-db ingest` | Scan `Incoming/`, normalize filenames |
| 2. **Dedup** | `lr_dedup` / `lr-db dedup` | Compare against `Database/` + within batch; move duplicates to `Set-Aside/` |
| 3. **Extract** | `lr_extract_pdfs` / `lr-db extract` | Read PDFs, return content + extraction prompt for **your LLM** |
| 4. **Assemble** | `lr_assemble` / `lr-db assemble` | Write extracted data to Excel with SOP formatting |
| 5. **Verify** | `lr_verify` / `lr-db verify` | Check counts, bijection, serials, spot-checks |
| 6. **Log** | `lr_log` / `lr-db log` | Append entry to `SOP-Log.md` |

**Key Design**: The toolkit handles **file operations & Excel formatting**. **Your LLM does the extraction** — no model lock-in.

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

**No LLM configuration** — uses your platform's model automatically.

---

## Requirements

- Node.js 20+
- Python 3.10+ with `openpyxl` (installed via `npm install`)
- Any LLM configured in your platform (OpenCode, Cursor, Claude, etc.)

---

## Project Structure

```
lr-database-skill/
├── core/                      # Universal TypeScript/Python pipeline
│   ├── src/
│   │   ├── schemas.ts         # Zod schemas (87 + 20 fields)
│   │   ├── pipeline.ts        # Core orchestrator
│   │   └── index.ts
│   ├── package.json
│   └── tsconfig.json
├── python/                    # Excel formatting (openpyxl)
│   ├── create_templates.py    # Creates SOP templates
│   ├── assemble_excel.py      # Appends with SOP formatting
│   ├── get_last_serial.py     # Helper
│   └── requirements.txt
├── adapters/                  # Platform-specific wrappers
│   ├── opencode/              # OpenCode skill (skill.json, tools, workflows)
│   ├── claude-mcp/            # MCP server for Claude Desktop
│   ├── cursor/                # Cursor rules (.mdc)
│   ├── cli/                   # Standalone CLI (npm link)
│   └── docker/                # Dockerfile + docker-compose
├── install-universal.sh       # Universal installer
└── README.md
```

---

## License

MIT — Free to use, modify, distribute.

---

**Version**: 0.1.0 | **SOP Version**: 2026-10