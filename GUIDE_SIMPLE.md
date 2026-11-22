# LR Database Skill - Simple User Guide

## What Is This?

This tool helps you organize research papers automatically. You drop PDFs in a folder, and it:

1. **Removes duplicates** - Finds papers you already have
2. **Reads each paper** - Extracts text from PDFs
3. **Asks AI to summarize** - Gets key info: authors, year, journal, findings, methods
4. **Creates Excel files** - Perfectly formatted for literature reviews
5. **Logs everything** - Keeps a record of what was processed

**Perfect for**: Researchers, PhD students, literature reviews, systematic reviews

---

## Quick Start (3 Steps)

### Step 1: Install (One Time)

**Windows:**
1. Install Python from python.org (check "Add to PATH")
2. Open Command Prompt and run:
   ```
   pip install openpyxl fpdf2
   ```

**Mac:**
1. Open Terminal
2. Run:
   ```
   brew install python
   pip3 install openpyxl fpdf2
   ```

**Linux:**
```
sudo apt install python3 python3-pip
pip3 install openpyxl fpdf2
```

### Step 2: Get the Tool

Download from GitHub:
- Go to: https://github.com/HarshAI01/lr-database-skill
- Click "Code" > "Download ZIP"
- Unzip to a folder (e.g., `Documents/lr-database-skill`)

### Step 3: Set Up Your Project

1. Create a project folder:
   ```
   mkdir my-research
   cd my-research
   ```

2. Create an "Incoming" folder:
   ```
   mkdir Incoming
   ```

3. Copy your PDF papers into `Incoming/`

---

## How to Run It

### Option A: Simple Command (Recommended)

Open terminal/command prompt in your project folder and run:

**Windows:**
```
python path\to\lr-database-skill\run.py
```

**Mac/Linux:**
```
python3 path/to/lr-database-skill/run.py
```

### Option B: Use the Python Script Directly

The tool will:
1. Scan `Incoming/` for PDFs
2. Check for duplicates against `Database/`
3. Show you each paper's text
4. Ask you to paste AI-extracted data (see below)
5. Create Excel files in `Outputs/`
6. Log everything in `SOP-Log.md`

---

## The AI Extraction Step (Important!)

When the tool reads a paper, it shows you the text and asks you to get structured data from an AI (ChatGPT, Claude, etc.).

### What to Do:

1. **Copy the prompt** the tool shows you
2. **Paste into your AI** (ChatGPT, Claude, Gemini, etc.) along with the paper text
3. **Copy the JSON result** the AI gives you
4. **Paste back into the tool** when it asks

### Example Prompt (Tool Provides This):

```
Extract from this paper:
- Authors (Last, First; Last, First)
- Title (exact)
- Year
- Journal
- DOI (bare format: 10.xxxx)
- Abstract (verbatim, max 1200 chars)
- Objective (40-70 words: research questions/hypotheses)
- Methodology (60-100 words: data, methods, sample, stats)
- Key Finding (60-100 words: specific numbers/results)
- Discussion (40-70 words)
- Limitation (30-60 words)
- Theory frameworks (if any)
- Research type (Quantitative/Qualitative/Mixed/Review)
- Variables (Independent, Dependent, Other)

Return as JSON with "metadata" (87 fields) and "lrMatrix" (20 fields).
```

### AI Response Format:
```json
{
  "metadata": { "File Name": "paper.pdf", "Author": "Smith, John; Doe, Jane", ... },
  "lrMatrix": { "File Name": "paper.pdf", "Title": "...", "Author": "John Smith, Jane Doe", ... }
}
```

---

## What You Get

### Excel Files (in `Outputs/`):

| File | Description |
|------|-------------|
| `LR Matrix.xlsx` | Master table with 20 columns (all papers) |
| `metadata batch YYYY-MM-DD.xlsx` | 87 columns for this batch |
| `Meta Data Filled.xlsx` | Frozen baseline (never changes) |

### Log File:
- `SOP-Log.md` - Human-readable history of every run

### Folder Structure Created:
```
your-project/
├── Incoming/          # Drop PDFs here
├── Database/          # Processed PDFs (auto)
├── Outputs/           # Excel files (auto)
├── Set-Aside/         # Duplicates (auto)
└── SOP-Log.md         # Run history (auto)
```

---

## Key Rules (SOP Standards)

| Field | Format |
|-------|--------|
| Authors (metadata) | "Smith, John; Doe, Jane" |
| Authors (LR table) | "John Smith, Jane Doe" |
| DOI (metadata) | "10.1016/j.xxx" (bare) |
| DOI (LR table) | "https://doi.org/10.1016/j.xxx" |
| Word counts | Objective: 40-70, Methods: 60-100, Findings: 60-100 |
| Decision/Reason | Leave BLANK (keep all papers by default) |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Python not found" | Reinstall Python, check "Add to PATH" on Windows |
| "Module not found" | Run `pip install openpyxl fpdf2` |
| "No PDFs found" | Put PDFs in `Incoming/` folder |
| "Permission denied" | Run terminal as Administrator (Windows) or use `sudo` (Mac/Linux) |
| "Duplicate not caught" | Check `Set-Aside/` folder; tool uses title + first page content |
| "Excel won't open" | Close Excel if open; file may be locked |

---

## Tips for Best Results

1. **Use good PDFs** - Searchable text works best (not scanned images)
2. **Name files clearly** - Helps with duplicate detection
3. **Batch similar papers** - Process 10-20 at a time
3. **Verify AI output** - Check the JSON before pasting back
4. **Keep SOP-Log.md** - It's your audit trail
5. **Backup `Outputs/`** - Before major changes

---

## What Makes This Different

- **No coding needed** - Just run the script
- **Works with any AI** - ChatGPT, Claude, Gemini, local models
- **Exact SOP formatting** - 10pt Calibri, dark blue headers, wrapped text, filters
- **Duplicate smart detection** - Checks title + first page content (DOI, authors)
- **Free & open source** - No subscriptions, no limits
- **Your data stays local** - Nothing uploaded anywhere

---

## Need Help?

- **GitHub**: https://github.com/HarshAI01/lr-database-skill
- **Issues**: Report bugs on GitHub Issues
- **SOP Reference**: See `SOP.md` in the download

---

**Version 0.1.0** | **SOP 2026-10** | **Free (MIT License)**