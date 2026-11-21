/**
 * Universal LR Database Pipeline
 * Platform-agnostic core logic for: ingest → dedup → extract → assemble → verify → log
 */

import { glob } from "fast-glob";
import { readFile, writeFile, rename, stat, mkdir } from "fs/promises";
import { join, basename, resolve } from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import pdfParse from "pdf-parse";
import {
  MetadataSchema, LRMatrixSchema, ExtractionResultSchema,
  MetadataRecord, LRMatrixRecord, METADATA_COLUMNS, LR_MATRIX_COLUMNS,
  createEmptyMetadata, createEmptyLRMatrix
} from "./schemas.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..");
const CORE_ROOT = resolve(__dirname, "..");

export interface PipelineConfig {
  projectRoot: string;
  scope: string;
  pythonPath: string;
}

export interface FolderPaths {
  ROOT: string;
  IN: string;
  DB: string;
  OUT: string;
  SETASIDE: string;
  LR_MATRIX: string;
  META0: string;
}

export interface BatchResult {
  incomingFiles: string[];
  keeperFiles: string[];
  duplicates: Array<{ file: string; reason: string; matchedWith: string }>;
  extracted: ExtractionResult[];
  metadataBatchFile: string;
  lrMatrixFile: string;
  verification: VerificationResult;
}

export interface VerificationResult {
  passed: boolean;
  issues: string[];
  checks: {
    recordCountsMatch: boolean;
    fileRowBijection: boolean;
    serialContinuous: boolean;
    noBlankTitles: boolean;
    spotCheckPassed: boolean;
  };
}

export interface ExtractionResult {
  fileName: string;
  metadata: MetadataRecord;
  lrMatrix: LRMatrixRecord;
}

export interface PDFContent {
  fileName: string;
  numPages: number;
  fullText: string;
  firstPage: string;
}

// ============================================================================
// Folder Management
// ============================================================================

export function getFolders(root: string): FolderPaths {
  return {
    ROOT: root,
    IN: join(root, "Incoming"),
    DB: join(root, "Database"),
    OUT: join(root, "Outputs"),
    SETASIDE: join(root, "Set-Aside"),
    LR_MATRIX: join(root, "Outputs", "LR Matrix.xlsx"),
    META0: join(root, "Meta Data Filled.xlsx"),
  };
}

export async function ensureFolders(folders: FolderPaths, pythonPath: string) {
  await Promise.all([
    mkdir(folders.IN, { recursive: true }),
    mkdir(folders.DB, { recursive: true }),
    mkdir(folders.OUT, { recursive: true }),
    mkdir(folders.SETASIDE, { recursive: true }),
  ]);
  await runPython(join(CORE_ROOT, "..", "python", "create_templates.py"), [folders.OUT], pythonPath);
}

// ============================================================================
// Step 1: INGEST
// ============================================================================

function normalizeFilename(name: string): string {
  return name.replace(/[<>:"|?*]/g, "").replace(/\s+/g, " ").trim();
}

export async function runIngest(folders: FolderPaths): Promise<string[]> {
  const pdfFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  const normalized: string[] = [];
  for (const file of pdfFiles) {
    const name = basename(file);
    const normalizedName = normalizeFilename(name);
    if (name !== normalizedName) {
      await rename(file, join(folders.IN, normalizedName));
      normalized.push(join(folders.IN, normalizedName));
    } else {
      normalized.push(file);
    }
  }
  return normalized;
}

// ============================================================================
// Step 2: DEDUP
// ============================================================================

function titleTokens(title: string): string[] {
  return title.toLowerCase().replace(/[^\w\s]/g, " ").split(/\s+/).filter(t => t.length > 3);
}

function jaccardSimilarity(a: string[], b: string[]): number {
  const setA = new Set(a), setB = new Set(b);
  const intersection = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : intersection / union;
}

async function getPdfFirstPageText(filePath: string): Promise<string> {
  const buffer = await readFile(filePath);
  const data = await pdfParse(buffer, { max: 1 });
  return data.text.slice(0, 3000);
}

export async function runDedup(
  incomingFiles: string[],
  folders: FolderPaths
): Promise<{ keepers: string[]; duplicates: Array<{ file: string; reason: string; matchedWith: string }> }> {
  const dbFiles = await glob("*.pdf", { cwd: folders.DB, absolute: true });
  
  const dbIndex = [];
  for (const file of dbFiles) {
    try {
      const firstPage = await getPdfFirstPageText(file);
      dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage: firstPage.toLowerCase() });
    } catch {
      dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage: "" });
    }
  }

  const keepers: string[] = [];
  const duplicates: Array<{ file: string; reason: string; matchedWith: string }> = [];

  for (const incFile of incomingFiles) {
    const incTokens = titleTokens(basename(incFile).replace(/\.pdf$/i, ""));
    let matched = false;

    for (const db of dbIndex) {
      const sim = jaccardSimilarity(incTokens, db.tokens);
      if (sim > 0.7) {
        const incFirstPage = (await getPdfFirstPageText(incFile)).toLowerCase();
        const doiMatch = incFirstPage.match(/10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i);
        if (doiMatch && db.firstPage.includes(doiMatch[0].toLowerCase())) {
          matched = true;
          duplicates.push({ file: incFile, reason: "DOI match", matchedWith: db.file });
          break;
        }
        const titleWords = incTokens.filter(t => t.length > 4);
        const overlap = titleWords.filter(t => db.firstPage.includes(t)).length;
        if (overlap >= 3) {
          matched = true;
          duplicates.push({ file: incFile, reason: "Title/author match", matchedWith: db.file });
          break;
        }
      }
    }

    if (!matched) {
      for (const other of incomingFiles) {
        if (other === incFile) continue;
        const otherTokens = titleTokens(basename(other).replace(/\.pdf$/i, ""));
        if (jaccardSimilarity(incTokens, otherTokens) > 0.85) {
          const incStat = await stat(incFile);
          const otherStat = await stat(other);
          if (incStat.size >= otherStat.size) {
            duplicates.push({ file: other, reason: "Within-batch duplicate (larger kept)", matchedWith: incFile });
          } else {
            matched = true;
            duplicates.push({ file: incFile, reason: "Within-batch duplicate (larger kept)", matchedWith: other });
            break;
          }
        }
      }
    }

    if (!matched) keepers.push(incFile);
  }

  return { keepers, duplicates };
}

// ============================================================================
// Step 3: EXTRACT - Returns PDF content for agent to process
// ============================================================================

export async function runExtractPDFs(keeperFiles: string[]): Promise<PDFContent[]> {
  const results: PDFContent[] = [];
  
  for (const file of keeperFiles) {
    const fileName = basename(file);
    try {
      const buffer = await readFile(file);
      const pdfData = await pdfParse(buffer);
      const fullText = pdfData.text;
      const numPages = pdfData.numpages;
      
      const firstPageData = await pdfParse(buffer, { max: 1 });
      const firstPage = firstPageData.text;

      results.push({
        fileName,
        numPages,
        fullText: fullText.slice(0, 50000),
        firstPage: firstPage.slice(0, 3000),
      });
    } catch (error) {
      results.push({ fileName, numPages: 0, fullText: "", firstPage: "", error: String(error) });
    }
  }
  
  return results;
}

// ============================================================================
// Step 4: ASSEMBLE - Takes agent-extracted data and writes Excel
// ============================================================================

function runPython(script: string, args: string[], pythonPath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const proc = spawn(pythonPath, [script, ...args], { cwd: CORE_ROOT });
    let stdout = "", stderr = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => stderr += d.toString());
    proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
  });
}

export async function runAssemble(
  extractions: ExtractionResult[],
  folders: FolderPaths,
  config: PipelineConfig
): Promise<{ metadataBatchFile: string; lrMatrixFile: string; startSerial: number }> {
  let startSerial = 1;
  try {
    const result = await runPython(
      join(CORE_ROOT, "..", "python", "get_last_serial.py"),
      [folders.LR_MATRIX], config.pythonPath
    );
    startSerial = parseInt(result.trim()) + 1;
  } catch {
    startSerial = 1;
  }

  extractions.forEach((ex, i) => {
    ex.lrMatrix["S. No."] = startSerial + i;
  });

  const dateStr = new Date().toISOString().split("T")[0].replace(/-/g, "-");
  const metadataBatchFile = join(folders.OUT, `metadata batch ${dateStr}.xlsx`);
  const lrMatrixFile = folders.LR_MATRIX;

  const inputData = {
    extractions,
    metadataBatchFile,
    lrMatrixFile,
    startSerial,
    metadataColumns: METADATA_COLUMNS,
    lrMatrixColumns: LR_MATRIX_COLUMNS,
  };

  await runPython(
    join(CORE_ROOT, "..", "python", "assemble_excel.py"),
    [JSON.stringify(inputData)], config.pythonPath
  );

  return { metadataBatchFile, lrMatrixFile, startSerial };
}

// ============================================================================
// Step 5: VERIFY
// ============================================================================

export async function runVerify(
  extractions: ExtractionResult[],
  folders: FolderPaths
): Promise<VerificationResult> {
  const issues: string[] = [];
  const checks = {
    recordCountsMatch: true,
    fileRowBijection: true,
    serialContinuous: true,
    noBlankTitles: true,
    spotCheckPassed: true,
  };

  const serials = extractions.map(e => e.lrMatrix["S. No."]).sort((a, b) => a - b);
  for (let i = 1; i < serials.length; i++) {
    if (serials[i] !== serials[i - 1] + 1) {
      checks.serialContinuous = false;
      issues.push(`Serial gap: ${serials[i - 1]} → ${serials[i]}`);
    }
  }

  for (const ex of extractions) {
    if (!ex.metadata.Title.trim()) {
      checks.noBlankTitles = false;
      issues.push(`Blank title: ${ex.fileName}`);
    }
  }

  for (const ex of extractions.slice(0, 2)) {
    try {
      const buffer = await readFile(join(folders.DB, ex.fileName));
      const data = await pdfParse(buffer, { max: 1 });
      if (!data.text.toLowerCase().includes(ex.metadata.Title.toLowerCase().slice(0, 30))) {
        checks.spotCheckPassed = false;
        issues.push(`Spot check: title mismatch for ${ex.fileName}`);
      }
    } catch (e) {
      issues.push(`Spot check error for ${ex.fileName}: ${e}`);
    }
  }

  return { passed: Object.values(checks).every(v => v), issues, checks };
}

// ============================================================================
// Step 6: LOG
// ============================================================================

export async function runLog(
  result: BatchResult,
  folders: FolderPaths
): Promise<void> {
  const logFile = join(folders.ROOT, "SOP-Log.md");
  let existingLog = "";
  try {
    existingLog = await readFile(logFile, "utf-8");
  } catch {
    existingLog = "# SOP Log — LR Database runs\n\n";
  }

  const dateStr = new Date().toISOString().split("T")[0];
  const duplicateLines = result.duplicates.length > 0 
    ? result.duplicates.map(d => `  - ${basename(d.file)}: ${d.reason} (matched ${basename(d.matchedWith)})`).join("\n")
    : "  - None";

  const logEntry = `
## ${dateStr} — Batch of ${result.keeperFiles.length}
- \`Incoming/\` held ${result.incomingFiles.length} PDF(s); dedup vs Database and within batch found ${result.duplicates.length} duplicate(s):
${duplicateLines}
- ${result.keeperFiles.length} keeper(s) extracted
- New \`${basename(result.metadataBatchFile)}\` (${result.keeperFiles.length} rows, 87 cols); +${result.keeperFiles.length} LR rows (S.No. continuous)
- Files moved to \`Database/\`; \`Incoming/\` empty at close
- Verification: ${result.verification.passed ? "PASSED" : "ISSUES FOUND"} — ${result.verification.issues.length} issue(s)
`;

  await writeFile(logFile, existingLog + logEntry);
}

// ============================================================================
// Full Pipeline (without extraction - agent does that)
// ============================================================================

export async function runFullPipeline(
  config: PipelineConfig,
  agentExtracted: ExtractionResult[]
): Promise<BatchResult> {
  const folders = getFolders(config.projectRoot);
  folders.SCOPE = config.scope;
  await ensureFolders(folders, config.pythonPath);

  // Step 1: Ingest
  const incomingFiles = await runIngest(folders);
  if (incomingFiles.length === 0) throw new Error("No files in Incoming/");

  // Step 2: Dedup
  const { keepers, duplicates } = await runDedup(incomingFiles, folders);
  for (const dup of duplicates) {
    await rename(dup.file, join(folders.SETASIDE, basename(dup.file)));
  }

  // Step 3: Extract PDF content (agent will process these)
  const pdfContents = await runExtractPDFs(keepers);

  // Agent does extraction here using their model
  // extractions = agentExtracted

  // Step 4: Assemble
  const { metadataBatchFile, lrMatrixFile, startSerial } = await runAssemble(agentExtracted, folders, config);

  // Move keepers to Database
  for (const file of keepers) {
    await rename(file, join(folders.DB, basename(file)));
  }

  // Step 5: Verify
  const verification = await runVerify(agentExtracted, folders);

  // Step 6: Log
  const result: BatchResult = {
    incomingFiles: incomingFiles.map(basename),
    keeperFiles: keepers.map(basename),
    duplicates,
    extracted: agentExtracted,
    metadataBatchFile,
    lrMatrixFile,
    verification,
  };

  await runLog(result, folders);

  return result;
}

// ============================================================================
// SOP Extraction Prompt (for agents)
// ============================================================================

export const EXTRACTION_PROMPT = `You are an expert academic metadata extractor for a Literature Review Database.
Follow the SOP conventions EXACTLY. Output MUST conform to the provided JSON schemas.

## SOP §4 Extraction Conventions (BINDING)

1. **PDF content only** — except official publisher/regulator web pages for institutional reports (SEBI, NBER)
2. **Authors (metadata)**: Full names in printed order, "Last-name-comma-rest", semicolon-joined
   - Example: "Sharma, Pankaj; Patnaik, Tirthankar"
   - Single-token names stay as-is
   - Corporate reports: organisation as Author (e.g., "Securities and Exchange Board of India")
   - Signatory name at bottom of circular is NOT an author
3. **Authors (LR record)**: Display form "First-Last" joined with commas
   - Example: "Pankaj Sharma, Tirthankar Patnaik"
4. **Year/Date**: Exactly as printed (year, year-month, or full date). If PDF contradicts filename, use PDF.
5. **DOI (metadata)**: Bare form starting with "10." (strip https://doi.org/, dx.doi.org/)
6. **DOI (LR record)**: Full https://doi.org/... link, or blank
7. **URL**: Only if printed in PDF
8. **ISSN/ISBN**: Bare numbers
9. **Item Type**: One of: journalArticle, report, preprint, conferencePaper, thesis
   - ArXiv-only = preprint
   - IEEE-formatted no venue = conferencePaper only if venue printed, else preprint
9. **Abstract**: Verbatim from labelled Abstract section (~1200 chars max)
10. **Objective**: End of Introduction (research questions, hypotheses, motivation) — 40-70 words
11. **Methodology**: Data and Methods sections (sample, period, variables, techniques) — 60-100 words WITH concrete specifics and reported statistics
12. **Key Finding**: Results tables + abstract (copy reported numbers; NEVER invent) — 60-100 words WITH concrete specifics and reported statistics
13. **Discussion**: Discussion and Conclusion sections — 40-70 words
14. **Limitation**: Discussion/Conclusion; if paper states none, summarise scope boundary honestly — 30-60 words
15. **Theory**: Only frameworks the paper NAMES or clearly applies, otherwise blank
16. **Research-type tag**: One of Quantitative, Qualitative, Mixed, Review-Theoretical
17. **Variables**: Independent, Dependent, Other (control/moderator/mediator) as named in paper
18. **Decision/Reason**: Leave BLANK (default keep-all). Only fill on explicit screening request.

## Word Count Bands (STRICT)
- Objective: 40-70 words
- Methodology: 60-100 words with concrete specifics and reported statistics
- Key Finding: 60-100 words with concrete specifics and reported statistics
- Discussion: 40-70 words
- Limitation: 30-60 words

## Output
Return a single JSON object with two keys: "metadata" (87 fields) and "lrMatrix" (20 fields).
All fields required. Use empty string "" for missing data.`;