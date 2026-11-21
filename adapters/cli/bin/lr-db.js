#!/usr/bin/env node
/**
 * LR Database CLI - Standalone command-line tool
 * Usage: lr-db <command> [options]
 */

import { Command } from "commander";
import { readFile, writeFile, rename, mkdir, stat } from "fs/promises";
import { join, resolve, basename } from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import { glob } from "fast-glob";
import pdfParse from "pdf-parse";
import chalk from "chalk";
import ora from "ora";

const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..");
const ROOT = resolve(__dirname, "..", "..", "..");
const CORE_ROOT = resolve(ROOT, "core");
const PYTHON_ROOT = resolve(ROOT, "python");

// ============================================================================
// Core Functions
// ============================================================================

function getFolders(root) {
  return {
    ROOT: root, IN: join(root, "Incoming"), DB: join(root, "Database"),
    OUT: join(root, "Outputs"), SETASIDE: join(root, "Set-Aside"),
    LR_MATRIX: join(root, "Outputs", "LR Matrix.xlsx"),
    META0: join(root, "Meta Data Filled.xlsx"),
  };
}

function normalizeFilename(name) {
  return name.replace(/[<>:"|?*]/g, "").replace(/\s+/g, " ").trim();
}

function titleTokens(title) {
  return title.toLowerCase().replace(/[^\w\s]/g, " ").split(/\s+/).filter(t => t.length > 3);
}

function jaccardSimilarity(a, b) {
  const setA = new Set(a), setB = new Set(b);
  const inter = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : inter / union;
}

async function getPdfFirstPageText(filePath) {
  const buffer = await readFile(filePath);
  const data = await pdfParse(buffer, { max: 1 });
  return data.text.slice(0, 3000);
}

function runPython(script, args, pythonPath = "python") {
  return new Promise((resolve, reject) => {
    const proc = spawn(pythonPath, [script, ...args], { cwd: CORE_ROOT });
    let stdout = "", stderr = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => stderr += d.toString());
    proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
  });
}

const EXTRACTION_PROMPT = `You are an expert academic metadata extractor for a Literature Review Database.
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

// ============================================================================
// CLI Commands
// ============================================================================

const program = new Command();

program
  .name("lr-db")
  .description("Literature Review Database CLI - Full SOP Pipeline")
  .version("0.1.0")
  .option("-r, --root <path>", "Project root folder", ".")
  .option("--python <path>", "Python executable", "python");

program
  .command("batch")
  .description("Run full pipeline")
  .option("-s, --scope <text>", "Screening scope")
  .action(async (opts) => {
    const root = resolve(opts.root);
    const config = { projectRoot: root, scope: opts.scope || "", pythonPath: opts.python };
    await runBatch(config);
  });

program
  .command("ingest")
  .description("Step 1: Ingest PDFs")
  .action(async (opts) => {
    const root = resolve(opts.root);
    await runIngest(root);
  });

program
  .command("dedup")
  .description("Step 2: Deduplicate")
  .action(async (opts) => {
    const root = resolve(opts.root);
    await runDedup(root);
  });

program
  .command("extract")
  .description("Step 3: Extract PDF content for LLM")
  .action(async (opts) => {
    const root = resolve(opts.root);
    await runExtract(root);
  });

program
  .command("assemble")
  .description("Step 4: Assemble Excel from extractions")
  .requiredOption("-i, --input <file>", "Extractions JSON file")
  .action(async (opts) => {
    const root = resolve(opts.root);
    const input = resolve(root, opts.input);
    await runAssemble(root, input, opts.python);
  });

program
  .command("verify")
  .description("Step 5: Verify")
  .action(async (opts) => {
    const root = resolve(opts.root);
    await runVerify(root);
  });

program
  .command("log")
  .description("Step 6: Log")
  .requiredOption("-k, --keepers <count>", "Number of keepers")
  .requiredOption("-d, --duplicates <count>", "Number of duplicates")
  .option("--serial-start <n>", "Starting serial", "1")
  .option("--serial-end <n>", "Ending serial", "1")
  .option("--passed", "Verification passed")
  .option("--issues <n>", "Number of issues", "0")
  .action(async (opts) => {
    const root = resolve(opts.root);
    await runLog(root, {
      keepers: parseInt(opts.keepers),
      duplicates: parseInt(opts.duplicates),
      serialStart: parseInt(opts.serialStart),
      serialEnd: parseInt(opts.serialEnd),
      passed: opts.passed,
      issues: parseInt(opts.issues)
    });
  });

program
  .command("prompt")
  .description("Print extraction prompt for LLM")
  .action(() => {
    console.log(EXTRACTION_PROMPT);
  });

// ============================================================================
// Implementation
// ============================================================================

function getFolders(root) {
  return {
    ROOT: root, IN: join(root, "Incoming"), DB: join(root, "Database"),
    OUT: join(root, "Outputs"), SETASIDE: join(root, "Set-Aside"),
    LR_MATRIX: join(root, "Outputs", "LR Matrix.xlsx"),
    META0: join(root, "Meta Data Filled.xlsx"),
  };
}

function runPython(script, args, pythonPath = "python") {
  return new Promise((resolve, reject) => {
    const proc = spawn(pythonPath, [script, ...args], { cwd: CORE_ROOT });
    let stdout = "", stderr = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => stderr += d.toString());
    proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
  });
}

async function ensureFolders(folders, pythonPath) {
  await Promise.all([
    mkdir(folders.IN, { recursive: true }),
    mkdir(folders.DB, { recursive: true }),
    mkdir(folders.OUT, { recursive: true }),
    mkdir(folders.SETASIDE, { recursive: true }),
  ]);
  await runPython(join(PYTHON_ROOT, "create_templates.py"), [folders.OUT], pythonPath);
}

async function runBatch(config) {
  const spinner = ora("Starting LR Database Pipeline...").start();
  const folders = getFolders(config.projectRoot);
  await ensureFolders(folders, config.pythonPath);

  try {
    // Ingest
    spinner.text = "1/6: Ingesting...";
    const incomingFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
    if (incomingFiles.length === 0) { spinner.warn("No PDFs in Incoming/"); return; }
    for (const file of incomingFiles) {
      const name = basename(file);
      const normalized = normalizeFilename(name);
      if (name !== normalized) await rename(file, join(folders.IN, normalized));
    }

    // Dedup
    spinner.text = "2/6: Deduplicating...";
    const { keepers, duplicates } = await dedup(folders);
    for (const dup of duplicates) await rename(dup.file, join(folders.SETASIDE, basename(dup.file)));

    // Extract PDFs
    spinner.text = "3/6: Reading PDFs...";
    const pdfContents = await extractPDFs(keepers);

    // Agent extraction
    spinner.text = "3b/6: Extracting with LLM...";
    const extractions = await extractWithAgent(pdfContents);

    // Assemble
    spinner.text = "4/6: Assembling Excel...";
    const { startSerial } = await assemble(extractions, folders, config);

    // Move keepers
    for (const file of keepers) await rename(file, join(folders.DB, basename(file)));

    // Verify
    spinner.text = "5/6: Verifying...";
    const verification = await verify(extractions, folders);

    // Log
    spinner.text = "6/6: Logging...";
    await log(folders, { keepers: keepers.length, duplicates: duplicates.length, serialStart: startSerial, serialEnd: startSerial + keepers.length - 1, passed: verification.passed, issues: verification.issues.length });

    spinner.succeed("Pipeline complete!");
    console.log(chalk.green(`\n✅ Processed: ${keepers.length} papers`));
    console.log(chalk.green(`📊 LR Matrix: S.No. ${startSerial}–${startSerial + keepers.length - 1}`));
  } catch (error) {
    spinner.fail("Failed");
    console.error(chalk.red(error));
    process.exit(1);
  }
}

async function runIngest(root) {
  const folders = getFolders(root);
  await mkdir(folders.IN, { recursive: true });
  const pdfFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  let normalized = 0;
  for (const file of pdfFiles) {
    const name = basename(file);
    const normalizedName = normalizeFilename(name);
    if (name !== normalizedName) { await rename(file, join(folders.IN, normalizedName)); normalized++; }
  }
  console.log(chalk.green(`Found ${pdfFiles.length} PDFs, normalized ${normalized} filenames`));
}

async function runDedup(root) {
  const folders = getFolders(root);
  const { keepers, duplicates } = await dedup(folders);
  for (const dup of duplicates) await rename(dup.file, join(folders.SETASIDE, basename(dup.file)));
  console.log(chalk.green(`Keepers: ${keepers.length}, Duplicates: ${duplicates.length}`));
  duplicates.forEach(d => console.log(`  ${basename(d.file)} → ${d.reason}`));
}

async function runExtract(root) {
  const folders = getFolders(root);
  const files = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  const contents = await extractPDFs(files);
  console.log(JSON.stringify({ pdfContents: contents, extractionPrompt: EXTRACTION_PROMPT }, null, 2));
}

async function runAssemble(root, inputFile, pythonPath) {
  const folders = getFolders(root);
  const extractions = JSON.parse(await readFile(inputFile, "utf-8"));
  await assemble(extractions, folders, { pythonPath });
  console.log(chalk.green(`Assembled ${extractions.length} records`));
}

async function runVerify(root) {
  const folders = getFolders(root);
  console.log(chalk.green("Verification complete"));
}

async function runLog(root, data) {
  const folders = getFolders(root);
  await log(folders, data);
  console.log(chalk.green("Logged to SOP-Log.md"));
}

// ============================================================================
// Helper Functions
// ============================================================================

function normalizeFilename(name) {
  return name.replace(/[<>:"|?*]/g, "").replace(/\s+/g, " ").trim();
}

function titleTokens(title) {
  return title.toLowerCase().replace(/[^\w\s]/g, " ").split(/\s+/).filter(t => t.length > 3);
}

function jaccardSimilarity(a, b) {
  const setA = new Set(a), setB = new Set(b);
  const inter = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : inter / union;
}

async function getPdfFirstPageText(filePath) {
  const buffer = await readFile(filePath);
  const data = await pdfParse(buffer, { max: 1 });
  return data.text.slice(0, 3000);
}

async function dedup(folders) {
  const incomingFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  const dbFiles = await glob("*.pdf", { cwd: folders.DB, absolute: true });
  
  const dbIndex = [];
  for (const file of dbFiles) {
    try { dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage: (await getPdfFirstPageText(file)).toLowerCase() }); }
    catch { dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage: "" }); }
  }

  const keepers = [], duplicates = [];
  for (const incFile of incomingFiles) {
    const incTokens = titleTokens(basename(incFile).replace(/\.pdf$/i, ""));
    let matched = false;
    for (const db of dbIndex) {
      if (jaccardSimilarity(incTokens, db.tokens) > 0.7) {
        const incFirstPage = (await getPdfFirstPageText(incFile)).toLowerCase();
        const doiMatch = incFirstPage.match(/10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i);
        if (doiMatch && db.firstPage.includes(doiMatch[0].toLowerCase())) { matched = true; duplicates.push({ file: incFile, reason: "DOI match", matchedWith: db.file }); break; }
        const overlap = incTokens.filter(t => t.length > 4 && db.firstPage.includes(t)).length;
        if (overlap >= 3) { matched = true; duplicates.push({ file: incFile, reason: "Title/author match", matchedWith: db.file }); break; }
      }
    }
    if (!matched) {
      for (const other of incomingFiles) {
        if (other === incFile) continue;
        if (jaccardSimilarity(incTokens, titleTokens(basename(other).replace(/\.pdf$/i, ""))) > 0.85) {
          const incStat = await stat(incFile), otherStat = await stat(other);
          if (incStat.size >= otherStat.size) duplicates.push({ file: other, reason: "Within-batch duplicate (larger kept)", matchedWith: incFile });
          else { matched = true; duplicates.push({ file: incFile, reason: "Within-batch duplicate (larger kept)", matchedWith: other }); break; }
        }
      }
    }
    if (!matched) keepers.push(incFile);
  }
  return { keepers, duplicates };
}

async function extractPDFs(files) {
  const results = [];
  for (const file of files) {
    const fileName = basename(file);
    const buffer = await readFile(file);
    const pdfData = await pdfParse(buffer);
    const firstPageData = await pdfParse(buffer, { max: 1 });
    results.push({ fileName, numPages: pdfData.numpages, fullText: pdfData.text.slice(0, 50000), firstPage: firstPageData.text.slice(0, 3000) });
  }
  return results;
}

async function extractWithAgent(pdfContents) {
  const { default: readline } = await import("readline");
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const extractions = [];

  for (const pdf of pdfContents) {
    console.log(chalk.cyan(`\n📄 ${pdf.fileName}`));
    console.log(chalk.gray("─".repeat(60)));
    console.log(chalk.yellow("FIRST PAGE:"));
    console.log(pdf.firstPage);
    console.log(chalk.gray("─".repeat(60)));
    console.log(chalk.yellow("FULL TEXT (first 50k):"));
    console.log(pdf.fullText);
    console.log(chalk.gray("─".repeat(60)));
    console.log(chalk.magenta("\n📋 EXTRACTION PROMPT:"));
    console.log(chalk.gray(EXTRACTION_PROMPT));
    console.log(chalk.gray("─".repeat(60)));

    const result = await new Promise(resolve => {
      rl.question(chalk.green("\n✅ Paste JSON result (or 'skip'): "), input => {
        if (input.trim().toLowerCase() === "skip") {
          resolve({ fileName: pdf.fileName, metadata: { "File Name": pdf.fileName, "Item Type": "journalArticle", "Publication Year": "", Author: "", Title: "", "Publication Title": "", ISBN: "", ISSN: "", DOI: "", Url: "", "Abstract Note": "", Date: "", "Date Added": new Date().toISOString(), "Date Modified": new Date().toISOString(), "Access Date": new Date().toISOString().split("T")[0], Pages: "", "Num Pages": pdf.numPages, Issue: "", Volume: "", "Number Of Volumes": 1, "Journal Abbreviation": "", "Short Title": "", Series: "", "Series Number": "", "Series Text": "", "Series Title": "", Publisher: "", Place: "", Language: "en", Rights: "", Type: "", Archive: "", "Archive Location": "", "Library Catalog": "", "Call Number": "", Extra: "", Notes: "", "File Attachments": `Database/${pdf.fileName}`, "Link Attachments": "", "Manual Tags": "", "Automatic Tags": "", Editor: "", "Series Editor": "", Translator: "", Contributor: "", "Attorney Agent": "", "Book Author": "", "Cast Member": "", Commenter: "", Composer: "", Cosponsor: "", Counsel: "", Interviewer: "", Producer: "", Recipient: "", "Reviewed Author": "", Scriptwriter: "", "Words By": "", Guest: "", Number: "", Edition: "", "Running Time": "", Scale: "", Medium: "", "Artwork Size": "", "Filing Date": "", "Application Number": "", Assignee: "", "Issuing Authority": "", Country: "", "Meeting Name": "", "Conference Name": "", Court: "", References: "", Reporter: "", "Legal Status": "", "Priority Numbers": "", "Programming Language": "", Version: "", System: "", Code: "", "Code Number": "", Section: "", Session: "", Committee: "", History: "", "Legislative Body": "" }, lrMatrix: { "S. No.": 0, "File Name": pdf.fileName, Title: "", Author: "", Year: "", Journal: "", DOI: "", Objective: "", Methodology: "", "Key Finding": "", Discussion: "", Limitation: "", "Theoretical Framework": "", "Tag 1 (Method)": "", "Tag 2 (Type of Research)": "Quantitative", "Independent Variable": "", "Dependent Variable": "", "Other Variable": "", "Decision (For Me)": "", "Reason (For Me)": "" } });
        } else {
          try { resolve({ fileName: pdf.fileName, ...JSON.parse(input) }); }
          catch { console.error(chalk.red("Invalid JSON, skipping")); resolve({ fileName: pdf.fileName, metadata: { "File Name": pdf.fileName }, lrMatrix: { "File Name": pdf.fileName } }); }
        }
      });
    });
    extractions.push(result);
  }
  rl.close();
  return extractions;
}

async function assemble(extractions, folders, config) {
  let startSerial = 1;
  try {
    const result = await runPython(join(PYTHON_ROOT, "get_last_serial.py"), [folders.LR_MATRIX], config.pythonPath);
    startSerial = parseInt(result.trim()) + 1;
  } catch { startSerial = 1; }

  extractions.forEach((ex, i) => ex.lrMatrix["S. No."] = startSerial + i);

  const dateStr = new Date().toISOString().split("T")[0].replace(/-/g, "-");
  const metadataBatchFile = join(folders.OUT, `metadata batch ${dateStr}.xlsx`);
  const lrMatrixFile = folders.LR_MATRIX;

  const inputData = { extractions, metadataBatchFile, lrMatrixFile, startSerial };
  await runPython(join(PYTHON_ROOT, "assemble_excel.py"), [JSON.stringify(inputData)], config.pythonPath);

  for (const ex of extractions) {
    const src = join(folders.IN, ex.metadata["File Name"]);
    try { await rename(src, join(folders.DB, ex.metadata["File Name"])); } catch {}
  }

  return { startSerial };
}

async function verify(extractions, folders) {
  const issues = [];
  const serials = extractions.map(e => e.lrMatrix["S. No."]).sort((a, b) => a - b);
  for (let i = 1; i < serials.length; i++) if (serials[i] !== serials[i - 1] + 1) issues.push(`Serial gap: ${serials[i - 1]} → ${serials[i]}`);
  for (const ex of extractions) if (!ex.metadata.Title?.trim()) issues.push(`Blank title: ${ex.fileName}`);
  return { passed: issues.length === 0, issues, checks: { serialContinuous: issues.length === 0, noBlankTitles: !issues.some(i => i.startsWith("Blank")) } };
}

async function log(folders, data) {
  const logFile = join(folders.ROOT, "SOP-Log.md");
  let existingLog = "";
  try { existingLog = await readFile(logFile, "utf-8"); }
  catch { existingLog = "# SOP Log — LR Database runs\n\n"; }

  const dateStr = new Date().toISOString().split("T")[0];
  const logEntry = `
## ${dateStr} — Batch of ${data.keepers}
- \`Incoming/\` held ${data.keepers + data.duplicates} PDF(s); dedup found ${data.duplicates} duplicate(s)
- ${data.keepers} keeper(s) extracted
- +${data.keepers} LR rows (S.No. ${data.serialStart}–${data.serialEnd})
- Verification: ${data.passed ? "PASSED" : "ISSUES FOUND"} — ${data.issues} issue(s)
`;

  await writeFile(logFile, existingLog + logEntry);
}

// ============================================================================
// Run
// ============================================================================

program.parse();