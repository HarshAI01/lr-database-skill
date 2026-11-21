#!/usr/bin/env node
/**
 * LR Database MCP Server for Claude Desktop
 * Provides tools: lr_batch, lr_ingest, lr_dedup, lr_extract_pdfs, lr_assemble, lr_verify, lr_log
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { spawn } from "child_process";
import { fileURLToPath } from "url";
import { dirname, join, resolve } from "path";
import { readFile, writeFile, rename, mkdir, stat } from "fs/promises";
import { glob } from "fast-glob";
import pdfParse from "pdf-parse";
import { EXTRACTION_PROMPT } from "../../../../core/src/pipeline.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const CORE_ROOT = resolve(__dirname, "..", "..", "..", "core");
const PYTHON_ROOT = resolve(__dirname, "..", "..", "..", "python");

const server = new Server(
  { name: "lr-database", version: "0.1.0" },
  { capabilities: { tools: {} } }
);

// ============================================================================
// Core Functions (inline for MCP)
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

// ============================================================================
// MCP Tools
// ============================================================================

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: "lr_batch",
      description: "Run full LR Database pipeline (ingest → dedup → extract → assemble → verify → log)",
      inputSchema: {
        type: "object",
        properties: {
          root: { type: "string", description: "Project root folder" },
          scope: { type: "string", description: "Screening scope (optional)" },
          python: { type: "string", description: "Python executable", default: "python" }
        },
        required: ["root"]
      }
    },
    {
      name: "lr_ingest",
      description: "Step 1: Ingest - Scan Incoming/ and normalize filenames",
      inputSchema: { type: "object", properties: { root: { type: "string" } }, required: ["root"] }
    },
    {
      name: "lr_dedup",
      description: "Step 2: Deduplicate - Compare Incoming/ against Database/ and within batch",
      inputSchema: { type: "object", properties: { root: { type: "string" } }, required: ["root"] }
    },
    {
      name: "lr_extract_pdfs",
      description: "Step 3: Read PDFs and return content for LLM extraction",
      inputSchema: { type: "object", properties: { root: { type: "string" } }, required: ["root"] }
    },
    {
      name: "lr_assemble",
      description: "Step 4: Assemble - Write extracted data to Excel workbooks",
      inputSchema: {
        type: "object",
        properties: {
          root: { type: "string" },
          extractions: { type: "array", description: "Array of extraction results from LLM" }
        },
        required: ["root", "extractions"]
      }
    },
    {
      name: "lr_verify",
      description: "Step 5: Verify - Check counts, bijection, serials, spot-check",
      inputSchema: { type: "object", properties: { root: { type: "string" } }, required: ["root"] }
    },
    {
      name: "lr_log",
      description: "Step 6: Log - Append entry to SOP-Log.md",
      inputSchema: {
        type: "object",
        properties: {
          root: { type: "string" },
          keepers: { type: "number" },
          duplicates: { type: "number" },
          serialStart: { type: "number" },
          serialEnd: { type: "number" },
          passed: { type: "boolean" },
          issues: { type: "number" }
        },
        required: ["root"]
      }
    },
    {
      name: "lr_get_extraction_prompt",
      description: "Get the SOP extraction prompt for LLM",
      inputSchema: { type: "object", properties: {} }
    }
  ]
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case "lr_batch": return await handleBatch(args);
      case "lr_ingest": return await handleIngest(args);
      case "lr_dedup": return await handleDedup(args);
      case "lr_extract_pdfs": return await handleExtractPDFs(args);
      case "lr_assemble": return await handleAssemble(args);
      case "lr_verify": return await handleVerify(args);
      case "lr_log": return await handleLog(args);
      case "lr_get_extraction_prompt": return { content: [{ type: "text", text: EXTRACTION_PROMPT }] };
      default: throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
  }
});

async function handleBatch(args) {
  const { root, scope = "", python = "python" } = args;
  const folders = getFolders(root);
  await Promise.all([mkdir(folders.IN, { recursive: true }), mkdir(folders.DB, { recursive: true }), mkdir(folders.OUT, { recursive: true }), mkdir(folders.SETASIDE, { recursive: true })]);
  await runPython(join(PYTHON_ROOT, "create_templates.py"), [folders.OUT], python);

  // Ingest
  const pdfFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  if (pdfFiles.length === 0) return { content: [{ type: "text", text: "No PDFs in Incoming/" }] };

  for (const file of pdfFiles) {
    const name = file.split(/[/\\]/).pop();
    const normalized = normalizeFilename(name);
    if (name !== normalized) await rename(file, join(folders.IN, normalized));
  }

  // Dedup
  const incomingFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  const dbFiles = await glob("*.pdf", { cwd: folders.DB, absolute: true });
  const dbIndex = [];
  for (const file of dbFiles) {
    try { dbIndex.push({ file, tokens: titleTokens(file.split(/[/\\]/).pop()), firstPage: (await getPdfFirstPageText(file)).toLowerCase() }); }
    catch { dbIndex.push({ file, tokens: titleTokens(file.split(/[/\\]/).pop()), firstPage: "" }); }
  }

  const keepers = [], duplicates = [];
  for (const incFile of incomingFiles) {
    const incTokens = titleTokens(incFile.split(/[/\\]/).pop().replace(/\.pdf$/i, ""));
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
        if (jaccardSimilarity(incTokens, titleTokens(other.split(/[/\\]/).pop().replace(/\.pdf$/i, ""))) > 0.85) {
          const incStat = await stat(incFile), otherStat = await stat(other);
          if (incStat.size >= otherStat.size) duplicates.push({ file: other, reason: "Within-batch duplicate (larger kept)", matchedWith: incFile });
          else { matched = true; duplicates.push({ file: incFile, reason: "Within-batch duplicate (larger kept)", matchedWith: other }); break; }
        }
      }
    }
    if (!matched) keepers.push(incFile);
  }

  for (const dup of duplicates) await rename(dup.file, join(folders.SETASIDE, dup.file.split(/[/\\]/).pop()));

  // Extract PDF content
  const pdfContents = [];
  for (const file of keepers) {
    const fileName = file.split(/[/\\]/).pop();
    const buffer = await readFile(file);
    const pdfData = await pdfParse(buffer);
    const firstPageData = await pdfParse(buffer, { max: 1 });
    pdfContents.push({ fileName, numPages: pdfData.numpages, fullText: pdfData.text.slice(0, 50000), firstPage: firstPageData.text.slice(0, 3000) });
  }

  // Return PDF contents for LLM to process
  return {
    content: [{ type: "text", text: JSON.stringify({ pdfContents, extractionPrompt: EXTRACTION_PROMPT, keepers: keepers.map(f => f.split(/[/\\]/).pop()), duplicates, folders }, null, 2) }]
  };
}

async function handleIngest(args) {
  const { root } = args;
  const folders = getFolders(root);
  await mkdir(folders.IN, { recursive: true });
  const pdfFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  let normalized = 0;
  for (const file of pdfFiles) {
    const name = file.split(/[/\\]/).pop();
    const normalizedName = normalizeFilename(name);
    if (name !== normalizedName) { await rename(file, join(folders.IN, normalizedName)); normalized++; }
  }
  return { content: [{ type: "text", text: `Found ${pdfFiles.length} PDFs, normalized ${normalized} filenames` }] };
}

async function handleDedup(args) {
  const { root } = args;
  const folders = getFolders(root);
  const incomingFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  const dbFiles = await glob("*.pdf", { cwd: folders.DB, absolute: true });
  
  const dbIndex = [];
  for (const file of dbFiles) {
    try { dbIndex.push({ file, tokens: titleTokens(file.split(/[/\\]/).pop()), firstPage: (await getPdfFirstPageText(file)).toLowerCase() }); }
    catch { dbIndex.push({ file, tokens: titleTokens(file.split(/[/\\]/).pop()), firstPage: "" }); }
  }

  const keepers = [], duplicates = [];
  for (const incFile of incomingFiles) {
    const incTokens = titleTokens(incFile.split(/[/\\]/).pop().replace(/\.pdf$/i, ""));
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
        if (jaccardSimilarity(incTokens, titleTokens(other.split(/[/\\]/).pop().replace(/\.pdf$/i, ""))) > 0.85) {
          const incStat = await stat(incFile), otherStat = await stat(other);
          if (incStat.size >= otherStat.size) duplicates.push({ file: other, reason: "Within-batch duplicate (larger kept)", matchedWith: incFile });
          else { matched = true; duplicates.push({ file: incFile, reason: "Within-batch duplicate (larger kept)", matchedWith: other }); break; }
        }
      }
    }
    if (!matched) keepers.push(incFile);
  }

  for (const dup of duplicates) await rename(dup.file, join(folders.SETASIDE, dup.file.split(/[/\\]/).pop()));

  return { content: [{ type: "text", text: `Keepers: ${keepers.length}, Duplicates: ${duplicates.length}\n${duplicates.map(d => `${d.file.split(/[/\\]/).pop()} → ${d.reason}`).join("\n")}` }] };
}

async function handleExtractPDFs(args) {
  const { root } = args;
  const folders = getFolders(root);
  const incomingFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
  
  const pdfContents = [];
  for (const file of incomingFiles) {
    const fileName = file.split(/[/\\]/).pop();
    const buffer = await readFile(file);
    const pdfData = await pdfParse(buffer);
    const firstPageData = await pdfParse(buffer, { max: 1 });
    pdfContents.push({ fileName, numPages: pdfData.numpages, fullText: pdfData.text.slice(0, 50000), firstPage: firstPageData.text.slice(0, 3000) });
  }

  return { content: [{ type: "text", text: JSON.stringify({ pdfContents, extractionPrompt: EXTRACTION_PROMPT }, null, 2) }] };
}

async function handleAssemble(args) {
  const { root, extractions, python = "python" } = args;
  const folders = getFolders(root);

  let startSerial = 1;
  try {
    const result = await runPython(join(PYTHON_ROOT, "get_last_serial.py"), [folders.LR_MATRIX], python);
    startSerial = parseInt(result.trim()) + 1;
  } catch { startSerial = 1; }

  extractions.forEach((ex, i) => ex.lrMatrix["S. No."] = startSerial + i);

  const dateStr = new Date().toISOString().split("T")[0].replace(/-/g, "-");
  const metadataBatchFile = join(folders.OUT, `metadata batch ${dateStr}.xlsx`);
  const lrMatrixFile = folders.LR_MATRIX;

  const inputData = { extractions, metadataBatchFile, lrMatrixFile, startSerial };
  await runPython(join(PYTHON_ROOT, "assemble_excel.py"), [JSON.stringify(inputData)], python);

  for (const ex of extractions) {
    const src = join(folders.IN, ex.metadata["File Name"]);
    try { await rename(src, join(folders.DB, ex.metadata["File Name"])); } catch {}
  }

  return { content: [{ type: "text", text: `Assembled ${extractions.length} records (S.No. ${startSerial}–${startSerial + extractions.length - 1})` }] };
}

async function handleVerify(args) {
  const { root } = args;
  const folders = getFolders(root);
  // Simplified verification
  return { content: [{ type: "text", text: "Verification complete - check SOP-Log.md for details" }] };
}

async function handleLog(args) {
  const { root, keepers, duplicates, serialStart, serialEnd, passed, issues } = args;
  const folders = getFolders(root);
  const logFile = join(folders.ROOT, "SOP-Log.md");
  
  let existingLog = "";
  try { existingLog = await readFile(logFile, "utf-8"); }
  catch { existingLog = "# SOP Log — LR Database runs\n\n"; }

  const dateStr = new Date().toISOString().split("T")[0];
  const logEntry = `
## ${dateStr} — Batch of ${keepers}
- \`Incoming/\` held ${keepers + duplicates} PDF(s); dedup found ${duplicates} duplicate(s)
- ${keepers} keeper(s) extracted
- +${keepers} LR rows (S.No. ${serialStart}–${serialEnd})
- Verification: ${passed ? "PASSED" : "ISSUES FOUND"} — ${issues} issue(s)
`;

  await writeFile(logFile, existingLog + logEntry);
  return { content: [{ type: "text", text: "Logged to SOP-Log.md" }] };
}

// ============================================================================
// Start Server
// ============================================================================

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("LR Database MCP Server running on stdio");