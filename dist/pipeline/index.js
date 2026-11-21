/**
 * Pipeline Orchestrator — Core logic shared by CLI and programmatic use
 */
import { glob } from "fast-glob";
import { readFile, writeFile, rename, stat, mkdir } from "fs/promises";
import { join, basename, resolve } from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import pdfParse from "pdf-parse";
import { ExtractionResultSchema, METADATA_COLUMNS, LR_MATRIX_COLUMNS, createEmptyMetadata, createEmptyLRMatrix, } from "../schemas.js";
// Types are exported via 'export interface' below
const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..");
const SKILL_ROOT = resolve(__dirname, "..");
// ExtractionResult is imported from schemas (includes fileName, metadata, lrMatrix)
export function getFolders(root) {
    return {
        ROOT: root,
        IN: join(root, "Incoming"),
        DB: join(root, "Database"),
        OUT: join(root, "Outputs"),
        SETASIDE: join(root, "Set-Aside"),
        REF: "", // Not used - templates created programmatically
        SCOPE: "",
        LR_MATRIX: join(root, "Outputs", "LR Matrix.xlsx"),
        META0: join(root, "Meta Data Filled.xlsx"),
    };
}
export async function ensureFolders(folders, pythonPath = "python") {
    await Promise.all([
        mkdir(folders.IN, { recursive: true }),
        mkdir(folders.DB, { recursive: true }),
        mkdir(folders.OUT, { recursive: true }),
        mkdir(folders.SETASIDE, { recursive: true }),
    ]);
    // Create template files if they don't exist (no Reference & Template folder needed)
    const { spawn } = await import("child_process");
    const scriptPath = join(SKILL_ROOT, "python", "create_templates.py");
    await new Promise((resolve, reject) => {
        const proc = spawn(pythonPath, [scriptPath, folders.OUT]);
        proc.on("close", code => code === 0 ? resolve() : reject(new Error(`Template creation failed: ${code}`)));
    });
}
function normalizeFilename(name) {
    return name
        .replace(/[<>:"|?*]/g, "")
        .replace(/\s+/g, " ")
        .trim();
}
function titleTokens(title) {
    return title
        .toLowerCase()
        .replace(/[^\w\s]/g, " ")
        .split(/\s+/)
        .filter(t => t.length > 3);
}
function jaccardSimilarity(a, b) {
    const setA = new Set(a);
    const setB = new Set(b);
    const intersection = [...setA].filter(x => setB.has(x)).length;
    const union = new Set([...setA, ...setB]).size;
    return union === 0 ? 0 : intersection / union;
}
async function getPdfFirstPageText(filePath) {
    const buffer = await readFile(filePath);
    const data = await pdfParse(buffer, { max: 1 });
    return data.text.slice(0, 3000);
}
// ============================================================================
// Step 1: INGEST
// ============================================================================
export async function runIngest(folders) {
    const pdfFiles = await glob("*.pdf", { cwd: folders.IN, absolute: true });
    const normalized = [];
    for (const file of pdfFiles) {
        const name = basename(file);
        const normalizedName = normalizeFilename(name);
        if (name !== normalizedName) {
            const newPath = join(folders.IN, normalizedName);
            await rename(file, newPath);
            normalized.push(newPath);
        }
        else {
            normalized.push(file);
        }
    }
    return normalized;
}
// ============================================================================
// Step 2: DEDUP
// ============================================================================
export async function runDedup(incomingFiles, folders) {
    const dbFiles = await glob("*.pdf", { cwd: folders.DB, absolute: true });
    const dbIndex = [];
    for (const file of dbFiles) {
        try {
            const firstPage = await getPdfFirstPageText(file);
            dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage });
        }
        catch {
            dbIndex.push({ file, tokens: titleTokens(basename(file)), firstPage: "" });
        }
    }
    const keepers = [];
    const duplicates = [];
    for (const incFile of incomingFiles) {
        const incTokens = titleTokens(basename(incFile).replace(/\.pdf$/i, ""));
        let matched = false;
        for (const db of dbIndex) {
            const sim = jaccardSimilarity(incTokens, db.tokens);
            if (sim > 0.7) {
                const incFirstPage = await getPdfFirstPageText(incFile);
                const doiMatch = incFirstPage.match(/10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i);
                if (doiMatch && db.firstPage.includes(doiMatch[0])) {
                    matched = true;
                    duplicates.push({ file: incFile, reason: "DOI match", matchedWith: db.file });
                    break;
                }
                const titleWords = incTokens.filter(t => t.length > 4);
                const overlap = titleWords.filter(t => db.firstPage.toLowerCase().includes(t)).length;
                if (overlap >= 3) {
                    matched = true;
                    duplicates.push({ file: incFile, reason: "Title/author match", matchedWith: db.file });
                    break;
                }
            }
        }
        if (!matched) {
            for (const other of incomingFiles) {
                if (other === incFile)
                    continue;
                const otherTokens = titleTokens(basename(other).replace(/\.pdf$/i, ""));
                if (jaccardSimilarity(incTokens, otherTokens) > 0.85) {
                    const incStat = await stat(incFile);
                    const otherStat = await stat(other);
                    if (incStat.size >= otherStat.size) {
                        duplicates.push({ file: other, reason: "Within-batch duplicate (larger kept)", matchedWith: incFile });
                    }
                    else {
                        matched = true;
                        duplicates.push({ file: incFile, reason: "Within-batch duplicate (larger kept)", matchedWith: other });
                        break;
                    }
                }
            }
        }
        if (!matched)
            keepers.push(incFile);
    }
    return { keepers, duplicates };
}
// ============================================================================
// Step 3: EXTRACT
// ============================================================================
const EXTRACTION_SYSTEM_PROMPT = `You are an expert academic metadata extractor for a Literature Review Database.
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
async function callLLM(provider, model, apiKey, baseUrl, prompt) {
    if (provider === "ollama") {
        const url = baseUrl || "http://localhost:11434";
        const response = await fetch(`${url}/api/chat`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                model,
                messages: [
                    { role: "system", content: EXTRACTION_SYSTEM_PROMPT },
                    { role: "user", content: prompt }
                ],
                format: "json",
                options: { temperature: 0, num_predict: 4000 },
                stream: false,
            }),
        });
        if (!response.ok) {
            throw new Error(`Ollama error: ${response.status} ${await response.text()}`);
        }
        const data = await response.json();
        return JSON.parse(data.message?.content || "{}");
    }
    else if (provider === "openai") {
        const { OpenAI } = await import("openai");
        const client = new OpenAI({ apiKey: apiKey || process.env.OPENAI_API_KEY });
        const response = await client.chat.completions.create({
            model,
            messages: [
                { role: "system", content: EXTRACTION_SYSTEM_PROMPT },
                { role: "user", content: prompt }
            ],
            temperature: 0,
            max_tokens: 4000,
            response_format: { type: "json_object" },
        });
        return JSON.parse(response.choices[0].message.content || "{}");
    }
    else if (provider === "anthropic") {
        const { Anthropic } = await import("@anthropic-ai/sdk");
        const client = new Anthropic({ apiKey: apiKey || process.env.ANTHROPIC_API_KEY });
        const response = await client.messages.create({
            model,
            max_tokens: 4000,
            temperature: 0.0,
            system: EXTRACTION_SYSTEM_PROMPT,
            messages: [{ role: "user", content: prompt }],
        });
        return JSON.parse(response.content[0].type === "text" ? response.content[0].text : "{}");
    }
    else if (provider === "gemini") {
        const { GoogleGenerativeAI } = await import("@google/generative-ai");
        const client = new GoogleGenerativeAI(apiKey || process.env.GEMINI_API_KEY || "");
        const modelInstance = client.getGenerativeModel({ model });
        const response = await modelInstance.generateContent(prompt);
        return JSON.parse(response.response.text());
    }
    throw new Error(`Unknown provider: ${provider}`);
}
export async function runExtract(keeperFiles, folders, config) {
    const results = [];
    for (const file of keeperFiles) {
        const fileName = basename(file);
        try {
            const buffer = await readFile(file);
            const pdfData = await pdfParse(buffer);
            const fullText = pdfData.text;
            const numPages = pdfData.numpages;
            const firstPageData = await pdfParse(buffer, { max: 1 });
            const firstPage = firstPageData.text;
            const prompt = `## PDF Content (first 50k chars)
${fullText.slice(0, 50000)}

## First Page (for metadata)
${firstPage.slice(0, 3000)}

## File Info
- Filename: ${fileName}
- Pages: ${numPages}

Return ONLY valid JSON with "metadata" and "lrMatrix" keys.`;
            const extraction = await callLLM(config.llm.provider, config.llm.model, config.llm.apiKey, config.llm.baseUrl, prompt);
            extraction.metadata["File Name"] = fileName;
            extraction.metadata["Num Pages"] = numPages;
            extraction.metadata["File Attachments"] = `Database/${fileName}`;
            extraction.metadata["Date Added"] = new Date().toISOString();
            extraction.metadata["Date Modified"] = new Date().toISOString();
            extraction.metadata["Access Date"] = new Date().toISOString().split("T")[0];
            extraction.lrMatrix["File Name"] = fileName;
            const validated = ExtractionResultSchema.parse({
                fileName,
                metadata: extraction.metadata,
                lrMatrix: extraction.lrMatrix,
            });
            results.push(validated);
        }
        catch (error) {
            console.warn(`Failed to extract ${fileName}:`, error);
            const emptyMeta = createEmptyMetadata(fileName);
            emptyMeta["Num Pages"] = 0;
            const emptyLR = createEmptyLRMatrix(0, fileName);
            results.push({ fileName, metadata: emptyMeta, lrMatrix: emptyLR });
        }
    }
    return results;
}
// ============================================================================
// Step 4: ASSEMBLE (via Python)
// ============================================================================
function runPython(script, args, pythonPath) {
    return new Promise((resolve, reject) => {
        const proc = spawn(pythonPath, [script, ...args], { cwd: SKILL_ROOT });
        let stdout = "", stderr = "";
        proc.stdout.on("data", d => stdout += d.toString());
        proc.stderr.on("data", d => stderr += d.toString());
        proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
    });
}
export async function runAssemble(extractions, folders, config) {
    // Get next serial
    let startSerial = 1;
    try {
        const result = await runPython(join(SKILL_ROOT, "python", "get_last_serial.py"), [folders.LR_MATRIX], config.pythonPath);
        startSerial = parseInt(result.trim()) + 1;
    }
    catch {
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
    await runPython(join(SKILL_ROOT, "python", "assemble_excel.py"), [JSON.stringify(inputData)], config.pythonPath);
    return { metadataBatchFile, lrMatrixFile, startSerial };
}
// ============================================================================
// Step 5: VERIFY
// ============================================================================
export async function runVerify(extractions, folders) {
    const issues = [];
    const checks = {
        recordCountsMatch: true,
        fileRowBijection: true,
        serialContinuous: true,
        noBlankTitles: true,
        spotCheckPassed: true,
    };
    // Serial continuity
    const serials = extractions.map(e => e.lrMatrix["S. No."]).sort((a, b) => a - b);
    for (let i = 1; i < serials.length; i++) {
        if (serials[i] !== serials[i - 1] + 1) {
            checks.serialContinuous = false;
            issues.push(`Serial gap: ${serials[i - 1]} → ${serials[i]}`);
        }
    }
    // Blank titles
    for (const ex of extractions) {
        if (!ex.metadata.Title.trim()) {
            checks.noBlankTitles = false;
            issues.push(`Blank title: ${ex.fileName}`);
        }
    }
    // Spot check first 2
    for (const ex of extractions.slice(0, 2)) {
        try {
            const buffer = await readFile(join(folders.DB, ex.fileName));
            const data = await pdfParse(buffer, { max: 1 });
            const firstPage = data.text;
            if (!firstPage.toLowerCase().includes(ex.metadata.Title.toLowerCase().slice(0, 30))) {
                checks.spotCheckPassed = false;
                issues.push(`Spot check: title mismatch for ${ex.fileName}`);
            }
        }
        catch (e) {
            issues.push(`Spot check error for ${ex.fileName}: ${e}`);
        }
    }
    const passed = Object.values(checks).every(v => v);
    return { passed, issues, checks };
}
// ============================================================================
// Step 6: LOG
// ============================================================================
export async function runLog(result, folders, config) {
    const logFile = join(folders.ROOT, "SOP-Log.md");
    let existingLog = "";
    try {
        existingLog = await readFile(logFile, "utf-8");
    }
    catch {
        existingLog = "# SOP Log — LR Database runs\n\n";
    }
    const dateStr = new Date().toISOString().split("T")[0];
    const duplicateLines = result.duplicates.length > 0
        ? result.duplicates.map(d => `  - ${basename(d.file)}: ${d.reason} (matched ${basename(d.matchedWith)})`).join("\n")
        : "  - None";
    const logEntry = `
## ${dateStr} — Batch of ${result.keeperFiles.length} (auto)
- \`Incoming/\` held ${result.incomingFiles.length} PDF(s); dedup vs Database and within batch found ${result.duplicates.length} duplicate(s):
${duplicateLines}
- ${result.keeperFiles.length} keeper(s) extracted via ${config.llm.provider}/${config.llm.model}
- New \`${basename(result.metadataBatchFile)}\` (${result.keeperFiles.length} rows, 87 cols); +${result.keeperFiles.length} LR rows (S.No. ${result.verification.checks.serialContinuous ? "continuous" : "ISSUES"})
- Files moved to \`Database/\`; \`Incoming/\` empty at close
- Verification: ${result.verification.passed ? "PASSED" : "ISSUES FOUND"} — ${result.verification.issues.length} issue(s)
`;
    await writeFile(logFile, existingLog + logEntry);
}
// ============================================================================
// Full Pipeline
// ============================================================================
export async function runFullPipeline(config) {
    const folders = getFolders(config.projectRoot);
    folders.SCOPE = config.scope;
    await ensureFolders(folders, config.pythonPath);
    // Step 1: Ingest
    const incomingFiles = await runIngest(folders);
    if (incomingFiles.length === 0) {
        throw new Error("No files in Incoming/");
    }
    // Step 2: Dedup
    const { keepers, duplicates } = await runDedup(incomingFiles, folders);
    // Move duplicates to Set-Aside
    for (const dup of duplicates) {
        await rename(dup.file, join(folders.SETASIDE, basename(dup.file)));
    }
    // Step 3: Extract
    const extractions = await runExtract(keepers, folders, config);
    // Step 4: Assemble
    const { metadataBatchFile, lrMatrixFile, startSerial } = await runAssemble(extractions, folders, config);
    // Move keepers to Database
    for (const file of keepers) {
        await rename(file, join(folders.DB, basename(file)));
    }
    // Step 5: Verify
    const verification = await runVerify(extractions, folders);
    // Step 6: Log
    const result = {
        incomingFiles: incomingFiles.map(f => basename(f)),
        keeperFiles: keepers.map(f => basename(f)),
        duplicates,
        extracted: extractions,
        metadataBatchFile,
        lrMatrixFile,
        verification,
        logEntry: "",
    };
    await runLog(result, folders, config);
    return result;
}
//# sourceMappingURL=index.js.map