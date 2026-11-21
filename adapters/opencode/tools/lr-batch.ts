#!/usr/bin/env tsx
/** lr-batch — Full LR Pipeline for OpenCode */

import { Command } from "commander";
import { readFile, writeFile } from "fs/promises";
import { join, resolve } from "path";
import { fileURLToPath } from "url";
import chalk from "chalk";
import ora from "ora";
import { runIngest, runDedup, runExtractPDFs, runAssemble, runVerify, runLog, getFolders, ensureFolders, EXTRACTION_PROMPT, type PipelineConfig, type ExtractionResult, type PDFContent } from "../../../../core/src/pipeline.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..");
const CORE_ROOT = resolve(__dirname, "..", "..", "..", "core");

const program = new Command();

program
  .name("lr-batch")
  .description("Literature Review Database — Full SOP Pipeline")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .option("-s, --scope <text>", "Screening scope (optional)")
  .option("--python <path>", "Python executable", "python")
  .action(async (opts) => {
    const config: PipelineConfig = {
      projectRoot: resolve(opts.root),
      scope: opts.scope || "",
      pythonPath: opts.python,
    };

    const spinner = ora("Starting LR Database Pipeline...").start();
    const folders = getFolders(config.projectRoot);

    try {
      await ensureFolders(folders, config.pythonPath);

      // Step 1: Ingest
      spinner.text = "Step 1/6: Ingesting PDFs...";
      const incomingFiles = await runIngest(folders);
      if (incomingFiles.length === 0) {
        spinner.warn("No PDFs in Incoming/");
        return;
      }

      // Step 2: Dedup
      spinner.text = "Step 2/6: Deduplicating...";
      const { keepers, duplicates } = await runDedup(incomingFiles, folders);
      for (const dup of duplicates) {
        const { rename } = await import("fs/promises");
        await rename(dup.file, join(folders.SETASIDE, basename(dup.file)));
      }

      // Step 3: Extract PDF content
      spinner.text = "Step 3/6: Reading PDF content...";
      const pdfContents = await runExtractPDFs(keepers);

      // Step 3b: Agent extracts using their model
      spinner.text = "Step 3b/6: Extracting metadata (using your OpenCode model)...";
      const extractions = await extractWithAgent(pdfContents, config);

      // Step 4: Assemble
      spinner.text = "Step 4/6: Assembling Excel workbooks...";
      const { metadataBatchFile, lrMatrixFile, startSerial } = await runAssemble(extractions, folders, config);

      // Move keepers to Database
      for (const file of keepers) {
        const { rename } = await import("fs/promises");
        await rename(file, join(folders.DB, basename(file)));
      }

      // Step 5: Verify
      spinner.text = "Step 5/6: Verifying...";
      const verification = await runVerify(extractions, folders);

      // Step 6: Log
      spinner.text = "Step 6/6: Logging...";
      const result = {
        incomingFiles: incomingFiles.map(basename),
        keeperFiles: keepers.map(basename),
        duplicates,
        extracted: extractions,
        metadataBatchFile,
        lrMatrixFile,
        verification,
      };
      await runLog(result, folders);

      spinner.succeed("Pipeline complete!");
      console.log(chalk.green(`\n✅ Processed: ${result.keeperFiles.length} papers`));
      console.log(chalk.green(`📊 LR Matrix: S.No. ${startSerial}–${startSerial + result.keeperFiles.length - 1}`));
      console.log(chalk.green(`📝 Log: SOP-Log.md updated`));

    } catch (error) {
      spinner.fail("Pipeline failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

async function extractWithAgent(pdfContents: PDFContent[], config: PipelineConfig): Promise<ExtractionResult[]> {
  const { default: readline } = await import("readline");
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  
  const extractions: ExtractionResult[] = [];
  
  for (const pdf of pdfContents) {
    console.log(chalk.cyan(`\n📄 Extracting: ${pdf.fileName}`));
    console.log(chalk.gray("─".repeat(60)));
    console.log(chalk.yellow("FIRST PAGE:"));
    console.log(pdf.firstPage);
    console.log(chalk.gray("─".repeat(60)));
    console.log(chalk.yellow("FULL TEXT (first 50k chars):"));
    console.log(pdf.fullText);
    console.log(chalk.gray("─".repeat(60)));
    
    // Provide extraction prompt
    console.log(chalk.magenta("\n📋 EXTRACTION PROMPT (copy and run with your model):"));
    console.log(chalk.gray(EXTRACTION_PROMPT));
    console.log(chalk.gray("─".repeat(60)));
    
    const result = await new Promise<ExtractionResult>((resolve) => {
      rl.question(chalk.green("\n✅ Paste JSON result (or 'skip'): "), (input) => {
        if (input.trim().toLowerCase() === "skip") {
          const { createEmptyMetadata, createEmptyLRMatrix } = require("../../../../core/src/schemas.js");
          const meta = createEmptyMetadata(pdf.fileName);
          meta["Num Pages"] = pdf.numPages;
          const lr = createEmptyLRMatrix(0, pdf.fileName);
          resolve({ fileName: pdf.fileName, metadata: meta, lrMatrix: lr });
        } else {
          try {
            const parsed = JSON.parse(input);
            resolve({ fileName: pdf.fileName, metadata: parsed.metadata, lrMatrix: parsed.lrMatrix });
          } catch (e) {
            console.error(chalk.red("Invalid JSON, skipping..."));
            const { createEmptyMetadata, createEmptyLRMatrix } = require("../../../../core/src/schemas.js");
            const meta = createEmptyMetadata(pdf.fileName);
            meta["Num Pages"] = pdf.numPages;
            const lr = createEmptyLRMatrix(0, pdf.fileName);
            resolve({ fileName: pdf.fileName, metadata: meta, lrMatrix: lr });
          }
        }
      });
    });
    
    extractions.push(result);
  }
  
  rl.close();
  return extractions;
}

function basename(path: string): string {
  return path.split(/[/\\]/).pop() || path;
}

program.parse();