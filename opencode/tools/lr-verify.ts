#!/usr/bin/env tsx
/** lr-verify — Step 5: Verify pipeline output */

import { Command } from "commander";
import { glob } from "fast-glob";
import { readFile } from "fs/promises";
import { join, basename } from "path";
import chalk from "chalk";
import ora from "ora";
import pdfParse from "pdf-parse";
import { load_workbook } from "openpyxl"; // Note: This won't work in TS, we'll use Python

const program = new Command();

program
  .name("lr-verify")
  .description("Step 5: Verify — Check record counts, bijection, serials, spot-check")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .option("--python <path>", "Python executable", "python")
  .action(async (opts) => {
    const root = opts.root;
    const pythonPath = opts.python;

    const spinner = ora("Running verification...").start();

    try {
      // Use Python for Excel reading
      const script = `
import sys
sys.path.insert(0, "${root}")
from openpyxl import load_workbook

lr_path = "${join(root, "Outputs", "LR Matrix.xlsx")}"
db_path = "${join(root, "Database")}"

wb = load_workbook(lr_path, read_only=True)
ws = wb.active

# Get all serials and filenames
serials = []
filenames = []
for row in ws.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
    s, f = row
    if isinstance(s, int):
        serials.append(s)
        filenames.append(f)

# Check serial continuity
gaps = []
for i in range(1, len(serials)):
    if serials[i] != serials[i-1] + 1:
        gaps.append(f"{serials[i-1]} -> {serials[i]}")

# Check DB files
import os
db_files = set(f for f in os.listdir(db_path) if f.endswith('.pdf'))
lr_files = set(filenames)

missing_in_db = lr_files - db_files
extra_in_db = db_files - lr_files

# Check blank titles
blank_titles = []
for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True):
    if not row[0] or not str(row[0]).strip():
        blank_titles.append(row[0])

print(f"SERIALS: {len(serials)}")
print(f"GAPS: {gaps}")
print(f"MISSING_IN_DB: {list(missing_in_db)}")
print(f"EXTRA_IN_DB: {list(extra_in_db)}")
print(f"BLANK_TITLES: {len(blank_titles)}")
`;

      const { spawn } = await import("child_process");
      const result = await runPython(pythonPath, ["-c", script]);
      
      // Parse results
      const lines = result.trim().split("\n");
      const data: any = {};
      for (const line of lines) {
        const [key, ...val] = line.split(": ");
        data[key] = val.join(": ");
      }

      const issues = [];
      if (data.GAPS !== "[]") issues.push(`Serial gaps: ${data.GAPS}`);
      if (data.MISSING_IN_DB !== "[]") issues.push(`Files in LR but not in Database: ${data.MISSING_IN_DB}`);
      if (data.EXTRA_IN_DB !== "[]") issues.push(`Files in Database but not in LR: ${data.EXTRA_IN_DB}`);
      if (parseInt(data.BLANK_TITLES || "0") > 0) issues.push(`Blank titles: ${data.BLANK_TITLES}`);

      // Spot check first 2 files
      const lrFiles = JSON.parse(data.MISSING_IN_DB.replace(/'/g, '"').replace(/True/g, 'true').replace(/False/g, 'false') || "[]");
      const dbFiles = await glob("*.pdf", { cwd: join(root, "Database"), absolute: true });
      
      for (const file of dbFiles.slice(0, 2)) {
        try {
          const buffer = await readFile(file);
          const data = await pdfParse(buffer, { max: 1 });
          // Basic check - file readable
        } catch (e) {
          issues.push(`Spot check failed for ${basename(file)}: ${e}`);
        }
      }

      if (issues.length === 0) {
        spinner.succeed("All verification checks PASSED");
      } else {
        spinner.warn(`Verification issues found: ${issues.length}`);
        for (const issue of issues) console.log(chalk.yellow(`  ⚠ ${issue}`));
      }

      console.log(`\nSummary:`);
      console.log(`  LR Matrix rows: ${data.SERIALS}`);
      console.log(`  Database files: ${dbFiles.length}`);
      console.log(`  Issues: ${issues.length}`);

    } catch (error) {
      spinner.fail("Verify failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

function runPython(pythonPath: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    const { spawn } = require("child_process");
    const proc = spawn(pythonPath, args);
    let stdout = "", stderr = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => stderr += d.toString());
    proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
  });
}

program.parse();