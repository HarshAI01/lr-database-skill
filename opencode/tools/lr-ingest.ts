#!/usr/bin/env tsx
/** lr-ingest — Step 1: Scan Incoming/, normalize filenames */

import { Command } from "commander";
import { glob } from "fast-glob";
import { readFile, writeFile, rename } from "fs/promises";
import { join, basename } from "path";
import chalk from "chalk";
import ora from "ora";

const program = new Command();

program
  .name("lr-ingest")
  .description("Step 1: Ingest — Scan Incoming/ and normalize filenames")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .action(async (opts) => {
    const root = opts.root;
    const incomingDir = join(root, "Incoming");
    const spinner = ora("Scanning Incoming/").start();

    try {
      const pdfFiles = await glob("*.pdf", { cwd: incomingDir, absolute: true });
      
      if (pdfFiles.length === 0) {
        spinner.warn("No PDFs found in Incoming/");
        return;
      }

      let normalized = 0;
      for (const file of pdfFiles) {
        const name = basename(file);
        // Normalize: remove illegal chars, collapse whitespace
        const normalizedName = name
          .replace(/[<>:"|?*]/g, "")
          .replace(/\s+/g, " ")
          .trim();
        
        if (name !== normalizedName) {
          await rename(file, join(incomingDir, normalizedName));
          normalized++;
        }
      }

      spinner.succeed(`Found ${pdfFiles.length} PDF(s), normalized ${normalized} filename(s)`);
      console.log("\nFiles ready for deduplication:");
      const finalFiles = await glob("*.pdf", { cwd: incomingDir });
      for (const f of finalFiles) console.log(`  ${f}`);

    } catch (error) {
      spinner.fail("Ingest failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

program.parse();