#!/usr/bin/env tsx
/** lr-log — Step 6: Append entry to SOP-Log.md */

import { Command } from "commander";
import { readFile, writeFile } from "fs/promises";
import { join } from "path";
import chalk from "chalk";
import ora from "ora";

const program = new Command();

program
  .name("lr-log")
  .description("Step 6: Log — Append structured entry to SOP-Log.md")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .requiredOption("-k, --keepers <count>", "Number of keeper files")
  .requiredOption("-d, --duplicates <count>", "Number of duplicates")
  .requiredOption("--provider <provider>", "LLM provider used")
  .requiredOption("--model <model>", "LLM model used")
  .option("--serial-start <n>", "Starting serial number", "1")
  .option("--serial-end <n>", "Ending serial number", "1")
  .option("--verification <status>", "passed|failed", "passed")
  .option("--issues <count>", "Number of verification issues", "0")
  .action(async (opts) => {
    const root = opts.root;
    const logFile = join(root, "SOP-Log.md");

    const spinner = ora("Writing SOP-Log.md...").start();

    try {
      let existingLog = "";
      try {
        existingLog = await readFile(logFile, "utf-8");
      } catch {
        existingLog = "# SOP Log — LR Database runs\n\n";
      }

      const dateStr = new Date().toISOString().split("T")[0];
      const keepers = parseInt(opts.keepers);
      const duplicates = parseInt(opts.duplicates);
      const serialStart = parseInt(opts.serialStart);
      const serialEnd = parseInt(opts.serialEnd);

      const logEntry = `
## ${dateStr} — Batch of ${keepers} (auto)
- \`Incoming/\` held ${keepers + duplicates} PDF(s); dedup vs Database and within batch found ${duplicates} duplicate(s)
- ${keepers} keeper(s) extracted via ${opts.provider}/${opts.model}
- New \`Outputs/metadata batch ${dateStr}.xlsx\` (${keepers} rows, 87 cols); +${keepers} LR rows (S.No. ${serialStart}–${serialEnd})
- Files moved to \`Database/\`; \`Incoming/\` empty at close
- Verification: ${opts.verification.toUpperCase()} — ${opts.issues} issue(s)
`;

      await writeFile(logFile, existingLog + logEntry);
      spinner.succeed(`Logged to ${logFile}`);

    } catch (error) {
      spinner.fail("Log failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

program.parse();