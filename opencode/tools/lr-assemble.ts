#!/usr/bin/env tsx
/** lr-assemble — Step 4: Append extractions to Excel workbooks */

import { Command } from "commander";
import { readFile, writeFile } from "fs/promises";
import { join, resolve } from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import chalk from "chalk";
import ora from "ora";

const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..", "..");
const SKILL_ROOT = resolve(__dirname, "..", "..");

const program = new Command();

program
  .name("lr-assemble")
  .description("Step 4: Assemble — Append extractions to Excel workbooks with SOP formatting")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .requiredOption("-i, --input <path>", "Extractions JSON file")
  .option("--python <path>", "Python executable", "python")
  .action(async (opts) => {
    const root = opts.root;
    const inputFile = join(root, opts.input);
    const pythonPath = opts.python;

    const spinner = ora("Loading extractions...").start();

    try {
      const extractions = JSON.parse(await readFile(inputFile, "utf-8"));
      if (!extractions.length) {
        spinner.warn("No extractions to assemble");
        return;
      }

      // Determine next serial from LR Matrix
      const lrMatrixFile = join(root, "Outputs", "LR Matrix.xlsx");
      let startSerial = 1;
      try {
        const result = await runPython(
          join(SKILL_ROOT, "python", "get_last_serial.py"),
          [lrMatrixFile],
          pythonPath
        );
        startSerial = parseInt(result.trim()) + 1;
      } catch {
        startSerial = 1;
      }

      // Assign serials
      extractions.forEach((ex: any, i: number) => {
        ex.lrMatrix["S. No."] = startSerial + i;
      });

      // Prepare input for Python assembler
      const inputData = {
        extractions,
        metadataBatchFile: join(root, "Outputs", `metadata batch ${new Date().toISOString().split("T")[0]}.xlsx`),
        lrMatrixFile,
        startSerial,
        metadataColumns: [
          "File Name", "Item Type", "Publication Year", "Author", "Title",
          "Publication Title", "ISBN", "ISSN", "DOI", "Url", "Abstract Note",
          "Date", "Date Added", "Date Modified", "Access Date", "Pages",
          "Num Pages", "Issue", "Volume", "Number Of Volumes", "Journal Abbreviation",
          "Short Title", "Series", "Series Number", "Series Text", "Series Title",
          "Publisher", "Place", "Language", "Rights", "Type", "Archive",
          "Archive Location", "Library Catalog", "Call Number", "Extra", "Notes",
          "File Attachments", "Link Attachments", "Manual Tags", "Automatic Tags",
          "Editor", "Series Editor", "Translator", "Contributor", "Attorney Agent",
          "Book Author", "Cast Member", "Commenter", "Composer", "Cosponsor",
          "Counsel", "Interviewer", "Producer", "Recipient", "Reviewed Author",
          "Scriptwriter", "Words By", "Guest", "Number", "Edition", "Running Time",
          "Scale", "Medium", "Artwork Size", "Filing Date", "Application Number",
          "Assignee", "Issuing Authority", "Country", "Meeting Name",
          "Conference Name", "Court", "References", "Reporter", "Legal Status",
          "Priority Numbers", "Programming Language", "Version", "System", "Code",
          "Code Number", "Section", "Session", "Committee", "History",
          "Legislative Body"
        ],
        lrMatrixColumns: [
          "S. No.", "File Name", "Title", "Author", "Year", "Journal", "DOI",
          "Objective", "Methodology", "Key Finding", "Discussion", "Limitation",
          "Theoretical Framework", "Tag 1 (Method)", "Tag 2 (Type of Research)",
          "Independent Variable", "Dependent Variable", "Other Variable",
          "Decision (For Me)", "Reason (For Me)"
        ],
      };

      spinner.text = "Assembling Excel workbooks...";
      await runPython(
        join(SKILL_ROOT, "python", "assemble_excel.py"),
        [JSON.stringify(inputData)],
        pythonPath
      );

      spinner.succeed(`Assembled ${extractions.length} records (S.No. ${startSerial}–${startSerial + extractions.length - 1})`);

    } catch (error) {
      spinner.fail("Assemble failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

function runPython(script: string, args: string[], pythonPath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const proc = spawn(pythonPath, [script, ...args], { cwd: SKILL_ROOT });
    let stdout = "", stderr = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => stderr += d.toString());
    proc.on("close", code => code === 0 ? resolve(stdout) : reject(new Error(stderr)));
  });
}

program.parse();