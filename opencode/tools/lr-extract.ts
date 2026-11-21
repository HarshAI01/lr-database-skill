#!/usr/bin/env tsx
/** lr-extract — Step 3: Read PDFs and return content for agent to extract */

import { Command } from "commander";
import { glob } from "fast-glob";
import { readFile, writeFile } from "fs/promises";
import { join, basename } from "path";
import chalk from "chalk";
import ora from "ora";
import pdfParse from "pdf-parse";

const program = new Command();

program
  .name("lr-extract")
  .description("Step 3: Extract — Read PDFs and return text/content for agent to extract metadata")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .option("-o, --output <path>", "Output JSON file", "pdf-contents.json")
  .action(async (opts) => {
    const root = opts.root;
    const incomingDir = join(root, "Incoming");
    const outputFile = join(root, opts.output);

    const spinner = ora(`Reading PDFs from ${incomingDir}...`).start();

    try {
      const files = await glob("*.pdf", { cwd: incomingDir, absolute: true });
      if (files.length === 0) {
        spinner.warn("No PDFs in Incoming/");
        return;
      }

      const results = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fileName = basename(file);
        spinner.text = `Reading ${i + 1}/${files.length}: ${fileName}`;

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
            fullText: fullText.slice(0, 50000),  // First 50k chars
            firstPage: firstPage.slice(0, 3000), // First 3k chars
          });

        } catch (error) {
          spinner.warn(`Failed to read ${fileName}: ${error}`);
          results.push({ fileName, error: String(error) });
        }
      }

      await writeFile(outputFile, JSON.stringify(results, null, 2));
      spinner.succeed(`Read ${results.length} PDFs → ${outputFile}`);
      console.log(chalk.gray("\nAgent will now extract metadata from each PDF using its model."));

    } catch (error) {
      spinner.fail("Extract failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

program.parse();