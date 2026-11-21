#!/usr/bin/env tsx
/** lr-dedup — Step 2: Deduplicate incoming against Database/ */

import { Command } from "commander";
import { glob } from "fast-glob";
import { readFile, rename, stat } from "fs/promises";
import { join, basename } from "path";
import { fileURLToPath } from "url";
import chalk from "chalk";
import ora from "ora";
import pdfParse from "pdf-parse";

const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..", "..");

const program = new Command();

program
  .name("lr-dedup")
  .description("Step 2: Deduplicate — Compare Incoming/ against Database/ and within batch")
  .version("0.1.0")
  .requiredOption("-r, --root <path>", "Project root folder")
  .action(async (opts) => {
    const root = opts.root;
    const incomingDir = join(root, "Incoming");
    const dbDir = join(root, "Database");
    const setAsideDir = join(root, "Set-Aside");

    const spinner = ora("Loading Database/ index...").start();

    try {
      // Get incoming files
      const incomingFiles = await glob("*.pdf", { cwd: incomingDir, absolute: true });
      if (incomingFiles.length === 0) {
        spinner.warn("No files in Incoming/");
        return;
      }

      // Get DB files
      const dbFiles = await glob("*.pdf", { cwd: dbDir, absolute: true });

      // Build DB index with first-page text
      const dbIndex = [];
      for (const file of dbFiles) {
        try {
          const buffer = await readFile(file);
          const data = await pdfParse(buffer, { max: 1 });
          dbIndex.push({
            file,
            name: basename(file),
            firstPage: data.text.slice(0, 3000).toLowerCase(),
            tokens: titleTokens(basename(file).replace(/\.pdf$/i, "")),
          });
        } catch {
          dbIndex.push({
            file,
            name: basename(file),
            firstPage: "",
            tokens: titleTokens(basename(file).replace(/\.pdf$/i, "")),
          });
        }
      }
      spinner.succeed(`Indexed ${dbIndex.length} Database files`);

      // Check each incoming file
      const keepers = [];
      const duplicates = [];

      for (const incFile of incomingFiles) {
        const incName = basename(incFile);
        const incTokens = titleTokens(incName.replace(/\.pdf$/i, ""));
        let matched = false;

        // Check against DB
        for (const db of dbIndex) {
          const sim = jaccard(incTokens, db.tokens);
          if (sim > 0.7) {
            const incBuffer = await readFile(incFile);
            const incData = await pdfParse(incBuffer, { max: 1 });
            const incFirstPage = incData.text.slice(0, 3000).toLowerCase();

            // DOI match
            const doiMatch = incFirstPage.match(/10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i);
            if (doiMatch && db.firstPage.includes(doiMatch[0].toLowerCase())) {
              matched = true;
              duplicates.push({ file: incFile, reason: "DOI match", matchedWith: db.file });
              break;
            }

            // Title/author overlap
            const titleWords = incTokens.filter(t => t.length > 4);
            const overlap = titleWords.filter(t => db.firstPage.includes(t)).length;
            if (overlap >= 3) {
              matched = true;
              duplicates.push({ file: incFile, reason: "Title/author match", matchedWith: db.file });
              break;
            }
          }
        }

        // Check within batch
        if (!matched) {
          for (const other of incomingFiles) {
            if (other === incFile) continue;
            const otherTokens = titleTokens(basename(other).replace(/\.pdf$/i, ""));
            if (jaccard(incTokens, otherTokens) > 0.85) {
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

      // Move duplicates to Set-Aside
      for (const dup of duplicates) {
        await rename(dup.file, join(setAsideDir, basename(dup.file)));
      }

      console.log(chalk.green(`\n✅ Keepers: ${keepers.length}, Duplicates moved to Set-Aside: ${duplicates.length}`));
      if (duplicates.length > 0) {
        console.log("\nDuplicates:");
        for (const d of duplicates) {
          console.log(`  ${basename(d.file)} → ${d.reason} (vs ${basename(d.matchedWith)})`);
        }
      }
      console.log("\nKeepers:");
      for (const k of keepers) console.log(`  ${basename(k)}`);

    } catch (error) {
      spinner.fail("Dedup failed");
      console.error(chalk.red(error));
      process.exit(1);
    }
  });

function titleTokens(title: string): string[] {
  return title.toLowerCase().replace(/[^\w\s]/g, " ").split(/\s+/).filter(t => t.length > 3);
}

function jaccard(a: string[], b: string[]): number {
  const setA = new Set(a), setB = new Set(b);
  const inter = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : inter / union;
}

program.parse();