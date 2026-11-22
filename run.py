#!/usr/bin/env python3
"""
LR Database - Simple Runner
Run this from your project folder: python path/to/lr-database-skill/run.py
"""

import sys
import os
import subprocess

# Add the core src to path
SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
CORE_SRC = os.path.join(SKILL_DIR, "core", "src")
PYTHON_DIR = os.path.join(SKILL_DIR, "python")

sys.path.insert(0, CORE_SRC)

from pipeline import (
    getFolders, ensureFolders, runIngest, runDedup, runExtractPDFs,
    runAssemble, runVerify, runLog, EXTRACTION_PROMPT
)

def main():
    print("=" * 60)
    print("LR DATABASE - Literature Review Pipeline")
    print("=" * 60)
    
    # Get project root (current directory)
    project_root = os.getcwd()
    print(f"\nProject: {project_root}")
    
    # Check for Incoming folder
    incoming = os.path.join(project_root, "Incoming")
    if not os.path.exists(incoming):
        print(f"\nCreating Incoming folder...")
        os.makedirs(incoming, exist_ok=True)
    
    # Check for PDFs
    pdfs = [f for f in os.listdir(incoming) if f.lower().endswith('.pdf')]
    if not pdfs:
        print(f"\nNo PDFs found in {incoming}")
        print("Please add PDF files to the Incoming folder and run again.")
        return
    
    print(f"\nFound {len(pdfs)} PDF(s) in Incoming/")
    
    # Run pipeline
    config = {
        "projectRoot": project_root,
        "scope": "",
        "pythonPath": "python"
    }
    
    try:
        import asyncio
        asyncio.run(run_pipeline(config))
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

async def run_pipeline(config):
    folders = getFolders(config["projectRoot"])
    await ensureFolders(folders, config["pythonPath"])
    
    # Step 1: Ingest
    print("\n[1/6] Ingesting PDFs...")
    incoming_files = await runIngest(folders)
    if not incoming_files:
        print("  No PDFs to process.")
        return
    print(f"  Found {len(incoming_files)} PDF(s)")
    
    # Step 2: Dedup
    print("\n[2/6] Checking for duplicates...")
    result = await runDedup(incoming_files, folders)
    keepers = result["keepers"]
    duplicates = result["duplicates"]
    
    for dup in duplicates:
        os.rename(dup["file"], os.path.join(folders["SETASIDE"], os.path.basename(dup["file"])))
    
    print(f"  Keepers: {len(keepers)}, Duplicates: {len(duplicates)}")
    for d in duplicates:
        print(f"    Duplicate: {os.path.basename(d['file'])} -> {d['reason']}")
    
    if not keepers:
        print("  No new papers to process.")
        return
    
    # Step 3: Extract PDF content
    print("\n[3/6] Reading PDF content...")
    pdf_contents = await runExtractPDFs(keepers)
    print(f"  Read {len(pdf_contents)} papers")
    
    # Step 3b: Agent extraction (interactive)
    print("\n[3b/6] AI Extraction Required")
    print("  For each paper, you'll need to:")
    print("  1. Copy the prompt and paper text below")
    print("  2. Paste into your AI (ChatGPT, Claude, etc.)")
    print("  3. Copy the JSON result and paste it here")
    print()
    
    extractions = []
    for pdf in pdf_contents:
        print("=" * 60)
        print(f"PAPER: {pdf['fileName']}")
        print("=" * 60)
        print(f"\nFIRST PAGE (for metadata):")
        print(pdf['firstPage'][:1000])
        print(f"\n... (showing first 1000 chars)")
        print(f"\nFULL TEXT (first 30000 chars):")
        print(pdf['fullText'][:3000])
        print(f"\n... (showing first 3000 chars)")
        print("\n" + "-" * 60)
        print("EXTRACTION PROMPT (copy this + paper text to your AI):")
        print("-" * 60)
        print(EXTRACTION_PROMPT)
        print("-" * 60)
        
        # Get JSON from user
        while True:
            user_input = input("\nPaste JSON result (or 'skip' to skip): ").strip()
            if user_input.lower() == 'skip':
                from schemas import createEmptyMetadata, createEmptyLRMatrix
                meta = createEmptyMetadata(pdf['fileName'])
                meta["Num Pages"] = pdf['numPages']
                lr = createEmptyLRMatrix(0, pdf['fileName'])
                extractions.append({"fileName": pdf['fileName'], "metadata": meta, "lrMatrix": lr})
                break
            try:
                parsed = eval(user_input) if user_input.startswith('{') else eval('{' + user_input + '}')
                if isinstance(parsed, dict) and 'metadata' in parsed and 'lrMatrix' in parsed:
                    parsed['fileName'] = pdf['fileName']
                    extractions.append(parsed)
                    print("  OK - Extraction saved")
                    break
                else:
                    print("  Invalid format. Must have 'metadata' and 'lrMatrix' keys.")
            except Exception as e:
                print(f"  Error parsing JSON: {e}")
                print("  Try again or type 'skip'")
    
    # Step 4: Assemble
    print("\n[4/6] Assembling Excel workbooks...")
    result = await runAssemble(extractions, folders, config)
    print(f"  Created: {os.path.basename(result['metadataBatchFile'])}")
    print(f"  Updated: {os.path.basename(result['lrMatrixFile'])} (S.No. {result['startSerial']}-{result['startSerial']+len(extractions)-1})")
    
    # Move keepers to Database
    for file in keepers:
        os.rename(file, os.path.join(folders["DB"], os.path.basename(file)))
    
    # Step 5: Verify
    print("\n[5/6] Verifying...")
    verification = await runVerify(extractions, folders)
    if verification["passed"]:
        print("  All checks passed!")
    else:
        print(f"  Issues found: {len(verification['issues'])}")
        for issue in verification["issues"]:
            print(f"    - {issue}")
    
    # Step 6: Log
    print("\n[6/6] Logging...")
    batch_result = {
        "incomingFiles": [os.path.basename(f) for f in incoming_files],
        "keeperFiles": [os.path.basename(f) for f in keepers],
        "duplicates": duplicates,
        "extracted": extractions,
        "metadataBatchFile": result["metadataBatchFile"],
        "lrMatrixFile": result["lrMatrixFile"],
        "verification": verification,
    }
    await runLog(batch_result, folders)
    print("  Logged to SOP-Log.md")
    
    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE!")
    print("=" * 60)
    print(f"Processed: {len(extractions)} papers")
    print(f"Duplicates: {len(duplicates)}")
    print(f"LR Matrix: S.No. {result['startSerial']}-{result['startSerial']+len(extractions)-1}")
    print(f"Outputs: {folders['OUT']}/")
    print(f"Log: {os.path.join(folders['ROOT'], 'SOP-Log.md')}")

if __name__ == "__main__":
    main()