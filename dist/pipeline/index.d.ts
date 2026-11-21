/**
 * Pipeline Orchestrator — Core logic shared by CLI and programmatic use
 */
import { ExtractionResult } from "../schemas.js";
export type { MetadataRecord, LRMatrixRecord } from "../schemas.js";
export interface PipelineConfig {
    projectRoot: string;
    scope: string;
    llm: {
        provider: "ollama" | "openai" | "anthropic" | "gemini";
        model: string;
        apiKey?: string;
        baseUrl?: string;
        temperature: number;
        maxTokens: number;
    };
    pythonPath: string;
}
export interface FolderPaths {
    ROOT: string;
    IN: string;
    DB: string;
    OUT: string;
    SETASIDE: string;
    REF: string;
    SCOPE: string;
    LR_MATRIX: string;
    META0: string;
}
export interface BatchResult {
    incomingFiles: string[];
    keeperFiles: string[];
    duplicates: Array<{
        file: string;
        reason: string;
        matchedWith: string;
    }>;
    extracted: ExtractionResult[];
    metadataBatchFile: string;
    lrMatrixFile: string;
    verification: VerificationResult;
    logEntry: string;
}
export interface VerificationResult {
    passed: boolean;
    issues: string[];
    checks: {
        recordCountsMatch: boolean;
        fileRowBijection: boolean;
        serialContinuous: boolean;
        noBlankTitles: boolean;
        spotCheckPassed: boolean;
    };
}
export declare function getFolders(root: string): FolderPaths;
export declare function ensureFolders(folders: FolderPaths, pythonPath?: string): Promise<void>;
export declare function runIngest(folders: FolderPaths): Promise<string[]>;
export declare function runDedup(incomingFiles: string[], folders: FolderPaths): Promise<{
    keepers: string[];
    duplicates: Array<{
        file: string;
        reason: string;
        matchedWith: string;
    }>;
}>;
export declare function runExtract(keeperFiles: string[], folders: FolderPaths, config: PipelineConfig): Promise<ExtractionResult[]>;
export declare function runAssemble(extractions: ExtractionResult[], folders: FolderPaths, config: PipelineConfig): Promise<{
    metadataBatchFile: string;
    lrMatrixFile: string;
    startSerial: number;
}>;
export declare function runVerify(extractions: ExtractionResult[], folders: FolderPaths): Promise<VerificationResult>;
export declare function runLog(result: BatchResult, folders: FolderPaths, config: PipelineConfig): Promise<void>;
export declare function runFullPipeline(config: PipelineConfig): Promise<BatchResult>;
//# sourceMappingURL=index.d.ts.map