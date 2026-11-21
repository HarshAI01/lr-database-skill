/**
 * LR Database Skill — Main Entry Point
 * 
 * This module exports all core functionality for programmatic use.
 * The CLI tools are in opencode/tools/ and are registered via skill.json.
 */

// Export schemas (includes ExtractionResult type)
export * from "./schemas.js";

// Export pipeline (includes PipelineConfig, BatchResult, etc. interfaces)
export * from "./pipeline/index.js";

// Version
export const VERSION = "0.1.0";
export const SOP_VERSION = "2026-10"; // Matches SOP.md date