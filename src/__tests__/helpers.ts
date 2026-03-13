/**
 * Test helpers for the MCP xlsx-engine tests.
 */

import * as os from "os";
import * as path from "path";
import * as fs from "fs/promises";
import * as crypto from "crypto";
import { createWorkbook, EngineError, ErrorCode } from "../xlsx-engine.js";
import type { ErrorCodeType } from "../xlsx-engine.js";

export { EngineError, ErrorCode };
export type { ErrorCodeType };

/** Generate a unique tmp file path for a .xlsx file */
export function tmpXlsxPath(): string {
  return path.join(os.tmpdir(), `mcp-test-${crypto.randomUUID()}.xlsx`);
}

/** List of tmp file paths to clean up after each test */
const tmpFiles: string[] = [];

/** Register a tmp path for cleanup */
export function trackTmpFile(p: string): string {
  tmpFiles.push(p);
  return p;
}

/** Remove all tracked tmp files */
export async function cleanupTmpFiles(): Promise<void> {
  for (const p of tmpFiles) {
    try {
      await fs.unlink(p);
    } catch {
      // ignore — file may not exist
    }
  }
  tmpFiles.length = 0;
}

/** Create a tmp xlsx and return its path (auto-tracked for cleanup) */
export async function createTmpWorkbook(
  sheetName?: string,
): Promise<string> {
  const p = tmpXlsxPath();
  trackTmpFile(p);
  await createWorkbook(p, sheetName);
  return p;
}
