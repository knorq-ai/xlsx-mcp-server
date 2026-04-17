/**
 * Server-side safety guards for F-002 (structural integrity).
 *
 * Two enforcement layers, both configured via environment variables at server
 * start so the LLM cannot disable them by setting tool parameters:
 *
 *  - XLSX_MAX_CELLS_PER_CALL: hard cap on cells touched by a single tool call.
 *  - XLSX_TEMPLATE_MODE + XLSX_TEMPLATE_RANGES: whitelist mode that rejects
 *    writes outside declared "Sheet!Range" entries.
 */
import { ErrorCode, EngineError } from "./xlsx-io.js";
import {
  type CellRange,
  parseRange,
  rangeToString,
  MAX_RANGE_CELLS,
} from "./cells.js";

function parsePositiveInt(name: string, raw: string | undefined, fallback: number): number {
  if (raw === undefined || raw === "") return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 1) {
    throw new Error(`Invalid ${name}: must be a positive integer, got ${JSON.stringify(raw)}`);
  }
  return n;
}

function parseBoolFlag(raw: string | undefined): boolean {
  if (raw === undefined) return false;
  const v = raw.toLowerCase();
  return v === "1" || v === "true" || v === "yes" || v === "on";
}

export interface TemplateRange {
  sheet: string;
  range: CellRange;
}

function parseTemplateRanges(raw: string | undefined): TemplateRange[] {
  if (!raw) return [];
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((spec) => {
      const bang = spec.indexOf("!");
      if (bang < 1 || bang === spec.length - 1) {
        throw new Error(
          `Invalid XLSX_TEMPLATE_RANGES entry ${JSON.stringify(spec)}. Expected "Sheet!Range" e.g. "Sheet1!A1:D10".`,
        );
      }
      const sheet = spec.slice(0, bang);
      const rangeStr = spec.slice(bang + 1);
      return { sheet, range: parseRange(rangeStr) };
    });
}

export interface SafetyConfig {
  maxCellsPerCall: number;
  templateMode: boolean;
  templateRanges: TemplateRange[];
}

export function loadSafetyConfig(env: NodeJS.ProcessEnv = process.env): SafetyConfig {
  const maxCellsPerCall = parsePositiveInt(
    "XLSX_MAX_CELLS_PER_CALL",
    env.XLSX_MAX_CELLS_PER_CALL,
    MAX_RANGE_CELLS,
  );
  const templateMode = parseBoolFlag(env.XLSX_TEMPLATE_MODE);
  const templateRanges = parseTemplateRanges(env.XLSX_TEMPLATE_RANGES);
  if (templateMode && templateRanges.length === 0) {
    throw new Error(
      "XLSX_TEMPLATE_MODE is enabled but XLSX_TEMPLATE_RANGES is empty. " +
        'Set XLSX_TEMPLATE_RANGES to a CSV of "Sheet!Range" entries (e.g. "Sheet1!A1:D10,Sheet1!F2:F100").',
    );
  }
  return { maxCellsPerCall, templateMode, templateRanges };
}

export function assertCellCount(cells: number, tool: string, cfg: SafetyConfig): void {
  if (cells > cfg.maxCellsPerCall) {
    throw new EngineError(
      ErrorCode.MAX_CELLS_EXCEEDED,
      `${tool} request of ${cells.toLocaleString()} cells exceeds XLSX_MAX_CELLS_PER_CALL=${cfg.maxCellsPerCall.toLocaleString()}.`,
    );
  }
}

function rangeContains(outer: CellRange, inner: CellRange): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startCol >= outer.startCol &&
    inner.endCol <= outer.endCol
  );
}

export function assertWithinTemplate(
  sheet: string | number,
  target: CellRange,
  cfg: SafetyConfig,
): void {
  if (!cfg.templateMode) return;
  if (typeof sheet === "number") {
    throw new EngineError(
      ErrorCode.OUTSIDE_TEMPLATE_RANGE,
      "Template mode requires sheet specified by name, not numeric index.",
    );
  }
  const candidates = cfg.templateRanges.filter((t) => t.sheet === sheet);
  if (candidates.length === 0) {
    throw new EngineError(
      ErrorCode.OUTSIDE_TEMPLATE_RANGE,
      `Sheet ${JSON.stringify(sheet)} has no declared template range. Set XLSX_TEMPLATE_RANGES to enable writes here.`,
    );
  }
  const ok = candidates.some((t) => rangeContains(t.range, target));
  if (!ok) {
    const allowed = candidates.map((c) => rangeToString(c.range)).join(", ");
    throw new EngineError(
      ErrorCode.OUTSIDE_TEMPLATE_RANGE,
      `Write target ${rangeToString(target)} on sheet ${JSON.stringify(sheet)} is outside declared template ranges (${allowed}).`,
    );
  }
}
