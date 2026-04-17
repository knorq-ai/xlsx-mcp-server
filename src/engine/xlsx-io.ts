/**
 * XLSX file I/O — ExcelJS Workbook wrapper, error types, sheet resolution.
 */

import * as fs from "fs/promises";
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// ExcelJS patch — tolerate duplicate/overlapping merge cells
// ---------------------------------------------------------------------------
// Some XLSX files (especially from Google Sheets) contain duplicate or
// overlapping <mergeCell> entries.  ExcelJS throws "Cannot merge already
// merged cells" in _mergeCellsInternal when this happens.  We patch
// _parseMergeCells to silently skip duplicates instead of crashing.
// ---------------------------------------------------------------------------

const WorksheetProto = (ExcelJS as unknown as Record<string, unknown>).Worksheet
  ? ((ExcelJS as unknown as Record<string, unknown>).Worksheet as { prototype: Record<string, unknown> }).prototype
  : null;

// ExcelJS doesn't export Worksheet directly — grab it from a temp workbook.
function getWorksheetProto(): Record<string, unknown> {
  if (WorksheetProto) return WorksheetProto;
  const tmp = new ExcelJS.Workbook();
  const ws = tmp.addWorksheet("_tmp_");
  const proto = Object.getPrototypeOf(ws) as Record<string, unknown>;
  tmp.removeWorksheet(ws.id);
  return proto;
}

const _wsProto = getWorksheetProto();
const _origParseMergeCells = _wsProto._parseMergeCells as
  | ((model: Record<string, unknown>) => void)
  | undefined;

if (typeof _origParseMergeCells === "function") {
  _wsProto._parseMergeCells = function patchedParseMergeCells(
    this: ExcelJS.Worksheet,
    model: Record<string, unknown>,
  ) {
    const merges = model.mergeCells;
    if (!Array.isArray(merges)) return;
    for (const merge of merges) {
      try {
        (this as unknown as { mergeCellsWithoutStyle: (...args: unknown[]) => void })
          .mergeCellsWithoutStyle(merge);
      } catch {
        // Skip overlapping / duplicate merge — not fatal for reading.
      }
    }
  };
}

// ---------------------------------------------------------------------------
// Error types
// ---------------------------------------------------------------------------

export const ErrorCode = {
  FILE_NOT_FOUND: "FILE_NOT_FOUND",
  INVALID_XLSX: "INVALID_XLSX",
  SHEET_NOT_FOUND: "SHEET_NOT_FOUND",
  CELL_OUT_OF_RANGE: "CELL_OUT_OF_RANGE",
  INVALID_RANGE: "INVALID_RANGE",
  ROW_OUT_OF_RANGE: "ROW_OUT_OF_RANGE",
  COLUMN_OUT_OF_RANGE: "COLUMN_OUT_OF_RANGE",
  NAMED_RANGE_NOT_FOUND: "NAMED_RANGE_NOT_FOUND",
  DUPLICATE_NAME: "DUPLICATE_NAME",
  INVALID_PARAMETER: "INVALID_PARAMETER",
  MAX_CELLS_EXCEEDED: "MAX_CELLS_EXCEEDED",
  OUTSIDE_TEMPLATE_RANGE: "OUTSIDE_TEMPLATE_RANGE",
} as const;

export type ErrorCodeType = (typeof ErrorCode)[keyof typeof ErrorCode];

export class EngineError extends Error {
  constructor(
    public readonly code: ErrorCodeType,
    message: string,
  ) {
    super(message);
    this.name = "EngineError";
  }
}

// ---------------------------------------------------------------------------
// Workbook I/O
// ---------------------------------------------------------------------------

export interface XlsxHandle {
  workbook: ExcelJS.Workbook;
  filePath: string;
}

/**
 * XLSX ファイルを開いて ExcelJS Workbook を返す。
 */
/** ファイルサイズ上限 (100 MB) */
const MAX_FILE_SIZE = 100 * 1024 * 1024;

export async function openXlsx(filePath: string): Promise<XlsxHandle> {
  let stat: Awaited<ReturnType<typeof fs.stat>>;
  try {
    stat = await fs.stat(filePath);
  } catch {
    throw new EngineError(ErrorCode.FILE_NOT_FOUND, `File not found: ${filePath}`);
  }

  if (stat.size > MAX_FILE_SIZE) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `File too large (${Math.round(stat.size / 1024 / 1024)} MB). Maximum supported size is 100 MB.`,
    );
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
  } catch (e) {
    throw new EngineError(
      ErrorCode.INVALID_XLSX,
      `Failed to read XLSX file: ${e instanceof Error ? e.message : String(e)}`,
    );
  }

  return { workbook, filePath };
}

/**
 * Workbook をファイルに保存する。
 */
export async function saveXlsx(handle: XlsxHandle): Promise<void> {
  await handle.workbook.xlsx.writeFile(handle.filePath);
}

// ---------------------------------------------------------------------------
// Sheet resolution
// ---------------------------------------------------------------------------

/**
 * シート名（string）または 1-based インデックス（number）でワークシートを解決する。
 */
export function resolveSheet(
  workbook: ExcelJS.Workbook,
  sheet: string | number,
): ExcelJS.Worksheet {
  let ws: ExcelJS.Worksheet | undefined;

  if (typeof sheet === "number") {
    // ExcelJS worksheets are 1-indexed internally; getWorksheet accepts id
    // But we want to support 1-based positional index
    const sheets = workbook.worksheets;
    if (sheet < 1 || sheet > sheets.length) {
      throw new EngineError(
        ErrorCode.SHEET_NOT_FOUND,
        `Sheet index ${sheet} out of range (1-${sheets.length})`,
      );
    }
    ws = sheets[sheet - 1];
  } else {
    ws = workbook.getWorksheet(sheet);
  }

  if (!ws) {
    throw new EngineError(
      ErrorCode.SHEET_NOT_FOUND,
      `Sheet not found: ${sheet}`,
    );
  }

  return ws;
}
