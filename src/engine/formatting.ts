/**
 * Cell formatting — font, fill, border, alignment, number format.
 */

import ExcelJS from "exceljs";

export interface CellFormatOptions {
  // Font
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontName?: string;
  fontSize?: number;
  fontColor?: string; // hex "FF0000"

  // Fill
  fillColor?: string; // hex "FFFF00"
  fillPattern?: "solid" | "none";

  // Border
  borderStyle?: "thin" | "medium" | "thick" | "double" | "dotted" | "dashed";
  borderColor?: string; // hex
  borderTop?: boolean;
  borderBottom?: boolean;
  borderLeft?: boolean;
  borderRight?: boolean;

  // Alignment
  horizontalAlignment?: "left" | "center" | "right" | "justify";
  verticalAlignment?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  textRotation?: number;

  // Number format
  numFmt?: string;
}

export interface CellFormatBulkGroup {
  range: string;
  format: CellFormatOptions;
}

/**
 * セルに書式を適用する。
 * 既存の書式とマージし、指定されたプロパティのみ上書きする。
 */
export function applyCellFormat(
  cell: ExcelJS.Cell,
  opts: CellFormatOptions,
): void {
  // Font
  if (
    opts.bold !== undefined ||
    opts.italic !== undefined ||
    opts.underline !== undefined ||
    opts.strikethrough !== undefined ||
    opts.fontName !== undefined ||
    opts.fontSize !== undefined ||
    opts.fontColor !== undefined
  ) {
    const existing = cell.font ?? {};
    const font: Partial<ExcelJS.Font> = { ...existing };
    if (opts.bold !== undefined) font.bold = opts.bold;
    if (opts.italic !== undefined) font.italic = opts.italic;
    if (opts.underline !== undefined) font.underline = opts.underline;
    if (opts.strikethrough !== undefined) font.strike = opts.strikethrough;
    if (opts.fontName !== undefined) font.name = opts.fontName;
    if (opts.fontSize !== undefined) font.size = opts.fontSize;
    if (opts.fontColor !== undefined) {
      font.color = { argb: `FF${opts.fontColor}` };
    }
    cell.font = font as ExcelJS.Font;
  }

  // Fill
  if (opts.fillColor !== undefined || opts.fillPattern !== undefined) {
    if (opts.fillPattern === "none") {
      cell.fill = { type: "pattern", pattern: "none" };
    } else {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: `FF${opts.fillColor ?? "FFFFFF"}` },
      };
    }
  }

  // Border
  if (
    opts.borderStyle !== undefined ||
    opts.borderTop !== undefined ||
    opts.borderBottom !== undefined ||
    opts.borderLeft !== undefined ||
    opts.borderRight !== undefined
  ) {
    const style = opts.borderStyle ?? "thin";
    const color = opts.borderColor ? { argb: `FF${opts.borderColor}` } : { argb: "FF000000" };
    const borderDef: Partial<ExcelJS.Border> = { style, color };
    const existing = cell.border ?? {};
    const border: Partial<ExcelJS.Borders> = { ...existing };

    // 個別指定がなければ全辺に適用。個別指定がある場合は true の辺のみ
    const hasAnySideSpec = opts.borderTop !== undefined ||
      opts.borderBottom !== undefined ||
      opts.borderLeft !== undefined ||
      opts.borderRight !== undefined;

    if (!hasAnySideSpec || opts.borderTop === true) border.top = borderDef;
    if (!hasAnySideSpec || opts.borderBottom === true) border.bottom = borderDef;
    if (!hasAnySideSpec || opts.borderLeft === true) border.left = borderDef;
    if (!hasAnySideSpec || opts.borderRight === true) border.right = borderDef;

    cell.border = border as ExcelJS.Borders;
  }

  // Alignment
  if (
    opts.horizontalAlignment !== undefined ||
    opts.verticalAlignment !== undefined ||
    opts.wrapText !== undefined ||
    opts.textRotation !== undefined
  ) {
    const existing = cell.alignment ?? {};
    const alignment: Partial<ExcelJS.Alignment> = { ...existing };
    if (opts.horizontalAlignment !== undefined) alignment.horizontal = opts.horizontalAlignment;
    if (opts.verticalAlignment !== undefined) alignment.vertical = opts.verticalAlignment;
    if (opts.wrapText !== undefined) alignment.wrapText = opts.wrapText;
    if (opts.textRotation !== undefined) alignment.textRotation = opts.textRotation;
    cell.alignment = alignment as ExcelJS.Alignment;
  }

  // Number format
  if (opts.numFmt !== undefined) {
    cell.numFmt = opts.numFmt;
  }
}
