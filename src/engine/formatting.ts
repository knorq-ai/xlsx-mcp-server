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
  // ExcelJS はファイル読み込み時に同一書式のセル間で style オブジェクトを
  // 共有する。そのまま部分更新すると無関係なセルの書式まで変わるため、
  // まず自前のシャローコピーに差し替えて共有を断つ
  // （font/fill/border/alignment は以下で常に新しいオブジェクトを代入する）。
  cell.style = { ...cell.style };

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
      // fillPattern: "solid" のみ指定のときは既存の塗り色を保持する（白で潰さない）
      const existingFg = (cell.fill as ExcelJS.FillPattern | undefined)?.fgColor;
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: opts.fillColor !== undefined
          ? { argb: `FF${opts.fillColor}` }
          : existingFg ?? { argb: "FFFFFFFF" },
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

// ---------------------------------------------------------------------------
// Style read-back
// ---------------------------------------------------------------------------

const BORDER_STYLES = ["thin", "medium", "thick", "double", "dotted", "dashed"] as const;
const H_ALIGN = ["left", "center", "right", "justify"] as const;
const V_ALIGN = ["top", "middle", "bottom"] as const;

function argbToHex(color: Partial<ExcelJS.Color> | undefined): string | undefined {
  const argb = color?.argb;
  if (!argb) return undefined;
  return argb.length === 8 ? argb.slice(2) : argb;
}

/**
 * セルの書式を CellFormatOptions（format_cells が受け取る形式）に要約する。
 * 読み取った書式をそのまま format_cells に渡して複製できるようにするための
 * 読み書き対称な表現。書式が何もなければ undefined。
 */
export function summarizeCellStyle(cell: ExcelJS.Cell): CellFormatOptions | undefined {
  const out: CellFormatOptions = {};

  const f = cell.font;
  if (f) {
    if (f.bold) out.bold = true;
    if (f.italic) out.italic = true;
    if (f.underline) out.underline = true;
    if (f.strike) out.strikethrough = true;
    if (f.name) out.fontName = f.name;
    if (f.size) out.fontSize = f.size;
    const fc = argbToHex(f.color);
    if (fc) out.fontColor = fc;
  }

  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  if (fill && fill.type === "pattern" && fill.pattern && fill.pattern !== "none") {
    const bg = argbToHex(fill.fgColor);
    if (bg) out.fillColor = bg;
  }

  const b = cell.border;
  if (b) {
    const sides: Array<[keyof ExcelJS.Borders, "borderTop" | "borderBottom" | "borderLeft" | "borderRight"]> = [
      ["top", "borderTop"],
      ["bottom", "borderBottom"],
      ["left", "borderLeft"],
      ["right", "borderRight"],
    ];
    for (const [side, flag] of sides) {
      const def = b[side] as Partial<ExcelJS.Border> | undefined;
      if (def?.style) {
        out[flag] = true;
        if (!out.borderStyle && (BORDER_STYLES as readonly string[]).includes(def.style)) {
          out.borderStyle = def.style as CellFormatOptions["borderStyle"];
        }
        if (!out.borderColor) {
          const bc = argbToHex(def.color as Partial<ExcelJS.Color> | undefined);
          if (bc) out.borderColor = bc;
        }
      }
    }
  }

  const a = cell.alignment;
  if (a) {
    if (a.horizontal && (H_ALIGN as readonly string[]).includes(a.horizontal)) {
      out.horizontalAlignment = a.horizontal as CellFormatOptions["horizontalAlignment"];
    }
    if (a.vertical && (V_ALIGN as readonly string[]).includes(a.vertical)) {
      out.verticalAlignment = a.vertical as CellFormatOptions["verticalAlignment"];
    }
    if (a.wrapText) out.wrapText = true;
    if (typeof a.textRotation === "number" && a.textRotation !== 0) {
      out.textRotation = a.textRotation;
    }
  }

  if (cell.numFmt) out.numFmt = cell.numFmt;

  return Object.keys(out).length > 0 ? out : undefined;
}
