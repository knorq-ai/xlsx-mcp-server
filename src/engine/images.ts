/**
 * Image listing from worksheets.
 */

import ExcelJS from "exceljs";

/** Excel のデフォルト列幅から算出した 1 列あたりの概算ピクセル数 */
const APPROX_PIXELS_PER_COLUMN = 64;
/** Excel のデフォルト行高から算出した 1 行あたりの概算ピクセル数 */
const APPROX_PIXELS_PER_ROW = 20;

export interface ImageInfo {
  imageId: string;
  name: string;
  extension: string;
  width: number;
  height: number;
}

/**
 * ワークシート内の画像一覧を返す。
 */
export function listSheetImages(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
): ImageInfo[] {
  const images: ImageInfo[] = [];
  const wsImages = ws.getImages();

  for (const img of wsImages) {
    let extension = "unknown";
    try {
      const imageData = workbook.getImage(Number(img.imageId));
      if (imageData) {
        extension = imageData.extension ?? "unknown";
      }
    } catch {
      // 画像データが見つからない場合は extension を "unknown" のままにする
    }

    const range = img.range;

    // デフォルト列幅・行高に基づく概算ピクセル寸法
    let width = 0;
    let height = 0;
    if (range && "tl" in range && "br" in range) {
      const tl = range.tl as { col: number; row: number };
      const br = range.br as { col: number; row: number };
      width = Math.max(0, Math.round((br.col - tl.col) * APPROX_PIXELS_PER_COLUMN));
      height = Math.max(0, Math.round((br.row - tl.row) * APPROX_PIXELS_PER_ROW));
    }

    images.push({
      imageId: img.imageId,
      name: `Image ${img.imageId}`,
      extension,
      width,
      height,
    });
  }

  return images;
}
