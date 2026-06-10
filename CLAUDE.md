# CLAUDE.md — xlsx-mcp-server

ローカル MCP サーバ。XLSX ファイルの読み取り・書き込み・書式設定・シート操作を提供する。

## ファイル構成

```
src/
  index.ts              … MCP サーバ本体（registerTool + annotations、stdio transport）
  xlsx-engine.ts        … バレルモジュール（engine/* を再エクスポート + 公開 API 関数）
  engine/
    xlsx-io.ts          … ファイル I/O（アトミック保存、.xlsm 書き込み拒否）、ErrorCode、EngineError
    cells.ts            … A1 記法解析、値読み書き、型変換、共有数式の実体化、SheetJson 変換
    formatting.ts       … セル書式（適用 + 読み戻し summarizeCellStyle）
    sheets.ts           … シート操作（追加/名前変更/削除/コピー、シート名検証）
    rows-columns.ts     … 行列操作（挿入/削除、結合・データ検証のシフト保全）
    data-validation.ts  … データ検証ルール
    images.ts           … 画像一覧
    view-settings.ts    … フリーズペイン、オートフィルタ
    named-ranges.ts     … 名前付き範囲
    file-lock.ts        … 2 層書き込みロック（プロセス内 Promise チェーン + .mcplock）
  __tests__/            … vitest テスト
```

### モジュール依存グラフ（非循環）

```
xlsx-io ← formatting ← cells
   ↑                     ↑
   ├── file-lock         └── data-validation
   ├── sheets
   ├── rows-columns ← cells
   ├── images
   ├── view-settings
   └── named-ranges
```

## ビルド・テスト

```bash
npm run build     # TypeScript → dist/
npx vitest run    # 全テスト実行
```

## ツール使用ワークフロー（推奨）

1. `get_workbook_info` でワークブックの構造を把握する
2. `read_sheet` で対象シートのデータを読む（range で範囲指定可能）
3. `search_cells` で編集対象のセルを特定する
4. 編集系ツール（`write_cell`, `write_rows`, `copy_range`, `find_replace` 等）で変更を行う

## セルアドレス

- **A1 記法**: セルアドレスは Excel 標準の A1 記法（例: `A1`, `BC42`）
- **範囲**: コロン区切り（例: `A1:C10`）。`A:A` のような全列・全行指定は不可
- **シート指定**: 名前（`"Sheet1"`）または 1-based インデックス（`1`）
- **上限**: 行 1,048,576・列 XFD (16,384)。超過は明示エラー

## 書き込み値

`write_cell` / `write_cells` / `write_row` / `write_rows` の値:

- 文字列・数値・真偽値・null
- `=` で始まる文字列 → 数式（リテラルにするには `'=` でエスケープ）
- `{ date: "2024-01-15" }` → Excel の日付値
- `{ hyperlink: "https://...", text?: "表示名" }` → ハイパーリンク
- 結合セルの**子**への書き込みはエラー（マスターを上書きしてしまうため）

## 構造化レスポンス

読み取り系ツールはテキストサマリの後に `<json>...</json>` ブロックで構造化データを返す。

### read_sheet の JSON 形式（マップ形式）

セルアドレスをキーにした密な形式。**キーが無い = 空セル**。

```json
{
  "sheetName": "Sheet1", "range": "A1:C10", "totalRows": 10, "totalColumns": 3,
  "cells":     { "A1": "Name", "B1": 42, "C1": true },
  "formulas":  { "C2": { "f": "A1*2", "v": 84 } },
  "dates":     { "A3": "2024-01-15T00:00:00.000Z" },
  "errors":    { "B4": "#DIV/0!" },
  "hyperlinks":{ "A5": "https://example.com" },
  "numFmts":   { "B1": "#,##0" },
  "notes":     { "A1": "コメント" },
  "styles":    { "A1": { "bold": true } },
  "mergedCells": ["A1:C1"],
  "truncated": true, "truncatedAtRow": 500
}
```

- `formulas[addr].v` が無い = 結果未計算（書き込み直後の数式。Excel で開くと再計算される）
- `styles` は `include_styles: true` のときのみ。**format_cells が受け取る形式と同一**なので、読んだ書式をそのまま複製できる
- 出力は 5,000 セルで打ち切り（`truncated` フラグ + 続きは range 指定）
- `read_cell` の `style` も同じ format_cells 形式

### 数式の表記

- **読み取り出力**: 先頭の `=` を**付けない**（例: `$C2*E2`）
- **書き込み入力**: 先頭に `=` を**付ける**（例: `=SUM(A1:A2)`）
- **共有数式**: スレーブセルは相対参照をずらした**そのセル固有の数式**で返す。マスターのアドレスを formula に返すことはない（解決不能な場合は `sharedGroupMaster` に分離）
- 構造変更（splice）や共有数式マスターの上書き時は、グループを通常数式に自動実体化する

## 書き込みロック・保存

- 書き込みは 2 層ロック: プロセス内 Promise チェーン + プロセス間 `.mcplock`（PID 記録、死活判定付き、10 秒タイムアウト）
- 保存はアトミック（同一ディレクトリの一時ファイル → rename）。クラッシュで元ファイルは壊れない
- 保存時に `fullCalcOnLoad` を立てる（Excel が開いたとき数式を再計算する）
- `XLSX_BACKUP_ON_WRITE=1` で保存前に `<file>.bak` を作成
- `.xlsm` / `.xltm` への書き込みは拒否（VBA が消えるため）。読み取りは可

## 入力検証

- **セルアドレス**: A1 記法 + Excel グリッド上限（行 1,048,576 / 列 16,384）
- **シート名**: 空・32 文字以上・`* ? : \ / [ ]`・先頭末尾アポストロフィを拒否
- **範囲サイズ**: 書き込み・書式・データ検証・copy_range・sort_range で 100,000 セル上限
- **読み取り出力**: read_sheet 5,000 セル / search_cells max_results（既定 100、最大 1,000）
- **ファイルサイズ**: 100 MB 上限
- **create_workbook**: O_EXCL（`wx`）で OS レベルの上書き防止

## 安全ガード（env 設定、LLM からは変更不可）

- `XLSX_MAX_CELLS_PER_CALL` … 1 コールのセル数上限
- `XLSX_TEMPLATE_MODE` + `XLSX_TEMPLATE_RANGES` … "Sheet!Range" ホワイトリスト外への書き込み拒否。構造変更ツール（行列挿入/削除、シート削除/改名）と find_replace はテンプレートモード中は全面拒否
- `XLSX_BACKUP_ON_WRITE` … 保存前バックアップ

## ExcelJS の制限事項

- **チャート・ピボットテーブル・スライサー**: 書き込み操作で**消える**（ExcelJS が保持しない）。これらを含むブックの編集は不可逆
- **条件付き書式**: 保存で保持はされるが、読み書きツールは未提供
- **VBA**: .xlsm は読み取り専用
- **数式の再計算**: サーバ側では行わない（fullCalcOnLoad で Excel 起動時に再計算）
- **数式参照の自動更新**: `insert_rows` 等で既存数式内の参照はシフトしない。構造変更は数式の書き込み**前**に行うこと
  - 結合セル・データ検証は splice 時に自動シフトされる（rows-columns.ts で保全）

## アンチパターン

- 大量のセルを個別に `write_cell` → `write_cells` / `write_rows` でまとめる
- 個別の `format_cells` 連打 → `format_cells_bulk`
- 検索 → 1 件ずつ write_cell で置換 → `find_replace`
- 書式付きブロックの再現を 1 セルずつ → `copy_range`
- 数式を書いた後に行列を挿入 → 参照がずれる。構造変更が先
