# CLAUDE.md — xlsx-mcp-server

ローカル MCP サーバ。XLSX ファイルの読み取り・書き込み・書式設定・シート操作を提供する。

## ファイル構成

```
src/
  index.ts              … MCP サーバ本体（ツール登録・stdio transport）
  xlsx-engine.ts        … バレルモジュール（engine/* を再エクスポート + 公開 API 関数）
  engine/
    xlsx-io.ts          … ファイル I/O（ExcelJS Workbook）、ErrorCode、EngineError
    cells.ts            … セルアドレス解析（A1 記法）、値読み書き、型変換
    formatting.ts       … セル書式（font/fill/border/alignment/numFmt）
    sheets.ts           … シート操作（追加/名前変更/削除/コピー）
    rows-columns.ts     … 行列操作（挿入/削除）
    data-validation.ts  … データ検証ルール
    images.ts           … 画像一覧
    view-settings.ts    … フリーズペイン、オートフィルタ
    named-ranges.ts     … 名前付き範囲
    file-lock.ts        … ファイル単位の Promise チェーン書き込みロック
  __tests__/
    helpers.ts
    xlsx-reading.test.ts
    xlsx-cell-editing.test.ts
    xlsx-formatting.test.ts
    xlsx-sheet-ops.test.ts
    xlsx-rows-columns.test.ts
    xlsx-data-validation.test.ts
    xlsx-view-settings.test.ts
    xlsx-named-ranges.test.ts
    xlsx-bulk-operations.test.ts
    xlsx-edge-cases.test.ts
    file-lock.test.ts
```

### モジュール依存グラフ（非循環）

```
file-lock (独立)

xlsx-io  ←  cells  ←  formatting
   ↑            ↑
   ├── sheets   └── data-validation
   ├── rows-columns
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
4. 編集系ツール（`write_cell`, `write_rows` 等）で変更を行う

## セルアドレス

- **A1 記法**: セルアドレスは Excel 標準の A1 記法（例: `A1`, `BC42`）
- **範囲**: コロン区切り（例: `A1:C10`）
- **シート指定**: 名前（`"Sheet1"`）または 1-based インデックス（`1`）

## デフォルト動作

| パラメータ | デフォルト値 | 備考 |
|---|---|---|
| `case_sensitive` | `false` | 検索時の大文字小文字区別 |
| シート指定省略 | — | `search_cells` のみ全シート検索 |

## パラメータ規約

- **ファイルパス**: すべて絶対パスで指定する
- **列**: 英字（A, B, ..., Z, AA, AB, ...）
- **行**: 1-based 数値
- **列幅**: 文字数単位（Excel の標準列幅と同じ）
- **行高**: ポイント（pt）

## 構造化レスポンス

`get_workbook_info`, `read_sheet`, `read_cell`, `search_cells`, `list_named_ranges`, `list_data_validations`, `list_images`, `get_sheet_properties` はテキストの後に `<json>...</json>` ブロックで構造化データを返す。LLM はテキスト部分で自然言語応答を構成し、プログラムは JSON 部分をパースして利用できる。

## 書き込みロック

書き込み関数は `withFileLock` でラップされており、同一ファイルへの並行書き込みを自動直列化する。読み取り関数はロック不要。

## 入力検証

- **セルアドレス**: Zod regex `/^[A-Za-z]+\d+$/` で A1 記法を検証
- **列文字**: Zod regex `/^[A-Za-z]+$/` で英字のみを検証
- **行番号**: 1 以上の整数（Zod `.int().min(1)`）
- **カウント**: 1 以上の整数（行挿入・削除の count 等）
- **シートインデックス**: 1 以上の整数（0 は不正）
- **列幅**: 0〜255 文字単位
- **行高**: 0〜409 ポイント
- **フォントサイズ**: 1〜409 ポイント
- **色**: 6 文字 hex（Zod regex `/^[0-9A-Fa-f]{6}$/`、例: `FF0000`）
- **範囲サイズ**: 書き込み・書式・データ検証で 100,000 セル上限
- **ファイルサイズ**: 100 MB 上限（`openXlsx` で検証）
- **`create_workbook`**: 既存ファイルがある場合はエラー（上書き防止）

## ExcelJS の制限事項

以下の機能は ExcelJS の制限上、サポートしない:
- チャート（Chart）の作成・編集
- ピボットテーブルの作成・編集
- 条件付き書式の作成・編集
- VBA マクロの読み書き（XLSM は開けるが VBA は保持のみ）
- **数式参照の自動更新**: `insert_rows` / `insert_columns` / `delete_rows` / `delete_columns` を実行しても、既存セルの数式参照は更新されない。行・列の構造変更は数式の書き込み**前**に行うこと

## アンチパターン

- 大量のセルを個別に `write_cell` で設定 → `write_cells` や `write_rows` でまとめて適用する（1 回のファイル I/O で済む）
- 複数範囲に個別に `format_cells` を適用 → `format_cells_bulk` でまとめて適用する
- 列幅を個別に `set_column_width` で設定 → `set_column_widths` でまとめて設定する
- 行高を個別に `set_row_height` で設定 → `set_row_heights` でまとめて設定する
- 数式を書いた後に `insert_rows` / `insert_columns` で行列を挿入 → 数式参照がずれる。構造変更を先に行い、数式は最後に書く
