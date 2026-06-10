# xlsx-mcp-server

[![CI](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml)

Excel (.xlsx) ファイルの読み取り・編集を行うローカル [MCP](https://modelcontextprotocol.io/) サーバ。Claude Code、Cursor、その他 MCP 対応クライアントで動作する。

セルデータ、書式設定、数式、範囲コピー・ソート・置換、シート管理、行列操作、データ入力規則、名前付き範囲、セル結合、セルノート、シート保護、ページ設定をカバーする **47 ツール** を提供。すべて stdio 経由でローカル実行され、ファイルのアップロードは不要である。

## 機能一覧

| カテゴリ | ツール |
|---|---|
| **読み取り** | `get_workbook_info`, `read_sheet`, `read_cell`, `search_cells`, `get_sheet_properties`, `list_named_ranges`, `list_data_validations`, `list_images` |
| **書き込み** | `write_cell`, `write_cells`, `write_row`, `write_rows`, `clear_cells`, `set_cell_note`, `create_workbook` |
| **範囲操作** | `copy_range`, `find_replace`, `sort_range` |
| **書式** | `format_cells`, `format_cells_bulk` |
| **行・列** | `set_column_width`, `set_column_widths`, `set_row_height`, `set_row_heights`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `set_row_visibility`, `set_column_visibility` |
| **シート操作** | `add_sheet`, `rename_sheet`, `delete_sheet`, `copy_sheet`, `set_sheet_properties`, `protect_sheet`, `unprotect_sheet` |
| **表示・レイアウト** | `set_freeze_panes`, `set_auto_filter`, `remove_auto_filter`, `set_page_setup` |
| **入力規則** | `add_data_validation`, `remove_data_validation` |
| **構造** | `add_named_range`, `delete_named_range`, `merge_cells`, `unmerge_cells` |

### 一括操作

書き込み・書式・行列ツールには一括バリアント（`write_cells`, `write_rows`, `format_cells_bulk`, `set_column_widths`, `set_row_heights`）がある。これらは 1 回のファイル読み書きサイクルで複数のターゲットを処理する。単一ターゲット版をループで呼ぶ代わりにこれらを使用すること。

### 数式サポート

値を `=` で始めると数式として書き込まれる:

```
write_cell  →  value: "=SUM(A1:A10)"
write_cells →  cells: [{cell: "B1", value: "=A1*2"}, {cell: "B2", value: "=VLOOKUP(...)"}]
```

`read_cell` は数式とキャッシュされた計算結果の両方を返す。`=` で始まる文字列をそのまま書き込みたい場合は先頭にシングルクォートを付ける（`'=text` は文字列 `=text` として書き込まれる — Excel のエスケープ規則と同じ）。編集時に数式は再計算されないが、保存のたびに「開いたときに再計算」フラグが有効化されるため、Excel で開けばすべて再計算される。

### 日付・ハイパーリンク値

すべての書き込みツールで、オブジェクト形式の値により真の Excel 日付とハイパーリンクを書き込める:

```
write_cell →  value: {date: "2024-01-15"}                            // 真の Excel 日付セル
write_cell →  value: {hyperlink: "https://example.com", text: "Docs"} // 表示テキスト付きリンク
```

### read_sheet の JSON 形式

`read_sheet` はセルデータをアドレスキーのマップとして `<json>...</json>` ブロックで返す。アドレスが存在しない場合は空セルを意味する:

```json
{
  "sheetName": "Sheet1",
  "range": "A1:C3",
  "cells": {"A1": "Product", "B1": "Price", "A2": "Widget", "B2": 9.99},
  "formulas": {"C2": {"f": "B2*2", "v": 19.98}},
  "dates": {"A3": "2024-01-15T00:00:00.000Z"},
  "mergedCells": ["A1:B1"]
}
```

追加のマップ（`errors`, `hyperlinks`, `numFmts`, `notes`、および `include_styles: true` 指定時の `styles`（`format_cells` と同じ語彙））は該当データがある場合のみ出現する。出力は 5,000 セルで打ち切られ、打ち切り時は `truncated: true` が付くため、大きいシートは `range` で分割して読むこと。

## クイックスタート

### 方法 1: npm からインストール

```bash
npm install -g @knorq/xlsx-mcp-server
```

インストール後、MCP 設定に追加する（下記 [設定](#設定) を参照）。

### 方法 2: npx を使用（インストール不要）

設定を追加するだけで `npx` が自動的にダウンロード・実行する:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### 方法 3: ソースからビルド

```bash
git clone https://github.com/knorq-ai/xlsx-mcp-server.git
cd xlsx-mcp-server
npm install
npm run build
npm link        # `xlsx-mcp-server` をグローバルで利用可能にする
```

## 設定

### Claude Code

プロジェクトの `.mcp.json`（プロジェクト単位）または `~/.claude/settings.json`（グローバル）に追加:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### Cursor

Cursor 設定の MCP サーバ構成に追加:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### ローカルビルドを使用する場合（npm 不要）

ソースからビルドして `npm link` を実行済みの場合:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "xlsx-mcp-server"
    }
  }
}
```

または、ビルド済みファイルを直接参照:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "node",
      "args": ["/absolute/path/to/xlsx-mcp-server/dist/index.js"]
    }
  }
}
```

## 配布方法

### npm 経由（推奨）

```bash
npm publish
```

受け取り側は以下でインストール:

```bash
npm install -g @knorq/xlsx-mcp-server
```

インストールを省略することも可能 — 上記の `npx` 設定を含む `.mcp.json` を共有するだけで動作する。

### zip / git 経由

リポジトリを共有し、受け取り側が以下を実行:

```bash
git clone https://github.com/knorq-ai/xlsx-mcp-server.git
cd xlsx-mcp-server
npm install
npm run build
npm link
```

その後、上記の設定を追加する。

## ツールリファレンス

### 読み取り

**`get_workbook_info`** — シート一覧、名前付き範囲数、ファイルプロパティ。
```
file_path
```

**`read_sheet`** — シートのセルデータをアドレスキーの JSON マップとして読み取る（[read_sheet の JSON 形式](#read_sheet-の-json-形式) を参照）。出力は 5,000 セル上限。
```
file_path, sheet, range?, include_styles?
```

**`read_cell`** — 単一セルの値、数式、型、書式情報。書式は `format_cells` が受け付けるのと同じ語彙（`bold`, `fillColor` 等）で返されるため、読み取った書式をそのまま書き戻せる。
```
file_path, sheet, cell
```

**`search_cells`** — セル全体からテキストまたは数値を検索する。
```
file_path, query, sheet?, case_sensitive?, max_results?
```

**`get_sheet_properties`** — シートの状態、サイズ、ウィンドウ枠固定、オートフィルタ、タブ色。
```
file_path, sheet
```

**`list_named_ranges`** — すべての名前付き範囲とその参照先を一覧表示。
```
file_path
```

**`list_data_validations`** — シート上のデータ入力規則を一覧表示。
```
file_path, sheet
```

**`list_images`** — 埋め込み画像のファイル名、拡張子、サイズを一覧表示。
```
file_path, sheet
```

### セル書き込み

**`write_cell`** — セルの値または数式を設定する。`=` で始めると数式になる。日付は `{date: "ISO"}`、リンクは `{hyperlink, text}` を使う。
```
file_path, sheet, cell, value
```

**`write_cells`** — 複数セルを一括設定する。
```
file_path, sheet, cells ({cell, value} の配列)
```

**`write_row`** — 指定位置から 1 行分の値を書き込む。
```
file_path, sheet, row, values, start_column?
```

**`write_rows`** — 複数行のデータを一括書き込みする。
```
file_path, sheet, start_row, rows (2 次元配列), start_column?
```

**`clear_cells`** — 範囲内のセル値・書式をクリアする。`mode: "values"`（デフォルト）は書式を保持、`"formats"` は値を保持、`"all"` は両方クリアする。
```
file_path, sheet, range, mode?
```

**`set_cell_note`** — セルノート（コメント）を設定または削除する。`null` を渡すと削除。
```
file_path, sheet, cell, note
```

**`create_workbook`** — 新しい空の .xlsx ワークブックを作成する。
```
file_path, sheet_name?
```

### 範囲操作

**`copy_range`** — 範囲（値・数式・書式・結合）を別の場所へコピーする。別シートへのコピーも可能。数式内の相対参照はコピー先に合わせてシフトされ、`$` 付きの絶対参照は固定のまま。
```
file_path, sheet, source_range, destination, dest_sheet?
```

**`find_replace`** — プレーン文字列セルを対象に一括置換する。数式・数値・リッチテキスト・ハイパーリンクは変更されない。シート未指定時は全シートを対象とする。
```
file_path, query, replacement, sheet?, case_sensitive?, match_entire_cell?
```

**`sort_range`** — 範囲の行をキー列でソートする。値・数式・書式は行と一緒に移動し、数式内の相対参照は再アンカーされる。範囲が結合セルと交差する場合は失敗する。
```
file_path, sheet, range, key_column, ascending?, has_header?
```

### 書式設定

**`format_cells`** — セル範囲に書式を適用: フォント（太字、斜体、下線、取り消し線、フォント名、サイズ、色）、塗りつぶし（色、パターン）、罫線（スタイル、色、辺）、配置（水平、垂直、折り返し、回転）、表示形式。
```
file_path, sheet, range, format
```

**`format_cells_bulk`** — 複数範囲に異なる書式を一括適用する。1 回のファイル読み書きサイクルで処理。
```
file_path, sheet, groups ({range, format} の配列)
```

### 行・列

**`set_column_width`** — 列の幅を設定する（文字数単位）。
```
file_path, sheet, column, width
```

**`set_column_widths`** — 複数列の幅を一括設定する。
```
file_path, sheet, columns ({column, width} の配列)
```

**`set_row_height`** — 行の高さを設定する（ポイント単位）。
```
file_path, sheet, row, height
```

**`set_row_heights`** — 複数行の高さを一括設定する。
```
file_path, sheet, rows ({row, height} の配列)
```

**`insert_rows`** — 指定位置に空の行を挿入する。`inherit_style: true` で直上の行から書式（と行高）を引き継ぐ。
```
file_path, sheet, row, count, inherit_style?
```

**`delete_rows`** — 指定位置の行を削除する。
```
file_path, sheet, row, count
```

**`insert_columns`** — 指定位置に空の列を挿入する。
```
file_path, sheet, column, count
```

**`delete_columns`** — 指定位置の列を削除する。
```
file_path, sheet, column, count
```

**`set_row_visibility`** — 行範囲を非表示または再表示する。
```
file_path, sheet, start_row, end_row, hidden
```

**`set_column_visibility`** — 列範囲を非表示または再表示する。
```
file_path, sheet, start_column, end_column, hidden
```

### シート操作

**`add_sheet`** — 新しい空のシートを追加する。
```
file_path, name
```

**`rename_sheet`** — 既存のシートの名前を変更する。
```
file_path, sheet, new_name
```

**`delete_sheet`** — ワークブックからシートを削除する。
```
file_path, sheet
```

**`copy_sheet`** — ワークブック内でシートをコピーする。
```
file_path, source_sheet, new_name
```

**`set_sheet_properties`** — シートの表示状態（`visible` / `hidden` / `veryHidden`）とタブ色を設定する。ワークブックには最低 1 枚の表示シートが必要。
```
file_path, sheet, state?, tab_color?
```

**`protect_sheet`** — Excel 上での編集からシートを保護する（パスワード指定可）。これは Excel UI レベルの保護であり暗号化ではない — 本サーバや他のツールからはファイルを変更できる。
```
file_path, sheet, password?
```

**`unprotect_sheet`** — シート保護を解除する。
```
file_path, sheet
```

### 表示・レイアウト

**`set_freeze_panes`** — 行・列のウィンドウ枠を固定する。両方に 0 を指定すると解除。
```
file_path, sheet, freeze_rows, freeze_columns
```

**`set_auto_filter`** — 範囲にオートフィルタを有効にする。
```
file_path, sheet, range
```

**`remove_auto_filter`** — シートからオートフィルタを解除する。
```
file_path, sheet
```

**`set_page_setup`** — 印刷 / PDF レイアウトを設定する: 用紙の向き、印刷範囲、ページに合わせる、用紙サイズ。
```
file_path, sheet, orientation?, print_area?, fit_to_width?, fit_to_height?, paper_size?
```

### データ入力規則

**`add_data_validation`** — 入力規則（リスト、整数、小数、日付、文字列長、カスタム）を追加する。演算子、エラーメッセージ、入力時メッセージの設定が可能。
```
file_path, sheet, range, type, formulae, operator?, allow_blank?, show_error_message?, error_title?, error?, show_input_message?, prompt_title?, prompt?
```

**`remove_data_validation`** — 範囲から入力規則を解除する。
```
file_path, sheet, range
```

### 名前付き範囲

**`add_named_range`** — 名前付き範囲を追加する（ブックスコープまたはシートスコープ）。
```
file_path, name, range, sheet?
```

**`delete_named_range`** — 名前付き範囲を削除する。
```
file_path, name
```

### セル結合

**`merge_cells`** — セル範囲を結合する。
```
file_path, sheet, range
```

**`unmerge_cells`** — 結合済みのセル範囲を解除する。
```
file_path, sheet, range
```

## 既知の制限事項

### 破壊または拒否される機能

| 機能 | 挙動 |
|------|------|
| **グラフ・ピボットテーブル・スライサー** | **書き込み操作で破壊される。** 読み取りは安全だが、ファイルを保存するツールを実行するとワークブックから消える。グラフ・ピボットを残す必要があるワークブックは編集しないこと。 |
| **VBA マクロ (.xlsm/.xltm)** | **読み取り専用。** 保存すると VBA プロジェクトが暗黙に破壊されるため、書き込みは拒否される。 |

### 非対応機能

| 機能 | 詳細 |
|------|------|
| **数式の再計算** | 編集時に数式は評価されない。キャッシュされた計算結果は読み取れるが、キャッシュのない数式セルは `(not calculated)` と読める。保存のたびに「開いたときに再計算」が有効化されるため、Excel で開けば再計算される。 |
| **条件付き書式** | 既存のルールは保存時に保持されるが、読み取り・編集するツールはない。 |
| **数式参照の自動更新** | 行・列の挿入/削除時に数式テキスト内のセル参照は自動シフトされない（例: `=SUM(A1:A10)` は行挿入後もそのまま）。結合セルとデータ入力規則は正しくシフトされる。構造変更は数式の書き込み前に行うこと。 |

### その他の制限

- **copy_sheet は部分的** — セル値、スタイル、列幅、行高、結合セルをコピーする。データ入力規則、条件付き書式、表示設定はコピーされない
- **範囲サイズ制限** — 書き込み・書式・データ検証ツールは 100,000 セルを超える範囲を拒否する（`XLSX_MAX_CELLS_PER_CALL` で下げられる）
- **ファイルサイズ制限** — 100 MB を超えるファイルは開けない

## 安全性と信頼性

- **アトミック保存** — 一時ファイルへの書き込み + リネームで保存するため、保存中のクラッシュでワークブックが破損することはない。
- **プロセス間書き込みロック** — `<file>.mcplock` のアドバイザリロック（孤児ロック検出付き）により、プロセス内の直列化に加えて複数サーバインスタンス間でも書き込みが直列化される。

### 環境変数

サーバ起動時に一度だけ読み込まれるため、ツールパラメータからは上書きできない。

| 変数 | 効果 |
|------|------|
| `XLSX_MAX_CELLS_PER_CALL` | 1 回のツール呼び出しで触れるセル数の上限。デフォルト 100,000。デプロイ側で下げられる。 |
| `XLSX_TEMPLATE_MODE=1` + `XLSX_TEMPLATE_RANGES=Sheet1!A1:D10,Sheet1!F2:F100` | テンプレートモード: すべての書き込み・書式・クリアは宣言された範囲内に収まる必要があり、範囲外は `OUTSIDE_TEMPLATE_RANGE` で拒否される。構造操作（行・列の挿入/削除、シートの削除/名前変更）と `find_replace` もブロックされる。 |
| `XLSX_BACKUP_ON_WRITE=1` | 保存のたびにワークブックを `<file>.bak` へコピーする。 |

## なぜ Raw Python ではなく MCP ツールか？

AI エージェントは Raw Python (openpyxl) でも Excel を操作できるが、MCP ツールの方がトークン効率が大幅に高い:

| 指標 | MCP ツール | Raw Python |
|------|-----------|------------|
| 操作あたりの出力トークン | **60–85% 削減** | ベースライン (エージェントがコード全体を生成) |
| 操作あたりのコスト | **50–80% 削減** | ベースライン |
| 損益分岐点 | **2 操作** | — |
| デバッグ反復 | なし (入力バリデーション済み) | 平均 ~1.5 回/タスク |

削減の主因は **コード生成の省略** である。出力トークンは入力トークンの 5 倍の単価であるため、MCP ツール呼び出し (~30–50 tokens の構造化パラメータ) と、Python コード生成 (~80–200 出力トークン/操作: import、スタイルオブジェクト、イテレーション、保存) の差が大きい。

特に書式設定操作で最大の削減 (~75%) が得られる。openpyxl のスタイル API (`PatternFill`, `Border`, `Side`, `Font`) が冗長なためである。単純なセル読み書きでも ~60% の削減がある。

詳細なシナリオ別分析は [docs/token-efficiency-analysis.md](docs/token-efficiency-analysis.md) を参照。

## 動作要件

- Node.js 18+

## ライセンス

MIT
