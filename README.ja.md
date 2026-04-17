# xlsx-mcp-server

[![CI](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml)

Excel (.xlsx) ファイルの読み取り・編集を行うローカル [MCP](https://modelcontextprotocol.io/) サーバ。Claude Code、Cursor、その他 MCP 対応クライアントで動作する。

セルデータ、書式設定、数式、シート管理、行列操作、データ入力規則、名前付き範囲、セル結合をカバーする **37 ツール** を提供。すべて stdio 経由でローカル実行され、ファイルのアップロードは不要である。

## 機能一覧

| カテゴリ | ツール |
|---|---|
| **読み取り** | `get_workbook_info`, `read_sheet`, `read_cell`, `search_cells`, `get_sheet_properties`, `list_named_ranges`, `list_data_validations`, `list_images` |
| **書き込み** | `write_cell`, `write_cells`, `write_row`, `write_rows`, `clear_cells`, `create_workbook` |
| **書式** | `format_cells`, `format_cells_bulk` |
| **行・列** | `set_column_width`, `set_column_widths`, `set_row_height`, `set_row_heights`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns` |
| **シート操作** | `add_sheet`, `rename_sheet`, `delete_sheet`, `copy_sheet` |
| **表示設定** | `set_freeze_panes`, `set_auto_filter`, `remove_auto_filter` |
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

`read_cell` は数式とキャッシュされた計算結果の両方を返す。

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

**`read_sheet`** — シートのセルデータを読み取る。範囲指定可能。
```
file_path, sheet, range?
```

**`read_cell`** — 単一セルの値、数式、型、書式情報。
```
file_path, sheet, cell
```

**`search_cells`** — セル全体からテキストまたは数値を検索する。
```
file_path, query, sheet?, case_sensitive?
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

**`write_cell`** — セルの値または数式を設定する。`=` で始めると数式になる。
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

**`clear_cells`** — 範囲内のセル値をクリアする（書式は保持）。
```
file_path, sheet, range
```

**`create_workbook`** — 新しい空の .xlsx ワークブックを作成する。
```
file_path, sheet_name?
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

**`insert_rows`** — 指定位置に空の行を挿入する。
```
file_path, sheet, row, count
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

### 表示設定

**`set_freeze_panes`** — 行・列のウィンドウ枠を固定する。0 を指定すると解除。
```
file_path, sheet, row, column
```

**`set_auto_filter`** — 範囲にオートフィルタを有効にする。
```
file_path, sheet, range
```

**`remove_auto_filter`** — シートからオートフィルタを解除する。
```
file_path, sheet
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

### 非対応機能（Python/openpyxl/xlwings で代替すること）

| 機能 | 詳細 |
|------|------|
| **数式の再計算** | キャッシュされた計算結果は読み取れるが、値を変更しても数式は再計算されない。再計算には Excel で開く必要がある。 |
| **グラフ** | グラフの読み取り・作成・編集はできない。保存時に既存のグラフは保持される。 |
| **ピボットテーブル** | ピボットテーブルの読み取り・作成はできない |
| **条件付き書式** | 条件付き書式ルールの読み取り・作成はできない |
| **VBA/マクロ** | マクロ有効ブック (.xlsm) はサポートされていない |
| **数式参照の自動更新** | 行・列の挿入/削除時に既存数式のセル参照は自動シフトされない（例: `=SUM(A1:A10)` は行挿入後もそのまま） |

### その他の制限

- **copy_sheet は部分的** — セル値、スタイル、列幅、行高、結合セルをコピーする。データ入力規則、条件付き書式、表示設定はコピーされない
- **範囲サイズ制限** — 書き込み・書式・データ検証ツールは 100,000 セルを超える範囲を拒否する
- **ファイルサイズ制限** — 100 MB を超えるファイルは開けない

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
