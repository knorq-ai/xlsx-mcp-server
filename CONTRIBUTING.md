# Contributing to xlsx-mcp-server

## Development setup

```bash
git clone https://github.com/knorq-ai/xlsx-mcp-server.git
cd xlsx-mcp-server
npm install
npm run build
```

## Running tests

```bash
npm test
```

Tests use [Vitest](https://vitest.dev/) and run entirely in-process — no Excel installation required. Each test creates a real `.xlsx` file in `/tmp`, runs the engine function against it, and deletes it on teardown.

## Project structure

```
src/
  index.ts            # MCP server — tool registrations, schema validation
  xlsx-engine.ts      # Barrel module — re-exports engine/* and public API functions
  engine/
    xlsx-io.ts        # File I/O, ExcelJS wrapper, error types, sheet resolution
    cells.ts          # Cell addressing, data access, search
    formatting.ts     # Cell formatting options
    sheets.ts         # Sheet add/rename/delete/copy
    rows-columns.ts   # Row/column insert/delete, width/height
    data-validation.ts # Data validation rules
    named-ranges.ts   # Named range management
    images.ts         # Image listing/metadata
    view-settings.ts  # Freeze panes, auto filter
    file-lock.ts      # Promise-based file write locking
  __tests__/
    helpers.ts        # Test utilities (tmp file management, fixture builders)
    xlsx-reading.test.ts
    xlsx-cell-editing.test.ts
    xlsx-formatting.test.ts
    xlsx-rows-columns.test.ts
    xlsx-sheet-ops.test.ts
    xlsx-edge-cases.test.ts
    xlsx-data-validation.test.ts
    xlsx-named-ranges.test.ts
    xlsx-bulk-operations.test.ts
    xlsx-view-settings.test.ts
```

## Architecture

Every public function in `xlsx-engine.ts` is **stateless**:

1. Open the `.xlsx` workbook from disk using ExcelJS
2. Resolve the target sheet (by name or 1-based index)
3. Perform the operation
4. Save back to disk

Write operations are wrapped in `withFileLock` to serialize concurrent writes to the same file.

## Error handling

All engine errors are thrown as `EngineError` with a machine-readable `code`:

| Code | When |
|------|------|
| `FILE_NOT_FOUND` | Path doesn't exist |
| `INVALID_XLSX` | Not a valid XLSX file |
| `SHEET_NOT_FOUND` | Sheet name or index doesn't exist |
| `CELL_OUT_OF_RANGE` | Invalid cell address |
| `INVALID_RANGE` | Malformed range string, or range exceeds 100K cells |
| `ROW_OUT_OF_RANGE` | Row number out of bounds (1–1,048,576) |
| `COLUMN_OUT_OF_RANGE` | Column number out of bounds (1–16,384) |
| `NAMED_RANGE_NOT_FOUND` | Named range doesn't exist |
| `DUPLICATE_NAME` | Sheet or named range name already exists |
| `INVALID_PARAMETER` | Invalid parameter value, or file exceeds 100 MB |

## Adding a new tool

1. Implement the function in the appropriate `engine/*.ts` module
2. Export it from `xlsx-engine.ts`
3. Register it in `index.ts` with a Zod schema
4. Write tests in the appropriate `__tests__/*.test.ts` file
5. Update the tool count and reference in `README.md`

## Pull requests

- Keep PRs focused — one feature or fix per PR
- All tests must pass (`npm test`)
- Build must succeed (`npm run build`)
- Update `README.md` if the tool interface changes
