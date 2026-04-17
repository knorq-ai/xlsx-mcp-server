# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] — 2026-04-17

### Changed (BREAKING)
- Renamed package from `xlsx-mcp-server` to `@knorq/xlsx-mcp-server`. Update your `.mcp.json` / install commands to the scoped name.
- Pinned `engines.node` to `>=18.0.0`.

### Added
- `XLSX_MAX_CELLS_PER_CALL` environment variable: bounds bulk write/format operations. Defaults to the existing 100,000-cell range cap; deployments can lower it. Reject is enforced server-side, before the LLM-supplied range reaches the engine.
- Template mode (`XLSX_TEMPLATE_MODE=1` + `XLSX_TEMPLATE_RANGES=Sheet1!A1:D10,Sheet1!F2:F100`): when enabled, all writes/formats/clears must fall inside one of the declared ranges. Writes outside are rejected with `OUTSIDE_TEMPLATE_RANGE`. Enforces structural integrity at the server, not in the prompt.
- GitHub Actions workflow that publishes to npm with `--provenance --access public` on tag push, signed via OIDC.

## [1.1.0] and earlier

See git history.
