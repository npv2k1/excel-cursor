# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0-rc.1] - 2026-08-02

### Added

- Append-only `StreamingExcelWriter` with sequential sync/async ingestion, progress reporting, row limits, cancellation, and explicit commit lifecycle.
- `setSafeText()` for neutralizing formula-like untrusted strings and `setTrustedFormula()` for explicit trusted formulas.
- Configurable `maxCells`, `maxRows`, and `maxCols` operation limits.
- Packed-package CommonJS and ESM smoke testing and Node.js 22/24 CI coverage.

### Changed

- Node.js 22 is now the minimum supported runtime.
- Both `new ExcelCursor(options)` and `new ExcelCursor(workbook, options)` constructor forms are supported.
- Cell addresses and ranges are strictly checked against Excel dimensions.
- Formula input is normalized with or without a leading `=`.
- Release publishing uses a verified tarball, npm trusted publishing, and provenance.

### Fixed

- Native ESM package resolution and CommonJS/ESM package metadata.
- Existing `Sheet1` reuse, sparse worksheet extent tracking, and overlap-safe range copies.
- Streaming save lifecycle now rejects unsupported `saveWorkbook(filepath)` calls instead of failing inside ExcelJS.
- Row merge errors are propagated rather than logged and swallowed.

### Security

- Formula-like imported strings can be written through an explicit safe-text API.
- Synchronous range operations have bounded default work and production dependency overrides are audited in CI.

## [1.0.3] - 2026-04-15

### Added

- Comprehensive unit test suite with 101 tests covering:
  - ExcelCursor class (68 tests): constructor, navigation, data operations, formatting, merging, sheet operations, tracking, and more
  - Helper functions (33 tests): column letter/number conversion, address parsing, position conversion
  - Utility functions (8 tests): createWorkbook and createStreamWorkbook

### Changed

- Improved jest configuration for better test reliability
- Updated dependencies to latest versions

## [1.0.2] - 2026-04-15

### Added

- New row manipulation and formatting methods
- Excel helper tool

### Changed

- Improved TypeScript type definitions
- Enhanced documentation with more examples

## [1.0.1] - 2023

### Fixed

- Dockerfile vulnerabilities
- pnpm release CI issues

## [1.0.0] - 2023

### Added

- Initial release with basic cursor operations
- Cell navigation and data manipulation
- Cell formatting and styling
- Merge cells functionality
- Worksheet management
- Row and column operations
