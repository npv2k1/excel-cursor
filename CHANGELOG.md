# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
