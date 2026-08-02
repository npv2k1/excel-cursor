# Excel Cursor — Release Readiness Plan

Target: release a backward-compatible `1.1.0` after all release gates pass.

## 0. Protect and reconcile the working tree

- [x] Create `codex/release-readiness-1.1.0` from `origin/main` (`v1.0.3`).
- [x] Preserve the current uncommitted changes in `API.md`, `README.md`, `example/index.ts`, and `src/core/ExcelCursor.ts`; do not stage `excel-builder.skill` unless explicitly requested.
- [x] Reconcile the constructor migration with upstream tests using overloads that support both `new ExcelCursor(options)` and `new ExcelCursor(workbook, options)`.
- [x] Record a clean baseline by running build, test, lint, CJS import, ESM import, and package dry-run.

## 1. Repair build, tests, and package contract (release blockers)

- [x] Fix Jest configuration so JSONC `tsconfig.json` is not loaded through Node's JSON parser.
- [x] Add the missing ESLint plugin; split `lint` and `lint:fix`; correct `test/` globs to `tests/`.
- [x] Repair dual-package output so both native ESM and CJS consumers work, including Node ESM resolution and declaration paths.
- [x] Add package consumer fixtures that install the packed tarball and verify `require()` and `import` rather than importing `dist` directly.
- [x] Add `engines`, `sideEffects`, repository/bugs/homepage, publish metadata, supported Node versions, and remove redundant build artifacts such as `.tsbuildinfo` from the package.

## 2. Fix Excel correctness and API lifecycle

- [x] Centralize address/position validation: anchored syntax, finite integers, Excel row/column bounds, and ordered ranges. (Stable typed error codes remain future work.)
- [x] Make constructor sheet selection idempotent for existing workbooks and initialize per-worksheet extent state from actual worksheet data.
- [x] Correct formula handling: normalize a leading `=`, update tracking, and add round-trip coverage.
- [x] Make `copyRange` overlap-safe by snapshotting source values and styles; document copied metadata.
- [x] Stop swallowing `rowSpan` errors. (Stable typed error codes remain future work.)
- [x] Correct extent tracking after formula, insert/delete, sheet switches, sparse rows, and copy operations.
- [ ] Define consistent behavior for `isBorderAll`, `setColWidth`, and naming compatibility for the misspelled `goBackToFirstCollumn` API.

## 3. Separate streaming from random-access behavior

- [x] Introduce an append-only `StreamingExcelWriter` capability with `addRow`, `addRows`/async iteration, and `commit`.
- [ ] Keep existing stream construction as a compatibility adapter, but reject unsupported random-access operations with typed errors.
- [x] Remove/redirect `saveWorkbook()` for streaming workbooks; use the constructor output target plus `commit()`.
- [x] Do not return mutable committed rows from the new streaming API.
- [ ] Add lifecycle cleanup for auto-created temporary output files.

## 4. Security and resource controls

- [x] Add configurable `maxCells`, `maxRows`, and `maxCols` limits for range and batch operations.
- [x] Add sequential `AsyncIterable` ingestion, progress callbacks, and `AbortSignal` support.
- [x] Add safe text APIs and document an explicit trusted-formula/external-link policy.
- [ ] Document file-path trust boundaries; provide atomic output and optional confined output-root helpers for service usage. (Boundaries documented; helpers remain future work.)
- [x] Upgrade/override vulnerable transitive dependencies and make `pnpm audit --prod` a release gate with documented exceptions only.
- [ ] Generate SBOM and run dependency-license checks.

## 5. Test matrix and quality gates

- [x] Unit/regression tests cover public cursor behavior, invalid boundaries, overlap cases, and sheet state.
- [ ] XLSX round-trip tests for values, dates, formulas, styles, comments, merges, conditional formatting, and sparse data. (Core package/formula round trips exist; full matrix remains.)
- [x] Streaming tests for sequential rows, limits, abort, progress, idempotent commit, and post-commit behavior.
- [ ] Property-based tests for address conversion and range validation.
- [ ] Performance fixtures with explicit time/RSS budgets; prevent event-loop/OOM regressions.
- [x] Enforce coverage thresholds in CI.
- [x] Test supported Node runtimes: GitHub Actions passed Node 22/24; packed artifact passed in CI on Node 22 and locally on Node 24. (macOS runner remains optional follow-up.)

## 6. CI/CD and supply-chain hardening

- [x] Replace the current workflow with PR and push gates: frozen install, lint, tests/coverage, build, packed-consumer smoke, audit, and package dry-run.
- [x] Use a Node 22/24 matrix and pnpm version aligned with package metadata and local tooling.
- [x] Pin third-party GitHub Actions to reviewed commit SHAs and declare least-privilege workflow permissions.
- [x] Fix release tag parsing and verify tag and package version are identical before publishing.
- [x] Use npm trusted publishing/OIDC with provenance and a protected release environment; remove long-lived token configuration.
- [ ] Upload coverage, SBOM, packed tarball, and test reports as release evidence.
- [x] Remove the Dockerfile because this repository is a library.

## 7. Documentation and OSS readiness

- [x] Align README/API examples with the tested constructor, formula, streaming, and error contracts.
- [x] Publish compatibility, operation limits, security boundaries, and migration notes.
- [ ] Add OSS governance files. (`SECURITY.md` and `CONTRIBUTING.md` added; code of conduct, CODEOWNERS, and templates remain.)
- [x] Update CHANGELOG with verified behavior and no unverified benchmark claims.

## 8. Release acceptance gate

- [x] Clean GitHub runners pass the PR workflow without local-only dependencies (PR #13).
- [x] Packed tarball passes native CJS and ESM consumer tests on supported Node 22/24.
- [x] No unresolved critical/high production vulnerability; audit retains one moderate finding below the release threshold.
- [x] Streaming cancellation is regression-tested; a local 50,000-row smoke completed in 0.21s with +48.3 MiB RSS on Node 24.
- [ ] Generated workbooks reopen successfully in ExcelJS and at least one independent office-suite smoke test.
- [ ] Publish `1.1.0-rc.1`, install it in a clean consumer project, then promote the exact tested artifact to `1.1.0`.

## Review

Core correctness, package, streaming, security-control, CI/CD, and documentation stages are implemented in staged commits. PR #13 passes Node 22, Node 24, packed-package, SonarCloud, and Snyk checks. Remaining items are explicitly unchecked above, especially typed error codes, atomic/confined output helpers, broad office-suite interoperability, SBOM/license evidence, remaining governance templates, and the live npm trusted-publishing release path.
