# Excel Cursor

TypeScript library for creating and editing `.xlsx` workbooks with a cursor-style API. It supports native CommonJS and ESM consumers, bounded range operations, safe handling of untrusted text, and a separate append-only writer for large exports.

[![npm version](https://badge.fury.io/js/excel-cursor.svg)](https://badge.fury.io/js/excel-cursor)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Requirements

- Node.js 22 or 24
- ExcelJS 4.4

```bash
npm install excel-cursor exceljs
```

## In-memory workbooks

Use `ExcelCursor` when cells must be revisited, formatted, merged, copied, or read.

```ts
import { Workbook } from 'exceljs';
import { ExcelCursor } from 'excel-cursor';

const workbook = new Workbook();
const cursor = new ExcelCursor(workbook, {
  sheetName: 'Report',
  maxCells: 100_000,
  maxRows: 10_000,
  maxCols: 100,
});

cursor.move('A1').setData('Revenue').nextRow().setData(1250);
cursor.move('B2').setTrustedFormula('A2*1.1');
await cursor.saveWorkbook('report.xlsx');
```

Both constructor forms are supported:

```ts
new ExcelCursor({ sheetName: 'Sheet1' });
new ExcelCursor(workbook, { sheetName: 'Sheet1' });
```

If no workbook is supplied, an in-memory workbook is created. Existing sheets are reused rather than duplicated.

## Untrusted input and formulas

`setData()` accepts native ExcelJS values and therefore must only receive values allowed by your application. Use `setSafeText()` for user/imported strings that must never become a spreadsheet formula:

```ts
cursor.setSafeText('=HYPERLINK("https://example.invalid","click")', 'A1');
cursor.setTrustedFormula('SUM(B2:B100)', 'C2');
```

`setSafeText()` prefixes formula-like text (`=`, `+`, `-`, or `@`, including leading control whitespace) with an apostrophe. `setFormula()` remains as a compatibility alias for `setTrustedFormula()`; both accept a leading `=` but store the normalized formula. Never pass untrusted input to either formula method.

## Large, append-only exports

Use `StreamingExcelWriter` when rows can be written once in order. Committed rows cannot be read or modified.

```ts
import { StreamingExcelWriter } from 'excel-cursor';

const controller = new AbortController();
const writer = new StreamingExcelWriter({
  filename: 'large-report.xlsx',
  sheetName: 'Rows',
  maxRows: 500_000,
  signal: controller.signal,
  onProgress: ({ rowsWritten }) => console.log(rowsWritten),
});

await writer.addRows(getRowsAsAsyncIterable());
await writer.commit();
```

The compatibility option `new ExcelCursor({ isStream: true, filename })` still exists, but random-access behavior is not appropriate after rows are committed. Prefer `StreamingExcelWriter` for new streaming code. In stream mode the output path is fixed at construction: call `commit()`, not `saveWorkbook()`.

## Limits and failure behavior

- Excel addresses are validated against `A1:XFD1048576`.
- Synchronous range/batch operations default to `maxCells: 100000`; `maxRows` and `maxCols` default to Excel's limits.
- `StreamingExcelWriter.maxRows` defaults to Excel's row limit.
- Invalid addresses, reversed/oversized ranges, invalid lifecycle transitions, and write failures throw. Errors are not swallowed; callers should catch them at their job/API boundary.
- `AbortSignal` is checked between streamed rows. It cannot undo rows already committed to the output stream.

These controls reduce accidental CPU/memory abuse; they are not tenant isolation. Choose lower limits for externally supplied workloads.

## File-system trust boundary

Output paths are caller-controlled. The library does not confine paths to a directory, prevent symlink traversal, encrypt workbooks, scan macros, or guarantee atomic replacement. A service must resolve and validate paths inside an application-owned output root, use restrictive permissions, write to a temporary file and atomically rename it when appropriate, and avoid exposing arbitrary paths to users.

Formula and hyperlink content may cause spreadsheet applications to access external resources. Only trusted application code should create them.

## Documentation

- [API reference](API.md)
- [Changelog](CHANGELOG.md)
- [Security policy](SECURITY.md)
- [Contributing](CONTRIBUTING.md)

## License

[MIT](LICENSE)
