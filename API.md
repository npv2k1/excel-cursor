# Excel Cursor API

## `ExcelCursor`

An in-memory/random-access cursor over an ExcelJS workbook.

```ts
new ExcelCursor(options?: ExcelCursorOptions)
new ExcelCursor(workbook?: Workbook | stream.xlsx.WorkbookWriter, options?: ExcelCursorOptions)
```

### Options

| Option | Default | Description |
| --- | --- | --- |
| `workbook` | new workbook | Alternative to the first constructor argument. |
| `sheetName` | `Sheet1` | Reuse or create the selected worksheet. |
| `filename` | temporary path | Compatibility streaming output path. |
| `isStream` | `false` | Create an ExcelJS streaming workbook. Prefer `StreamingExcelWriter`. |
| `isBorderAll` | `false` | Add a thin border when tracked data is written. |
| `maxCells` | `100000` | Maximum cells touched by one synchronous range/batch operation. |
| `maxRows` | `1048576` | Maximum rows in one synchronous range/batch operation. |
| `maxCols` | `16384` | Maximum columns in one synchronous range/batch operation. |

Limits must be positive safe integers and cannot exceed Excel's dimensions.

### Navigation and position

- `move(address)`, `moveTo(row, col)`
- `nextRow(n?)`, `prevRow(n?)`, `nextCol(n?)`, `prevCol(n?)`
- `getCurrentAddress()`, `getCurrentPosition()`
- `moveLastRow()`, `moveLastCol()`
- `getLastRow()`, `getLastCol()`, `getLastColAddress()`, `getLastCellAddress()`
- `goBackToFirstCollumn()` — legacy misspelling retained for compatibility

Addresses are case-insensitive but must be within `A1:XFD1048576`. Invalid coordinates throw.

### Values and formulas

- `setData(value, address?)` — assign an ExcelJS cell value; trusted-input API.
- `setSafeText(value, address?)` — store untrusted string input as text and neutralize formula prefixes.
- `getCellValue(address?)` — return the ExcelJS cell value.
- `setTrustedFormula(formula, address?)` — set a trusted formula. Leading `=` is normalized away.
- `setFormula(formula, address?)` — compatibility alias for `setTrustedFormula`.
- `addComment(text, author?, address?)`

Formula APIs do not sanitize external links and must not receive user-controlled expressions.

### Layout and formatting

- `formatCell(style, address?)`
- `applyStyleToRange(style, startAddress, endAddress)`
- `formatCellNumber(address?, format?)`, `borderAll(address?)`, `center(address?)`
- `setColWidth(width, columnOrAddress?)`, `setRowHeight(height, rowOrAddress?)`
- `colSpan(count, address?)`, `rowSpan(count, address?)`
- `addConditionalFormatting(range, type, rules)`
- `createRegion(rows, cols)`

Range operations reject reversed ranges and operations exceeding configured limits.

### Rows, ranges, and worksheets

- `insertRow(values?)`, `deleteRow()`, `addRow(values)`, `addRows(rows)`
- `copyRange(sourceStart, sourceEnd, targetStart)` — snapshots values and styles first, so overlapping copies are safe.
- `createSheet(name)`, `switchSheet(name)`, `setWorksheet(worksheet)`
- `getWorkbook()`

`copyRange` copies cell values and styles only. It does not promise to copy comments, merges, row/column dimensions, or conditional formatting.

### Persistence

- `saveWorkbook(filepath)` writes an in-memory workbook.
- `commit()` commits a compatibility streaming workbook; it is a no-op for in-memory workbooks.

`saveWorkbook()` throws in streaming mode because its destination was fixed when the writer was created.

## `StreamingExcelWriter`

Append-only writer intended for large exports.

```ts
new StreamingExcelWriter({
  filename: string,
  sheetName?: string,
  maxRows?: number,
  signal?: AbortSignal,
  onProgress?: ({ rowsWritten, maxRows }) => void,
  useSharedStrings?: boolean,
  useStyles?: boolean,
})
```

- `addRow(values): this` immediately commits one row.
- `addRows(iterable | asyncIterable): Promise<number>` processes rows sequentially and returns the total written.
- `getRowsWritten(): number`
- `commit(): Promise<void>` finalizes the workbook; repeated calls return the same promise.

Adding rows after commit starts, exceeding `maxRows`, aborting, or a failed underlying write throws. Cancellation is cooperative between rows and does not remove an already-created partial file.

## Error handling and security

The current public contract uses standard `Error` and `TypeError`, with descriptive messages. Do not branch business logic on message text. Wrap library calls at the application boundary and treat any rejected write/commit as a failed export.

Paths, formulas, hyperlinks, and raw ExcelJS values are trust boundaries. See [SECURITY.md](SECURITY.md) and the README before processing untrusted workloads.
