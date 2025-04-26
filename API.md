# Excel Cursor API Documentation

## Table of Contents

- [Installation](#installation)
- [Basic Usage](#basic-usage)
- [API Reference](#api-reference)
  - [Constructor](#constructor)
  - [Navigation Methods](#navigation-methods)
  - [Data Operations](#data-operations)
  - [Formatting](#formatting)
  - [Cell Operations](#cell-operations)
  - [Worksheet Operations](#worksheet-operations)

## Installation

```bash
npm install excel-cursor
# or
yarn add excel-cursor
# or
pnpm add excel-cursor
```

## Basic Usage

```typescript
import { Workbook } from 'exceljs';
import { ExcelCursor } from 'excel-cursor';

const workbook = new Workbook();
const cursor = new ExcelCursor(workbook);

// Basic operations
cursor.move('A1').setData('Hello').nextRow().setData('World');

// Save the workbook
await workbook.xlsx.writeFile('output.xlsx');
```

## API Reference

### Constructor

```typescript
new ExcelCursor(workbook: Workbook, sheetName?: string)
```

Creates a new Excel cursor instance.

- `workbook`: ExcelJS Workbook instance
- `sheetName`: Optional worksheet name (defaults to 'Sheet1')

### Navigation Methods

#### move(address: string): ExcelCursor

Moves the cursor to a specific cell address (e.g., 'A1', 'B2').

#### moveTo(row: number, col: number): ExcelCursor

Moves the cursor to specific row and column coordinates.

#### nextRow(n = 1): ExcelCursor

Moves the cursor down by n rows.

#### prevRow(n = 1): ExcelCursor

Moves the cursor up by n rows.

#### nextCol(n = 1): ExcelCursor

Moves the cursor right by n columns.

#### prevCol(n = 1): ExcelCursor

Moves the cursor left by n columns.

### Data Operations

#### setData(data: any, address?: string): ExcelCursor

Sets data in the current cell or at a specific address.

#### getData(address?: string): any

Gets data from the current cell or a specific address.

### Formatting

#### formatCell(format: CellFormat, address?: string): ExcelCursor

Applies formatting to the current cell or a specific address.

```typescript
interface CellFormat {
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    color?: string;
  };
  alignment?: {
    vertical?: 'top' | 'middle' | 'bottom';
    horizontal?: 'left' | 'center' | 'right';
  };
  fill?: {
    type?: 'pattern';
    pattern?: 'solid';
    fgColor?: string;
  };
  border?: {
    top?: { style?: string; color?: string };
    left?: { style?: string; color?: string };
    bottom?: { style?: string; color?: string };
    right?: { style?: string; color?: string };
  };
}
```

### Cell Operations

#### colSpan(n: number, address?: string): ExcelCursor

Merges n columns from the current position or specified address.

#### rowSpan(n: number, address?: string): ExcelCursor

Merges n rows from the current position or specified address.

#### setColWidth(width: number, colOrAddress?: number | string): ExcelCursor

Sets the width of the current or specified column.

#### setRowHeight(height: number, row?: number): ExcelCursor

Sets the height of the current or specified row.

### Worksheet Operations

#### getWorkbook(): Workbook

Returns the current workbook instance.

#### getCurrentAddress(): string

Returns the current cell address (e.g., 'A1').

#### getCurrentPosition(): CellPosition

Returns the current position as {row: number, col: number}.

## Error Handling

The library throws descriptive errors for invalid operations:

- Invalid cell addresses
- Out of range operations
- Invalid formatting options
- Worksheet operation errors

## Best Practices

1. Chain operations for cleaner code:

```typescript
cursor
  .move('A1')
  .setData('Header')
  .formatCell({ font: { bold: true } })
  .nextRow();
```

2. Use position tracking for dynamic operations:

```typescript
const currentPos = cursor.getCurrentPosition();
cursor.moveTo(currentPos.row + 1, currentPos.col);
```

3. Handle errors appropriately:

```typescript
try {
  cursor.move('InvalidAddress');
} catch (error) {
  console.error('Invalid cell address:', error.message);
}
```
