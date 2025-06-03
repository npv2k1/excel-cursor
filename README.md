# Excel Cursor

A powerful TypeScript library for easy Excel file manipulation using a cursor-based API. Built on top of ExcelJS, it provides an intuitive interface for navigating and modifying Excel workbooks.

[![npm version](https://badge.fury.io/js/excel-cursor.svg)](https://badge.fury.io/js/excel-cursor)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=npv2k1_excel-cursor&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=npv2k1_excel-cursor)

## Features

- 🚀 Intuitive cursor-based navigation in Excel files
- 📝 Easy data reading and writing operations
- 🎨 Comprehensive cell formatting (fonts, colors, alignment)
- 🔄 Cell merging and spanning
- 📊 Row and column management
- 📑 Multi-worksheet support
- ➗ Excel formula support
- 🎯 Conditional formatting
- 🔍 Type-safe operations
- 📏 Auto-sizing columns
- 🛡️ Input validation
- 💾 Memory-efficient operations

## Installation

```bash
npm install excel-cursor
# or
yarn add excel-cursor
# or
pnpm add excel-cursor
```

## Quick Start

```typescript
import { Workbook } from 'exceljs';
import { ExcelCursor } from 'excel-cursor';

// Initialize workbook and cursor
const workbook = new Workbook();
const cursor = new ExcelCursor(workbook);

// Navigate and input data
cursor
  .move('A1')
  .setData('Hello')
  .nextRow()
  .setData('World')
  .formatCell({
    font: { bold: true },
    alignment: { vertical: 'middle', horizontal: 'center' }
  });

// Save the workbook
await workbook.xlsx.writeFile('output.xlsx');
```

## Documentation

For detailed API documentation and examples, please check:
- [API Documentation](./API.md)
- [Changelog](./CHANGELOG.md)

## Examples

### Cell Formatting

```typescript
cursor
  .move('A1')
  .setData('Styled Cell')
  .formatCell({
    font: { 
      bold: true,
      color: '#FF0000',
      size: 14
    },
    fill: {
      type: 'pattern',
      pattern: 'solid',
      fgColor: '#FFFF00'
    },
    border: {
      top: { style: 'thin', color: '#000000' },
      bottom: { style: 'thin', color: '#000000' }
    }
  });
```

### Cell Merging

```typescript
cursor
  .move('A1')
  .setData('Merged Cells')
  .colSpan(3)  // Merge 3 columns
  .formatCell({
    alignment: { horizontal: 'center' }
  });
```

### Working with Multiple Sheets

```typescript
const cursor = new ExcelCursor(workbook, 'Sheet1');
// Work with Sheet1
cursor.move('A1').setData('Sheet 1 Data');

// Create and switch to a new sheet
const sheet2 = workbook.addWorksheet('Sheet2');
cursor.switchSheet('Sheet2');
cursor.move('A1').setData('Sheet 2 Data');
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Pham Van Nguyen - [@npv2k1](https://github.com/npv2k1)

## Support

If you encounter any issues or have questions, please create an issue in the repository.
