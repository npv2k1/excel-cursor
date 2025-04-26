import { stream } from 'exceljs';
import { ExcelCursor } from '../src/index';

// Basic data input example
async function basicExample() {
  const options = {
    filename: './result/basic-example.xlsx',
    useStyles: true,
    useSharedStrings: true,
  };
  const workbook = new stream.xlsx.WorkbookWriter(options);
  const cursor = new ExcelCursor(workbook, 'Basic Example');

  cursor
    .move('A1')
    .setData('Basic Data Input')
    .nextRow()
    .setData('Simple text')
    .nextRow()
    .setData(42)
    .nextRow()
    .setData(new Date());

  cursor.commit();
}

// Formatting example
async function formattingExample() {
  const options = {
    filename: './result/formatting-example.xlsx',
    useStyles: true,
    useSharedStrings: true,
  };
  const workbook = new stream.xlsx.WorkbookWriter(options);
  const cursor = new ExcelCursor(workbook, 'Formatting');

  cursor
    .move('A1')
    .setData('Formatting Example')
    .formatCell({
      font: { bold: true, size: 14 },
      alignment: { horizontal: 'center' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90CAF9' } }
    })
    .nextRow(2)
    .setData('Bold Text')
    .formatCell({ font: { bold: true } })
    .nextCol()
    .setData('Italic Text')
    .formatCell({ font: { italic: true } })
    .nextCol()
    .setData('Underline')
    .formatCell({ font: { underline: true } });

  cursor.commit();
}

// Formulas and comments example
async function formulasExample() {
  const options = {
    filename: './result/formulas-example.xlsx',
    useStyles: true,
    useSharedStrings: true,
  };
  const workbook = new stream.xlsx.WorkbookWriter(options);
  const cursor = new ExcelCursor(workbook, 'Formulas');

  // Set up some numbers
  cursor
    .move('A1')
    .setData('Numbers')
    .nextRow()
    .setData(10)
    .nextRow()
    .setData(20)
    .nextRow()
    .setData(30);

  // Add formulas
  cursor
    .move('B1')
    .setData('Formulas')
    .nextRow()
    .setFormula('=A2*2')
    .addComment('Doubles the value in A2')
    .nextRow()
    .setFormula('=SUM(A2:A4)')
    .addComment('Sums all numbers')
    .nextRow()
    .setFormula('=AVERAGE(A2:A4)')
    .addComment('Calculates average');

  cursor.commit();
}

// Run all examples
async function main() {
  await basicExample();
  await formattingExample();
  await formulasExample();
}

main();
