import { stream } from 'exceljs';

import { ExcelCursor } from '../src/index';

async function main() {
  // construct a streaming XLSX workbook writer with styles and shared strings
  const options = {
    filename: './result/streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true,
  };
  const workbook = new stream.xlsx.WorkbookWriter(options);

  const cursor = new ExcelCursor(workbook, 'Sheet1');

  cursor.move('A1').setData('Hello, World!');

  cursor.move('B2').setData('ExcelJS Cursor Example');

  cursor.move('C3').setData('This is a test.');

  for (let i = 0; i < 1000; i++) {
    cursor
      .move('A' + (i + 4))
      .setData('Row ' + (i + 1))
      .addComment('This is a comment')
      .nextCol(1)
      // .setData('Column ' + (i + 1))
      .setFormula('=1+2');
  }

  cursor.commit();
}

main();
