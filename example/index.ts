import { mkdir } from 'node:fs/promises';
import { Workbook } from 'exceljs';

import { ExcelCursor, StreamingExcelWriter } from '../src';

async function createFormattedWorkbook() {
  const workbook = new Workbook();
  const cursor = new ExcelCursor(workbook, { sheetName: 'Report', maxCells: 10_000 });

  cursor
    .move('A1')
    .setData('Revenue report')
    .colSpan(2)
    .formatCell({ font: { bold: true, size: 14 }, alignment: { horizontal: 'center' } })
    .move('A2')
    .setData(10)
    .nextRow()
    .setData(20)
    .nextRow()
    .setData(30)
    .move('B2')
    .setTrustedFormula('A2*2')
    .nextRow()
    .setTrustedFormula('=A3*2')
    .nextRow()
    .setTrustedFormula('SUM(A2:A4)')
    .move('C2')
    .setSafeText('=untrusted imported value');

  await cursor.saveWorkbook('./result/report.xlsx');
}

async function createStreamingWorkbook() {
  const writer = new StreamingExcelWriter({
    filename: './result/streaming-report.xlsx',
    sheetName: 'Rows',
    maxRows: 100_000,
  });
  await writer.addRows([
    ['id', 'name'],
    [1, 'Alpha'],
    [2, 'Beta'],
  ]);
  await writer.commit();
}

async function main() {
  await mkdir('./result', { recursive: true });
  await createFormattedWorkbook();
  await createStreamingWorkbook();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
