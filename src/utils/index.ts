import { stream, Workbook } from 'exceljs';

export function createStreamWorkbook(options: stream.xlsx.WorkbookWriterOptions) {
  const workbook = new stream.xlsx.WorkbookWriter(options);
  return workbook;
}

export function createWorkbook() {
  const workbook = new Workbook();
  return workbook;
}
