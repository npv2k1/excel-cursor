import { stream, Workbook, Cell } from 'exceljs';

export function createStreamWorkbook(options: stream.xlsx.WorkbookWriterOptions) {
  const workbook = new stream.xlsx.WorkbookWriter(options);
  return workbook;
}

export function createWorkbook() {
  const workbook = new Workbook();
  return workbook;
}

/**
 * Check if a cell contains a formula
 * @param cell - The cell to check
 * @returns true if the cell contains a formula, false otherwise
 */
export function isFormulaCell(cell: Cell): boolean {
  return cell.type === 6 || (cell.value && typeof cell.value === 'object' && 'formula' in cell.value);
}

/**
 * Get the formula string from a formula cell
 * @param cell - The cell to extract the formula from
 * @returns the formula string or null if the cell doesn't contain a formula
 */
export function getFormulaFromCell(cell: Cell): string | null {
  if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) {
    return (cell.value as any).formula;
  }
  return null;
}

/**
 * Get the calculated value from a formula cell, or the regular value for non-formula cells
 * @param cell - The cell to get the value from
 * @returns the calculated result if available, otherwise the cell's raw value
 */
export function getFormulaCellValue(cell: Cell): any {
  if (cell.value && typeof cell.value === 'object') {
    const cellValue = cell.value as any;
    // If the cell has a calculated result, return it
    if ('result' in cellValue && cellValue.result !== undefined) {
      return cellValue.result;
    }
    // If it's a formula without result, return the formula object
    if ('formula' in cellValue) {
      return cell.value;
    }
  }
  // Return the raw value for regular cells
  return cell.value;
}

/**
 * Process a formula cell and return detailed information about it
 * @param cell - The cell to process
 * @returns an object with formula information
 */
export function processFormulaCell(cell: Cell): {
  isFormula: boolean;
  formula: string | null;
  hasResult: boolean;
  result: any;
  value: any;
} {
  const isFormula = isFormulaCell(cell);
  const formula = getFormulaFromCell(cell);
  const result = getFormulaCellValue(cell);
  const hasResult = isFormula && cell.value && typeof cell.value === 'object' && 'result' in (cell.value as any) && (cell.value as any).result !== undefined;

  return {
    isFormula,
    formula,
    hasResult,
    result: hasResult ? (cell.value as any).result : null,
    value: cell.value,
  };
}
