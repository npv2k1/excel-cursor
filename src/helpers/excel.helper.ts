export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLUMNS = 16_384;

function assertIntegerInRange(value: number, min: number, max: number, label: string): void {
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new Error(`Invalid Excel ${label}: ${value}`);
  }
}

/**
 * Convert column letter (A, B, C...) to number (1, 2, 3...)
 */
export function colLetterToNumber(colLetter: string): number {
  if (!/^[A-Za-z]+$/.test(colLetter)) {
    throw new Error(`Invalid Excel column: ${colLetter}`);
  }

  let result = 0;
  const normalized = colLetter.toUpperCase();
  for (let i = 0; i < normalized.length; i++) {
    result = result * 26 + (normalized.charCodeAt(i) - 64);
  }
  assertIntegerInRange(result, 1, EXCEL_MAX_COLUMNS, 'column');
  return result;
}

/**
 * Convert column number (1, 2, 3...) to letter (A, B, C...)
 */
export function colNumberToLetter(colNumber: number): string {
  assertIntegerInRange(colNumber, 1, EXCEL_MAX_COLUMNS, 'column');
  let dividend = colNumber;
  let columnName = '';
  let modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

/**
 * Parse cell address (A1, B2...) to row and column position
 */
export function parseAddress(address: string): { row: number; col: number } {
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(address);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  const colLetter = match[1];
  const rowNumber = parseInt(match[2], 10);

  const col = colLetterToNumber(colLetter);
  if (rowNumber > EXCEL_MAX_ROWS) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return { row: rowNumber, col };
}

/**
 * Convert row and column position to cell address (A1, B2...)
 */
export function positionToAddress(row: number, col: number): string {
  assertIntegerInRange(row, 1, EXCEL_MAX_ROWS, 'row');
  return `${colNumberToLetter(col)}${row}`;
}
