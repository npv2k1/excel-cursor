/**
 * Convert column letter (A, B, C...) to number (1, 2, 3...)
 */
export function colLetterToNumber(colLetter: string): number {
  let result = 0;
  for (let i = 0; i < colLetter.length; i++) {
    result = result * 26 + (colLetter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Convert column number (1, 2, 3...) to letter (A, B, C...)
 */
export function colNumberToLetter(colNumber: number): string {
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
  const match = address.match(/([A-Z]+)(\d+)/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  const colLetter = match[1];
  const rowNumber = parseInt(match[2], 10);

  return {
    row: rowNumber,
    col: colLetterToNumber(colLetter),
  };
}

/**
 * Convert row and column position to cell address (A1, B2...)
 */
export function positionToAddress(row: number, col: number): string {
  return `${colNumberToLetter(col)}${row}`;
}
