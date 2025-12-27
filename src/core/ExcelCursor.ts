import { stream, Style, Workbook } from 'exceljs';
import { isNil, isString, merge } from 'lodash';

import * as os from 'os';
import * as path from 'path';
import { Cell, CellPosition, Worksheet } from '../types';
export type ExcelCursorOptions = {
  filename?: string;
  workbook?: stream.xlsx.WorkbookWriter | Workbook;
  sheetName?: string;
  isStream?: boolean;
  isBorderAll?: boolean;
};

export class ExcelCursor {
  private workbook: stream.xlsx.WorkbookWriter | Workbook;
  private worksheet: Worksheet;
  private position: CellPosition = { row: 1, col: 1 };
  private lastRow = 1;
  private lastCol = 1;
  private options: ExcelCursorOptions = {};

  constructor(options?: ExcelCursorOptions) {
    const { workbook, sheetName, filename, isStream } = options ?? {};
    this.workbook = new Workbook();
    this.options = options || {};

    if (workbook) {
      this.workbook = workbook;
    } else if (isStream) {
      this.workbook = new stream.xlsx.WorkbookWriter({
        filename: filename || path.join(os.tmpdir(), `excel-cursor-stream-${Date.now()}.xlsx`),
        useStyles: true,
        useSharedStrings: true,
      });
    }

    // Reset position and tracking
    this.position = { row: 1, col: 1 };
    this.lastRow = 1;
    this.lastCol = 1;

    if (sheetName && this.workbook.getWorksheet(sheetName)) {
      this.worksheet = this.workbook.getWorksheet(sheetName);
    } else if (sheetName) {
      this.worksheet = this.workbook.addWorksheet(sheetName);
    } else {
      this.worksheet = this.workbook.addWorksheet('Sheet1');
    }
  }

  getWorkbook(): any {
    return this.workbook;
  }

  setWorksheet(worksheet: any): ExcelCursor {
    this.worksheet = worksheet;
    return this;
  }

  // Chuyển đổi địa chỉ cột dạng chữ (A, B, C...) sang số (1, 2, 3...)
  private colLetterToNumber(colLetter: string): number {
    let result = 0;
    for (let i = 0; i < colLetter.length; i++) {
      result = result * 26 + (colLetter.charCodeAt(i) - 64);
    }
    return result;
  }

  // Chuyển đổi số cột (1, 2, 3...) sang dạng chữ (A, B, C...)
  private colNumberToLetter(colNumber: number): string {
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

  // Phân tích địa chỉ ô (A1, B2...) thành vị trí hàng và cột
  private parseAddress(address: string): CellPosition {
    const match = address.match(/([A-Z]+)(\d+)/);
    if (!match) {
      throw new Error(`Invalid cell address: ${address}`);
    }

    const colLetter = match[1];
    const rowNumber = parseInt(match[2], 10);

    return {
      row: rowNumber,
      col: this.colLetterToNumber(colLetter),
    };
  }

  // Chuyển đổi vị trí hàng và cột thành địa chỉ ô (A1, B2...)
  private positionToAddress(position: CellPosition): string {
    return `${this.colNumberToLetter(position.col)}${position.row}`;
  }

  // Lấy ô từ vị trí hoặc địa chỉ
  private getCell(positionOrAddress?: CellPosition | string): Cell {
    let position: CellPosition;

    if (isString(positionOrAddress)) {
      position = this.parseAddress(positionOrAddress as string);
    } else if (positionOrAddress) {
      position = positionOrAddress as CellPosition;
    } else {
      position = this.position;
    }

    const row = this.worksheet.getRow(position.row);
    return row.getCell(position.col);
  }

  // Di chuyển đến ô có địa chỉ nào đó
  move(address: string): ExcelCursor {
    this.position = this.parseAddress(address);
    return this;
  }

  // Di chuyển đến vị trí (row, col) cụ thể
  moveTo(row: number, col: number): ExcelCursor {
    this.position = { row, col };
    return this;
  }

  // Gán dữ liệu cho ô hiện tại hoặc ô có địa chỉ bất kỳ
  setData(data: any, address?: string): ExcelCursor {
    let position: CellPosition;

    if (address) {
      position = this.parseAddress(address);
    } else {
      position = { ...this.position };
    }

    const cell = this.getCell(position);
    cell.value = data;

    // Update tracking when setting data
    this.updateLastPosition(position);

    return this;
  }

  // Di chuyển xuống n hàng
  nextRow(n = 1): ExcelCursor {
    this.position.row += n;
    return this;
  }

  // Di chuyển lên n hàng
  prevRow(n = 1): ExcelCursor {
    this.position.row = Math.max(1, this.position.row - n);
    return this;
  }

  // Di chuyển sang phải n cột
  nextCol(n = 1): ExcelCursor {
    this.position.col += n;
    return this;
  }

  // Di chuyển sang trái n cột
  prevCol(n = 1): ExcelCursor {
    this.position.col = Math.max(1, this.position.col - n);
    return this;
  }

  // Span n cột từ vị trí hiện tại hoặc địa chỉ bất kỳ
  colSpan(n: number, address?: string): ExcelCursor {
    const startPos = address ? this.parseAddress(address) : this.position;
    const endCol = startPos.col + n - 1;

    this.worksheet.mergeCells(startPos.row, startPos.col, startPos.row, endCol);

    return this;
  }

  // Span n hàng từ vị trí hiện tại hoặc địa chỉ bất kỳ
  rowSpan(n: number, address?: string): ExcelCursor {
    try {
      const startPos = address ? this.parseAddress(address) : this.position;
      const endRow = startPos.row + n - 1;

      this.worksheet.mergeCells(startPos.row, startPos.col, endRow, startPos.col);

      return this;
    } catch (error) {
      console.error('Error in rowSpan:', address ?? this.getCurrentAddress());
      return this;
    }
  }

  // Format ô hiện tại hoặc ô có địa chỉ bất kỳ
  formatCell(format: Partial<Style>, address?: string): ExcelCursor {
    const cell = this.getCell(address);

    if (!cell.style) {
      cell.style = {};
    }

    merge(cell.style, format);
    return this;
  }

  // Lấy địa chỉ ô hiện tại
  getCurrentAddress(): string {
    return this.positionToAddress(this.position);
  }

  // Lấy vị trí ô hiện tại
  getCurrentPosition(): CellPosition {
    return { ...this.position };
  }

  // Đặt chiều rộng cột
  setColWidth(width: number, colOrAddress?: number | string): ExcelCursor {
    let col: number | string;

    if (isString(colOrAddress)) {
      col = this.parseAddress(`${colOrAddress}`).col;
    } else if (!isNil(colOrAddress)) {
      col = colOrAddress;
    } else {
      col = this.position.col;
    }

    this.worksheet.getColumn(col).width = width;
    return this;
  }

  // Đặt chiều cao hàng
  setRowHeight(height: number, rowOrAddress?: number | string): ExcelCursor {
    let row: number;

    if (isString(rowOrAddress)) {
      row = this.parseAddress(`${rowOrAddress}`).row;
    } else if (!isNil(rowOrAddress)) {
      row = rowOrAddress as number;
    } else {
      row = this.position.row;
    }

    this.worksheet.getRow(row).height = height;
    return this;
  }

  // Chèn hàng tại vị trí hiện tại
  insertRow(values?: any[]): ExcelCursor {
    this.worksheet.insertRow(this.position.row, values || []);

    // Update tracking for inserted row
    if (values && values.length > 0) {
      this.lastRow = Math.max(this.lastRow, this.position.row);
      this.lastCol = Math.max(this.lastCol, this.position.col + values.length - 1);
    }

    return this;
  }

  // Xóa hàng tại vị trí hiện tại
  deleteRow(): ExcelCursor {
    this.worksheet.spliceRows(this.position.row, 1);
    return this;
  }

  // Thêm công thức cho ô
  setFormula(formula: string, address?: string): ExcelCursor {
    const cell = this.getCell(address);
    cell.value = { formula };
    return this;
  }

  // Thêm comment cho ô
  addComment(text: string, author?: string, address?: string): ExcelCursor {
    const cell = this.getCell(address);
    cell.note = {
      texts: [{ text, font: { name: 'Calibri', size: 11 } }],
      margins: { insetmode: 'auto' },
      editAs: 'twoCells',
      ...(author ? { author } : {}),
    };
    return this;
  }

  // Áp dụng định dạng có điều kiện
  addConditionalFormatting(
    range: string,
    type: 'cellIs' | 'containsText' | 'colorScale',
    rules: any
  ): ExcelCursor {
    this.worksheet.addConditionalFormatting({
      ref: range,
      rules: [
        {
          type,
          ...rules,
        },
      ],
    });
    return this;
  }

  // Thay đổi sheet hiện tại
  switchSheet(sheetName: string): ExcelCursor {
    const sheet = this.workbook.getWorksheet(sheetName);
    if (sheet) {
      this.worksheet = sheet;
      this.position = { row: 1, col: 1 };

      // Reset tracking for new sheet
      this.lastRow = 1;
      this.lastCol = 1;
    } else {
      throw new Error(`Sheet ${sheetName} not found`);
    }
    return this;
  }

  // Tạo sheet mới
  createSheet(sheetName: string): ExcelCursor {
    this.worksheet = this.workbook.addWorksheet(sheetName);
    this.position = { row: 1, col: 1 };

    // Reset tracking for new sheet
    this.lastRow = 1;
    this.lastCol = 1;

    return this;
  }

  // Lưu workbook
  async saveWorkbook(filepath: string): Promise<void> {
    await this.workbook.xlsx.writeFile(filepath);
  }

  async commit(): Promise<void> {
    if (this.workbook instanceof stream.xlsx.WorkbookWriter) {
      await this.workbook.commit();
    }
  }

  // Method hỗ trợ tạo vùng từ vị trí hiện tại với n hàng và m cột
  createRegion(rows: number, cols: number): string {
    const startAddress = this.getCurrentAddress();
    const endRow = this.position.row + rows - 1;
    const endCol = this.position.col + cols - 1;
    const endAddress = this.positionToAddress({ row: endRow, col: endCol });

    return `${startAddress}:${endAddress}`;
  }

  // Lấy giá trị của ô
  getCellValue(address?: string): any {
    const cell = this.getCell(address);
    return cell.value;
  }

  // Áp dụng style cho vùng
  applyStyleToRange(format: Partial<Style>, startAddress: string, endAddress: string): ExcelCursor {
    const startPos = this.parseAddress(startAddress);
    const endPos = this.parseAddress(endAddress);

    for (let row = startPos.row; row <= endPos.row; row++) {
      for (let col = startPos.col; col <= endPos.col; col++) {
        this.formatCell(format, this.positionToAddress({ row, col }));
      }
    }

    return this;
  }

  // goBack to first collumn
  goBackToFirstCollumn(): ExcelCursor {
    this.position.col = 1;
    return this;
  }

  // Sao chép dữ liệu từ vùng này sang vùng khác
  copyRange(
    sourceStartAddress: string,
    sourceEndAddress: string,
    targetStartAddress: string
  ): ExcelCursor {
    const sourceStartPos = this.parseAddress(sourceStartAddress);
    const sourceEndPos = this.parseAddress(sourceEndAddress);
    const targetStartPos = this.parseAddress(targetStartAddress);

    const rowOffset = targetStartPos.row - sourceStartPos.row;
    const colOffset = targetStartPos.col - sourceStartPos.col;

    // Calculate target end position
    const targetEndRow = sourceEndPos.row + rowOffset;
    const targetEndCol = sourceEndPos.col + colOffset;

    // Update tracking
    this.updateLastPosition({ row: targetEndRow, col: targetEndCol });

    for (let row = sourceStartPos.row; row <= sourceEndPos.row; row++) {
      for (let col = sourceStartPos.col; col <= sourceEndPos.col; col++) {
        const sourceCell = this.getCell({ row, col });
        const targetCell = this.getCell({
          row: row + rowOffset,
          col: col + colOffset,
        });

        targetCell.value = sourceCell.value;
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style || {}));
      }
    }

    return this;
  }

  // Update tracking of last row and column
  private updateLastPosition(position: CellPosition): void {
    this.lastRow = Math.max(this.lastRow, position.row);
    this.lastCol = Math.max(this.lastCol, position.col);
    if (this.options.isBorderAll) {
      this.borderAll(this.positionToAddress(position));
    }
  }

  // Get the last row that has data
  getLastRow(): number {
    return this.lastRow;
  }

  // Get the last column that has data
  getLastCol(): number {
    return this.lastCol;
  }

  // Get the last column address (like 'A', 'B', 'AA', etc.)
  getLastColAddress(): string {
    return this.colNumberToLetter(this.lastCol);
  }

  // Get the address of the last cell with data (like 'A1', 'B2', etc.)
  getLastCellAddress(): string {
    return this.positionToAddress({ row: this.lastRow, col: this.lastCol });
  }

  moveLastRow(): ExcelCursor {
    this.position.row = this.lastRow;
    return this;
  }

  moveLastCol(): ExcelCursor {
    this.position.col = this.lastCol;
    return this;
  }

  // Add row to worksheet and return the row for further manipulation
  addRow(values: any[]): any {
    const row = this.worksheet.addRow(values);

    // Update tracking
    if (values && values.length > 0) {
      this.lastRow = Math.max(this.lastRow, row.number);
      this.lastCol = Math.max(this.lastCol, values.length);
    }

    // Commit the row if using stream writer
    if (this.workbook instanceof stream.xlsx.WorkbookWriter) {
      row.commit();
    }

    return row;
  }

  // Add multiple rows at once
  addRows(data: any[][]): any[] {
    const rows = [];
    data.forEach((rowData) => {
      rows.push(this.addRow(rowData));
    });
    return rows;
  }

  formatCellNumber(address?: string, format = '#,##0.00'): ExcelCursor {
    const cell = this.getCell(address);
    cell.numFmt = format;
    return this;
  }

  borderAll(address?: string): ExcelCursor {
    const cell = this.getCell(address);
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    return this;
  }

  center(address?: string): ExcelCursor {
    const cell = this.getCell(address);
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    return this;
  }
}
