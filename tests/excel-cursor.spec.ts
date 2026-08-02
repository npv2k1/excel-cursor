import { stream, Workbook } from 'exceljs';

import { ExcelCursor } from '../src/core/ExcelCursor';

describe('ExcelCursor', () => {
  describe('Constructor', () => {
    it('should create a cursor with default options', () => {
      const cursor = new ExcelCursor();
      expect(cursor).toBeInstanceOf(ExcelCursor);
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 1 });
    });

    it('should create a cursor with custom sheet name', () => {
      const cursor = new ExcelCursor({ sheetName: 'CustomSheet' });
      const workbook = cursor.getWorkbook();
      const worksheet = workbook.getWorksheet('CustomSheet');
      expect(worksheet).toBeDefined();
    });

    it('should support the legacy workbook plus options signature when workbook is omitted', () => {
      const cursor = new ExcelCursor(undefined, { sheetName: 'LegacySheet' });
      expect(cursor.getWorkbook().getWorksheet('LegacySheet')).toBeDefined();
    });

    it('should create a cursor with existing workbook', () => {
      const workbook = new Workbook();
      const cursor = new ExcelCursor({ workbook });
      expect(cursor.getWorkbook()).toBe(workbook);
    });

    it('should use existing worksheet if sheetName provided', () => {
      const workbook = new Workbook();
      workbook.addWorksheet('ExistingSheet');
      const cursor = new ExcelCursor({ workbook, sheetName: 'ExistingSheet' });
      expect(cursor.getWorkbook()).toBe(workbook);
    });

    it('should reuse an existing default Sheet1', () => {
      const workbook = new Workbook();
      workbook.addWorksheet('Sheet1').getCell('C4').value = 'existing';
      const cursor = new ExcelCursor(workbook);
      expect(cursor.getCellValue('C4')).toBe('existing');
      expect(cursor.getLastRow()).toBe(4);
      expect(cursor.getLastCol()).toBe(3);
    });

    it('should create stream workbook when isStream is true', () => {
      const cursor = new ExcelCursor({ isStream: true });
      const workbook = cursor.getWorkbook();
      expect(workbook).toBeInstanceOf(stream.xlsx.WorkbookWriter);
    });
  });

  describe('Navigation', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should move to address correctly', () => {
      cursor.move('B5');
      expect(cursor.getCurrentPosition()).toEqual({ row: 5, col: 2 });
      expect(cursor.getCurrentAddress()).toBe('B5');
    });

    it('should move to position correctly', () => {
      cursor.moveTo(3, 4);
      expect(cursor.getCurrentPosition()).toEqual({ row: 3, col: 4 });
    });

    it('should move to next row', () => {
      cursor.nextRow();
      expect(cursor.getCurrentPosition()).toEqual({ row: 2, col: 1 });
      cursor.nextRow(3);
      expect(cursor.getCurrentPosition()).toEqual({ row: 5, col: 1 });
    });

    it('should move to previous row but not below 1', () => {
      cursor.prevRow();
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 1 });
      cursor.moveTo(5, 1).prevRow(2);
      expect(cursor.getCurrentPosition()).toEqual({ row: 3, col: 1 });
    });

    it('should move to next column', () => {
      cursor.nextCol();
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 2 });
      cursor.nextCol(3);
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 5 });
    });

    it('should move to previous column but not below 1', () => {
      cursor.prevCol();
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 1 });
      cursor.moveTo(1, 5).prevCol(2);
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 3 });
    });

    it('should go back to first column', () => {
      cursor.moveTo(5, 10).goBackToFirstCollumn();
      expect(cursor.getCurrentPosition()).toEqual({ row: 5, col: 1 });
    });

    it('should reject invalid or out-of-bounds navigation', () => {
      expect(() => cursor.moveTo(0, 1)).toThrow();
      expect(() => cursor.move('XFE1')).toThrow();
      expect(() => cursor.nextRow(-1)).toThrow('Invalid navigation distance');
      cursor.move('XFD1');
      expect(() => cursor.nextCol()).toThrow();
    });

    it('should move to last row', () => {
      cursor.setData('test').setData('test2', 'A5');
      cursor.moveLastRow();
      expect(cursor.getCurrentPosition().row).toBe(5);
    });

    it('should move to last column', () => {
      cursor.setData('test').setData('test2', 'D1');
      cursor.moveLastCol();
      expect(cursor.getCurrentPosition().col).toBe(4);
    });
  });

  describe('Data Operations', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should set data at current position', () => {
      cursor.setData('Hello');
      expect(cursor.getCellValue()).toBe('Hello');
    });

    it('should set data at specific address', () => {
      cursor.setData('World', 'C3');
      expect(cursor.getCellValue('C3')).toBe('World');
    });

    it('should get cell value', () => {
      cursor.setData(42);
      expect(cursor.getCellValue()).toBe(42);
    });

    it('should set formula', () => {
      cursor.moveTo(1, 1).setData(10).moveTo(1, 2).setData(20).moveTo(1, 3).setFormula('=A1+B1');
      const cell = cursor.getCellValue('C1');
      expect(cell).toEqual({ formula: 'A1+B1' });
    });

    it('should add row', () => {
      cursor.addRow(['A', 'B', 'C']);
      expect(cursor.getCellValue('A1')).toBe('A');
      expect(cursor.getCellValue('B1')).toBe('B');
      expect(cursor.getCellValue('C1')).toBe('C');
    });

    it('should add multiple rows', () => {
      cursor.addRows([
        ['A1', 'B1'],
        ['A2', 'B2'],
      ]);
      expect(cursor.getCellValue('A1')).toBe('A1');
      expect(cursor.getCellValue('A2')).toBe('A2');
    });

    it('should insert row', () => {
      cursor.addRow(['Row1']);
      cursor.addRow(['Row3']);
      cursor.moveTo(2, 1).insertRow(['Row2']);
      expect(cursor.getCellValue('A2')).toBe('Row2');
    });

    it('should delete row', () => {
      cursor.addRow(['Row1']);
      cursor.addRow(['Row2']);
      cursor.moveTo(2, 1).deleteRow();
      expect(cursor.getCellValue('A2')).toBeNull();
    });

    it('should update lastRow and lastCol when setting data', () => {
      cursor.setData('test', 'E5');
      expect(cursor.getLastRow()).toBe(5);
      expect(cursor.getLastCol()).toBe(5);
    });
  });

  describe('Cell Formatting', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should format cell', () => {
      cursor.setData('Bold').formatCell({ font: { bold: true } });
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.style.font?.bold).toBe(true);
    });

    it('should format cell number', () => {
      cursor.setData(1000).formatCellNumber(undefined, '#,##0.00');
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.numFmt).toBe('#,##0.00');
    });

    it('should apply border to cell', () => {
      cursor.setData('Bordered').borderAll();
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.border).toEqual({
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      });
    });

    it('should center cell content', () => {
      cursor.setData('Centered').center();
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.alignment).toEqual({
        vertical: 'middle',
        horizontal: 'center',
      });
    });

    it('should apply style to range', () => {
      cursor.setData('A1', 'A1').setData('B1', 'B1').setData('A2', 'A2').setData('B2', 'B2');
      cursor.applyStyleToRange({ font: { bold: true } }, 'A1', 'B2');
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getCell('A1').style.font?.bold).toBe(true);
      expect(worksheet.getCell('B2').style.font?.bold).toBe(true);
    });
  });

  describe('Merging Cells', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should merge columns (colSpan)', () => {
      cursor.moveTo(1, 1).colSpan(3);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(() => worksheet.getCell('A1')).not.toThrow();
    });

    it('should merge rows (rowSpan)', () => {
      cursor.moveTo(1, 1).rowSpan(3);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(() => worksheet.getCell('A1')).not.toThrow();
    });

    it('should propagate invalid row span errors', () => {
      expect(() => cursor.rowSpan(0)).toThrow('Invalid row span: 0');
    });
  });

  describe('Column and Row Operations', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should set column width', () => {
      cursor.setColWidth(20);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getColumn(1).width).toBe(20);
    });

    it('should set column width by address', () => {
      cursor.setColWidth(15, 'B1');
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getColumn(2).width).toBe(15);
    });

    it('should set column width by number', () => {
      cursor.setColWidth(25, 3);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getColumn(3).width).toBe(25);
    });

    it('should set row height', () => {
      cursor.setRowHeight(30);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getRow(1).height).toBe(30);
    });

    it('should set row height by address', () => {
      cursor.setRowHeight(25, 'A5');
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getRow(5).height).toBe(25);
    });

    it('should set row height by number', () => {
      cursor.setRowHeight(35, 3);
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet.getRow(3).height).toBe(35);
    });
  });

  describe('Comments', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should add comment to cell', () => {
      cursor.setData('Value').addComment('This is a comment');
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.note).toBeDefined();
      expect(cell.note?.texts[0].text).toBe('This is a comment');
    });

    it('should add comment with author', () => {
      cursor.setData('Value').addComment('Comment text', 'Author');
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.note?.author).toBe('Author');
    });
  });

  describe('Sheet Operations', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should create new sheet', () => {
      cursor.createSheet('NewSheet');
      const worksheet = cursor.getWorkbook().getWorksheet('NewSheet');
      expect(worksheet).toBeDefined();
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 1 });
    });

    it('should switch to existing sheet', () => {
      cursor.createSheet('Sheet2');
      cursor.switchSheet('Sheet1');
      expect(cursor.getCurrentPosition()).toEqual({ row: 1, col: 1 });
    });

    it('should throw error when switching to non-existent sheet', () => {
      expect(() => cursor.switchSheet('NonExistent')).toThrow('Sheet NonExistent not found');
    });

    it('should set worksheet', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TestSheet');
      const newCursor = new ExcelCursor({ workbook });
      newCursor.setWorksheet(worksheet);
      expect(newCursor.getWorkbook()).toBe(workbook);
    });
  });

  describe('Tracking', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should track last row correctly', () => {
      cursor.addRow(['Row1']);
      cursor.addRow(['Row2']);
      cursor.addRow(['Row3']);
      expect(cursor.getLastRow()).toBe(3);
    });

    it('should track last column correctly', () => {
      cursor.addRow(['A', 'B', 'C', 'D']);
      expect(cursor.getLastCol()).toBe(4);
    });

    it('should get last column address', () => {
      cursor.addRow(['A', 'B', 'C']);
      expect(cursor.getLastColAddress()).toBe('C');
    });

    it('should get last cell address', () => {
      cursor.addRow(['A', 'B', 'C']);
      expect(cursor.getLastCellAddress()).toBe('C1');
    });
  });

  describe('Region Operations', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should create region from current position', () => {
      cursor.moveTo(1, 1);
      const region = cursor.createRegion(3, 2);
      expect(region).toBe('A1:B3');
    });

    it('should create region from different position', () => {
      cursor.moveTo(2, 3);
      const region = cursor.createRegion(2, 3);
      expect(region).toBe('C2:E3');
    });
  });

  describe('Copy Range', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should copy range to target', () => {
      cursor.setData('A1', 'A1');
      cursor.setData('B1', 'B1');
      cursor.setData('A2', 'A2');
      cursor.setData('B2', 'B2');

      cursor.copyRange('A1', 'B2', 'D4');

      expect(cursor.getCellValue('D4')).toBe('A1');
      expect(cursor.getCellValue('E4')).toBe('B1');
      expect(cursor.getCellValue('D5')).toBe('A2');
      expect(cursor.getCellValue('E5')).toBe('B2');
    });

    it('should copy overlapping ranges from a stable snapshot', () => {
      cursor.setData('one', 'A1').setData('two', 'B1');
      cursor.copyRange('A1', 'B1', 'B1');
      expect(cursor.getCellValue('B1')).toBe('one');
      expect(cursor.getCellValue('C1')).toBe('two');
    });

    it('should reject reversed and out-of-bounds ranges', () => {
      expect(() => cursor.copyRange('B2', 'A1', 'C3')).toThrow('Range end');
      expect(() => cursor.copyRange('XFD1', 'XFD1', 'XFD2')).not.toThrow();
      expect(() => cursor.copyRange('XFD1', 'XFD1', 'XFE1')).toThrow();
    });
  });

  describe('Conditional Formatting', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should add conditional formatting', () => {
      cursor.addConditionalFormatting('A1:A10', 'cellIs', { operator: 'greaterThan', values: [5] });
      const worksheet = cursor.getWorkbook().getWorksheet('Sheet1');
      expect(worksheet).toBeDefined();
    });
  });

  describe('isBorderAll option', () => {
    it('should apply border to all cells when isBorderAll is true', () => {
      const cursor = new ExcelCursor({ isBorderAll: true });
      cursor.setData('Test');
      const cell = cursor.getWorkbook().getWorksheet('Sheet1').getCell('A1');
      expect(cell.border).toEqual({
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      });
    });
  });

  describe('Address Parsing Errors', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should throw error for invalid address', () => {
      expect(() => cursor.move('InvalidAddress')).toThrow('Invalid cell address: InvalidAddress');
    });
  });

  describe('Column Letter Conversion', () => {
    let cursor: ExcelCursor;

    beforeEach(() => {
      cursor = new ExcelCursor();
    });

    it('should handle single letter columns', () => {
      cursor.move('A1');
      expect(cursor.getCurrentPosition().col).toBe(1);
      cursor.move('Z1');
      expect(cursor.getCurrentPosition().col).toBe(26);
    });

    it('should handle double letter columns', () => {
      cursor.move('AA1');
      expect(cursor.getCurrentPosition().col).toBe(27);
      cursor.move('AZ1');
      expect(cursor.getCurrentPosition().col).toBe(52);
      cursor.move('BA1');
      expect(cursor.getCurrentPosition().col).toBe(53);
    });

    it('should handle triple letter columns', () => {
      cursor.move('AAA1');
      expect(cursor.getCurrentPosition().col).toBe(703);
    });
  });
});
