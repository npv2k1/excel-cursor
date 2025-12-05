import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/core/ExcelCursor';

describe('ExcelCursor Formula Methods', () => {
  let workbook: Workbook;
  let cursor: ExcelCursor;

  beforeEach(() => {
    workbook = new Workbook();
    cursor = new ExcelCursor({
      workbook,
      sheetName: 'TestSheet',
    });
  });

  describe('isFormulaCell', () => {
    it('should return false for regular cells', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setData('text');
      
      expect(cursor.isFormulaCell('A1')).toBe(false);
      expect(cursor.isFormulaCell('B1')).toBe(false);
    });

    it('should return true for formula cells', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setFormula('=A1*2');
      
      expect(cursor.isFormulaCell('B1')).toBe(true);
    });

    it('should work with current position when no address provided', () => {
      cursor.move('A1').setData(10);
      expect(cursor.isFormulaCell()).toBe(false);
      
      cursor.move('B1').setFormula('=A1*2');
      expect(cursor.isFormulaCell()).toBe(true);
    });
  });

  describe('getFormula', () => {
    it('should return null for non-formula cells', () => {
      cursor.move('A1').setData(10);
      
      expect(cursor.getFormula('A1')).toBeNull();
    });

    it('should return formula string for formula cells', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setFormula('=A1*2');
      
      expect(cursor.getFormula('B1')).toBe('=A1*2');
    });

    it('should work with current position when no address provided', () => {
      cursor.move('A1').setData(10);
      expect(cursor.getFormula()).toBeNull();
      
      cursor.move('B1').setFormula('=SUM(A1:A2)');
      expect(cursor.getFormula()).toBe('=SUM(A1:A2)');
    });
  });

  describe('getFormulaCellValue', () => {
    it('should return raw value for regular cells', () => {
      cursor.move('A1').setData(42);
      cursor.move('B1').setData('hello');
      
      expect(cursor.getFormulaCellValue('A1')).toBe(42);
      expect(cursor.getFormulaCellValue('B1')).toBe('hello');
    });

    it('should return formula object for formula cells without result', () => {
      cursor.move('B1').setFormula('=A1*2');
      
      const result = cursor.getFormulaCellValue('B1');
      expect(result).toEqual({ formula: '=A1*2' });
    });

    it('should return calculated result when available', () => {
      // Manually set a formula cell with result
      const cell = cursor['getCell']('B1');
      cell.value = { formula: '=10*2', result: 20 };
      
      expect(cursor.getFormulaCellValue('B1')).toBe(20);
    });
  });

  describe('processFormulaCell', () => {
    it('should process regular cells correctly', () => {
      cursor.move('A1').setData(42);
      
      const result = cursor.processFormulaCell('A1');
      
      expect(result).toEqual({
        isFormula: false,
        formula: null,
        hasResult: false,
        result: null,
        value: 42,
      });
    });

    it('should process formula cells without result correctly', () => {
      cursor.move('B1').setFormula('=A1*2');
      
      const result = cursor.processFormulaCell('B1');
      
      expect(result).toEqual({
        isFormula: true,
        formula: '=A1*2',
        hasResult: false,
        result: null,
        value: { formula: '=A1*2' },
      });
    });

    it('should process formula cells with result correctly', () => {
      // Manually set a formula cell with result
      const cell = cursor['getCell']('B1');
      cell.value = { formula: '=10*2', result: 20 };
      
      const result = cursor.processFormulaCell('B1');
      
      expect(result).toEqual({
        isFormula: true,
        formula: '=10*2',
        hasResult: true,
        result: 20,
        value: { formula: '=10*2', result: 20 },
      });
    });

    it('should work with current position when no address provided', () => {
      cursor.move('A1').setData(42);
      
      const result = cursor.processFormulaCell();
      expect(result.isFormula).toBe(false);
      expect(result.value).toBe(42);
    });
  });

  describe('Integration with existing methods', () => {
    it('should work together with setFormula and getCellValue', () => {
      cursor
        .move('A1').setData(10)
        .move('A2').setData(20)
        .move('B1').setFormula('=A1+A2');

      expect(cursor.isFormulaCell('B1')).toBe(true);
      expect(cursor.getFormula('B1')).toBe('=A1+A2');
      
      // getCellValue returns the raw value
      expect(cursor.getCellValue('B1')).toEqual({ formula: '=A1+A2' });
      
      // getFormulaCellValue also returns the raw value when no result is calculated
      expect(cursor.getFormulaCellValue('B1')).toEqual({ formula: '=A1+A2' });
    });
  });
});