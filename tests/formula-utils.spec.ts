import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/core/ExcelCursor';
import { isFormulaCell, getFormulaFromCell, getFormulaCellValue, processFormulaCell } from '../src/utils';

describe('Formula Utils', () => {
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
      
      const cell1 = cursor['getCell']('A1');
      const cell2 = cursor['getCell']('B1');
      
      expect(isFormulaCell(cell1)).toBe(false);
      expect(isFormulaCell(cell2)).toBe(false);
    });

    it('should return true for formula cells', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setFormula('=A1*2');
      
      const cell = cursor['getCell']('B1');
      expect(isFormulaCell(cell)).toBe(true);
    });
  });

  describe('getFormulaFromCell', () => {
    it('should return null for non-formula cells', () => {
      cursor.move('A1').setData(10);
      
      const cell = cursor['getCell']('A1');
      expect(getFormulaFromCell(cell)).toBeNull();
    });

    it('should return formula string for formula cells', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setFormula('=A1*2');
      
      const cell = cursor['getCell']('B1');
      expect(getFormulaFromCell(cell)).toBe('=A1*2');
    });
  });

  describe('getFormulaCellValue', () => {
    it('should return raw value for regular cells', () => {
      cursor.move('A1').setData(42);
      cursor.move('B1').setData('hello');
      
      const cell1 = cursor['getCell']('A1');
      const cell2 = cursor['getCell']('B1');
      
      expect(getFormulaCellValue(cell1)).toBe(42);
      expect(getFormulaCellValue(cell2)).toBe('hello');
    });

    it('should return formula object for formula cells without result', () => {
      cursor.move('B1').setFormula('=A1*2');
      
      const cell = cursor['getCell']('B1');
      const result = getFormulaCellValue(cell);
      
      expect(result).toEqual({ formula: '=A1*2' });
    });

    it('should return calculated result when available', () => {
      // Simulate a formula cell with calculated result
      const cell = cursor['getCell']('B1');
      cell.value = { formula: '=10*2', result: 20 };
      
      expect(getFormulaCellValue(cell)).toBe(20);
    });
  });

  describe('processFormulaCell', () => {
    it('should process regular cells correctly', () => {
      cursor.move('A1').setData(42);
      
      const cell = cursor['getCell']('A1');
      const result = processFormulaCell(cell);
      
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
      
      const cell = cursor['getCell']('B1');
      const result = processFormulaCell(cell);
      
      expect(result).toEqual({
        isFormula: true,
        formula: '=A1*2',
        hasResult: false,
        result: null,
        value: { formula: '=A1*2' },
      });
    });

    it('should process formula cells with result correctly', () => {
      // Simulate a formula cell with calculated result
      const cell = cursor['getCell']('B1');
      cell.value = { formula: '=10*2', result: 20 };
      
      const result = processFormulaCell(cell);
      
      expect(result).toEqual({
        isFormula: true,
        formula: '=10*2',
        hasResult: true,
        result: 20,
        value: { formula: '=10*2', result: 20 },
      });
    });
  });
});