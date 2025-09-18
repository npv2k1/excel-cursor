import { Workbook } from 'exceljs';
import { ExcelCursor } from '../src/core/ExcelCursor';

describe('Formula Calculation Engine', () => {
  let workbook: Workbook;
  let cursor: ExcelCursor;

  beforeEach(() => {
    workbook = new Workbook();
    cursor = new ExcelCursor({
      workbook,
      sheetName: 'TestSheet',
    });
  });

  describe('calculateFormula', () => {
    it('should calculate basic arithmetic formulas', () => {
      const result1 = cursor.calculateFormula('10 + 5');
      expect(result1.error).toBeNull();
      expect(result1.result).toBe(15);

      const result2 = cursor.calculateFormula('20 * 3');
      expect(result2.error).toBeNull();
      expect(result2.result).toBe(60);

      const result3 = cursor.calculateFormula('100 / 4');
      expect(result3.error).toBeNull();
      expect(result3.result).toBe(25);
    });

    it('should calculate formulas with cell references', () => {
      // Set up some data
      cursor.move('A1').setData(100);
      cursor.move('A2').setData(200);

      const result1 = cursor.calculateFormula('A1 + A2');
      expect(result1.error).toBeNull();
      expect(result1.result).toBe(300);

      const result2 = cursor.calculateFormula('A1 * 2');
      expect(result1.error).toBeNull();
      expect(result2.result).toBe(200);
    });

    it('should calculate SUM formulas', () => {
      // Set up some data
      cursor.move('A1').setData(10);
      cursor.move('A2').setData(20);
      cursor.move('A3').setData(30);

      const result1 = cursor.calculateFormula('SUM(A1,A2,A3)');
      expect(result1.error).toBeNull();
      expect(result1.result).toBe(60);

      const result2 = cursor.calculateFormula('SUM(A1:A3)');
      expect(result2.error).toBeNull();
      expect(result2.result).toBe(60);
    });

    it('should calculate AVERAGE formulas', () => {
      cursor.move('B1').setData(100);
      cursor.move('B2').setData(200);
      cursor.move('B3').setData(300);

      const result = cursor.calculateFormula('AVERAGE(B1,B2,B3)');
      expect(result.error).toBeNull();
      expect(result.result).toBe(200);
    });

    it('should calculate MAX and MIN formulas', () => {
      cursor.move('C1').setData(5);
      cursor.move('C2').setData(15);
      cursor.move('C3').setData(10);

      const maxResult = cursor.calculateFormula('MAX(C1,C2,C3)');
      expect(maxResult.error).toBeNull();
      expect(maxResult.result).toBe(15);

      const minResult = cursor.calculateFormula('MIN(C1,C2,C3)');
      expect(minResult.error).toBeNull();
      expect(minResult.result).toBe(5);
    });

    it('should handle errors gracefully', () => {
      const result = cursor.calculateFormula('INVALID_FUNCTION(1,2,3)');
      expect(result.error).not.toBeNull();
      expect(result.result).toBeNull();
    });
  });

  describe('getCalculatedValue', () => {
    it('should return calculated values for formula cells', () => {
      // Set up data
      cursor.move('A1').setData(50);
      cursor.move('A2').setData(100);

      // Set formula
      cursor.move('B1').setFormula('=A1+A2');

      // Get calculated value
      const value = cursor.getCalculatedValue('B1');
      expect(value).toBe(150);
    });

    it('should return raw values for non-formula cells', () => {
      cursor.move('A1').setData(42);
      
      const value = cursor.getCalculatedValue('A1');
      expect(value).toBe(42);
    });

    it('should cache calculated results', () => {
      cursor.move('A1').setData(10);
      cursor.move('B1').setFormula('=A1*5');

      // First call should calculate
      const value1 = cursor.getCalculatedValue('B1');
      expect(value1).toBe(50);

      // Check that result is cached in cell
      const cellInfo = cursor.processFormulaCell('B1');
      expect(cellInfo.hasResult).toBe(true);
      expect(cellInfo.result).toBe(50);

      // Second call should use cached value
      const value2 = cursor.getCalculatedValue('B1');
      expect(value2).toBe(50);
    });
  });

  describe('calculateAndUpdateFormulaCell', () => {
    it('should update formula cells with calculated results', () => {
      cursor.move('A1').setData(25);
      cursor.move('B1').setFormula('=A1*4');

      // Initially no result
      expect(cursor.processFormulaCell('B1').hasResult).toBe(false);

      // Calculate and update
      cursor.calculateAndUpdateFormulaCell('B1');

      // Now should have result
      const info = cursor.processFormulaCell('B1');
      expect(info.hasResult).toBe(true);
      expect(info.result).toBe(100);
    });
  });

  describe('Integration with existing methods', () => {
    it('should work with existing formula methods', () => {
      cursor.move('A1').setData(10);
      cursor.move('A2').setData(20);
      cursor.move('B1').setFormula('=SUM(A1,A2)');

      expect(cursor.isFormulaCell('B1')).toBe(true);
      expect(cursor.getFormula('B1')).toBe('=SUM(A1,A2)');

      // Calculate the formula
      const calculated = cursor.getCalculatedValue('B1');
      expect(calculated).toBe(30);

      // Check that getFormulaCellValue now returns the calculated result
      expect(cursor.getFormulaCellValue('B1')).toBe(30);
    });
  });
});